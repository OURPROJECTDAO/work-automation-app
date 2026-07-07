"""기준마진율 최적화 (두뇌④) — P×C마다 권장 기준마진율 작업목록 + 측정 루프(Gate3).

ADR 0026·0027 / workflows/margin-optimizer.md. 두뇌③(가격 A/B) 로더 재사용.
- 마진 = 순이익/정산액(택배 실배분 차감) — 채널마진모니터 기준마진율 시트와 동일 정의.
- 베이스 = 순이익 누적 85% proven 채널 순이익가중평균 / 볼륨×(마진vs베이스) 4분면 / 절반스텝.
- 게이트 🟢수락·🟡검토·🔴필수. 출력=권장값(저장은 사용자가 cmm/여기서). 선택→결정 원장(Gate3) 기록.
- 측정 루프: 결정 후 30일+매출 발생 → 결정일 이후 실적으로 효과 측정(개선=유지/악화=되돌림).
core/intelligence/margin_optimizer.py(순수함수) + decision_log.py 사용.
"""
import datetime as _dt
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import pandas as pd
import streamlit as st

from core.dashboard import store
from core.dashboard.sales_data import make_box_lookup
from core.intelligence import channel_compare as cc
from core.intelligence import decision_log as dl
from core.intelligence import margin_optimizer as mo
from core.intelligence import orders as _orders
from core.intelligence import ship_alloc
from core.intelligence import stockout

_REF = Path(__file__).parent.parent.parent / "reference"

st.set_page_config(layout="wide")  # 표 좌우 폭 확장(컬럼 많음)
st.title("🎯 기준마진율 최적화")
st.caption("상품×채널마다 **권장 기준마진율**을 데이터로. 손볼 가치 있는 소수만 임팩트(월순이익)순으로. "
           "🟢 수락 · 🟡 검토 · 🔴 사람 판단 필수. 적용은 **채널마진모니터**에서 저장(여기선 권장값·결정 기록).")


def _nfc(s) -> str:
    return unicodedata.normalize("NFC", str(s)).strip()


def _data_secret() -> tuple[str, str]:
    repo = "OURPROJECTDAO/work-automation-data"
    try:
        d = st.secrets["data"]
        return d["pat"], d.get("repo", repo)
    except Exception:
        return st.secrets.get("GITHUB_PAT", ""), repo


@st.cache_data(ttl=3600, show_spinner="매출 데이터 불러오는 중...")
def load_sales_min(pat: str, repo: str) -> pd.DataFrame:
    df = store.load_master(pat, repo)
    if df.empty:
        return df
    attr = pd.read_csv(_REF / "product_attributes.csv", dtype=str, encoding="utf-8-sig")
    pm = pd.read_csv(_REF / "product_master.csv", dtype=str, encoding="utf-8-sig")
    hap = {_nfc(k): v for k, v in zip(attr["관리코드"], attr["합포수량"])}

    def _hapq(c):
        try:
            f = float(str(hap.get(_nfc(c), "")).strip())
            return f if f > 0 else float("nan")
        except (TypeError, ValueError):
            return float("nan")

    boxq = make_box_lookup(pm)
    df["합포수량"] = df["관리코드"].map(_hapq)
    df["박스내품"] = df["관리코드"].map(boxq)
    return df


@st.cache_data(ttl=3600, show_spinner="EasyAdmin 주문 불러오는 중...")
def load_orders(pat: str, repo: str) -> pd.DataFrame:
    return _orders.read_all(pat, repo)


@st.cache_data(ttl=3600, show_spinner="송장 실배분 계산 중...")
def load_ship_rate(pat: str, repo: str) -> dict:
    od = load_orders(pat, repo)
    if od.empty:
        return {"rate": {}, "ch_rate": {}, "reconcile": {}, "stats": {"codes": 0}}
    pm = pd.read_csv(_REF / "product_master.csv", dtype=str, encoding="utf-8-sig")
    try:
        hapo = set(pd.read_csv(_REF / "hapo_175_190.csv", dtype=str,
                               encoding="utf-8-sig")["관리코드"].dropna())
    except Exception:
        hapo = None
    return ship_alloc.compute_ship_rate(od, load_sales_min(pat, repo),
                                        make_box_lookup(pm), hapo_codes=hapo)


@st.cache_data(ttl=3600, show_spinner=False)
def load_group_map(pat: str, repo: str) -> dict:
    g = store.read_groups(pat, repo)
    if g.empty:
        return {}
    out = {}
    for _, r in g.iterrows():
        nm, grp = r.get("상호명"), r.get("그룹")
        if pd.notna(nm) and pd.notna(grp) and str(grp).strip():
            out[_nfc(nm)] = str(grp).strip()
    return out


@st.cache_data(ttl=3600, show_spinner=False)
def load_giftset_codes() -> set:
    """product_attributes.csv 식품음료=='선물세트' 관리코드 — 명절세트는 별도관리(시즌·부분월) → 작업목록/측정 노이즈 제외(⑧)."""
    try:
        a = pd.read_csv(_REF / "product_attributes.csv", dtype=str, encoding="utf-8-sig")
        return {_nfc(c) for c, k in zip(a["관리코드"], a["식품음료"]) if _nfc(k) == "선물세트"}
    except Exception:
        return set()


@st.cache_data(ttl=3600, show_spinner=False)
def load_turnover(pat: str, repo: str) -> dict:
    """② 회전 신호 — 관리코드별 소진예측일(현재고÷최근3개월 일소진). 두뇌② forecast 재사용."""
    try:
        pm = pd.read_csv(_REF / "product_master.csv", dtype=str, encoding="utf-8-sig")
        sales = load_sales_min(pat, repo)
        if sales.empty:
            return {}
        dep = stockout.depletion_rate(sales, months=3)
        fc = stockout.forecast(pm, dep, {}, default_lead_days=14.0)  # 소진예측일은 cadence 무관
        if fc.empty:
            return {}
        return {_nfc(c): (float(d) if pd.notna(d) else None)
                for c, d in zip(fc["관리코드"], fc["소진예측일"])}
    except Exception:
        return {}


@st.cache_data(ttl=1800, show_spinner=False)
def load_locked() -> dict:
    """가격 제한(마진 민감) 상품 — reference/margin_floor.csv. {관리코드(NFC): 제한내용}.
    채널마진모니터 '제한 상품'과 동일 소스. 이들은 가격을 못 움직임 → 작업목록서 제외."""
    try:
        m = pd.read_csv(_REF / "margin_floor.csv", dtype=str, encoding="utf-8-sig")
        out = {}
        for _, r in m.iterrows():
            code = _nfc(str(r.get("관리코드") or ""))
            if not code:
                continue
            note = _nfc(str(r.get("제한내용") or "")) or _nfc(str(r.get("비고") or "")) or "가격 제한"
            out[code] = note
        return out
    except Exception:
        return {}


# 8개 관리채널(기준마진율 baseline 컬럼 보유) — 그 외(나들·제이티유통·11번가·셀러허브·리테일앤인사이트·멸치 등
# 미관리/유통 상호명)는 작업목록 스코프에서 제외(기준마진율 적용 대상 아님)
_MANAGED = {"스마트스토어", "ESM(G마켓·옥션)", "식봄", "캐시노트", "알리", "쿠팡", "배민상회", "올웨이즈"}


@st.cache_data(ttl=1800, show_spinner="온라인 매출 정제 중...")
def build_prod(pat: str, repo: str, unit: float) -> tuple[pd.DataFrame, dict]:
    """온라인 18개월 매출 → 택배 실배분·채널 라벨·나들 제외. 작업목록·측정 공용 토대."""
    df = load_sales_min(pat, repo)
    if df.empty:
        return pd.DataFrame()
    gmap = load_group_map(pat, repo)
    online = {s for s in df["상호명"].astype(str).unique() if gmap.get(_nfc(s)) == "온라인"}
    if not online:
        return pd.DataFrame()
    dmax = df["거래일자"].max()
    ts0 = pd.Timestamp((pd.Timestamp(dmax) - pd.DateOffset(months=18)).date())
    view = df[(df["거래일자"] >= ts0) & (df["상호명"].astype(str).isin(online))].copy()
    if view.empty:
        return pd.DataFrame()
    ship = load_ship_rate(pat, repo)
    prod = cc.compute_online_margin(view, ship, unit, use_actual=True)
    prod["채널"] = prod["상호명"].astype(str).map(cc.label)
    # ⑦ 나들 floor anchor — 나들 순마진(분수) per 관리코드, 제외 전에 1패스 계산
    _nd = prod[prod["채널"].astype(str).str.contains("나들", na=False)]
    nadl_map = {}
    if not _nd.empty:
        _g = (_nd.assign(_c=_nd["관리코드"].astype(str).map(_nfc))
                 .groupby("_c").agg(_s=("_순", "sum"), _r=("판매금액", "sum")))
        nadl_map = {c: float(r._s / r._r) for c, r in _g.iterrows() if r._r > 0}
    prod = prod[prod["채널"].astype(str).isin(_MANAGED)]  # 8 관리채널만(나들·제이티·11번가·셀러허브·리테일·멸치 등 제외)
    _gs = load_giftset_codes()  # ⑧ 선물세트(명절세트) 제외 — 별도관리·시즌 노이즈
    if _gs:
        prod = prod[~prod["관리코드"].astype(str).map(_nfc).isin(_gs)]
    return prod, nadl_map


@st.cache_data(ttl=1800, show_spinner="권장 기준마진율 계산 중...")
def build_worklist(pat: str, repo: str, unit: float) -> tuple[pd.DataFrame, pd.DataFrame]:
    prod, nadl_map = build_prod(pat, repo, unit)
    if prod.empty:
        return pd.DataFrame(), pd.DataFrame()
    cells = mo.cell_stats(prod)
    return mo.worklist(cells, nadl_map=nadl_map, turnover_map=load_turnover(pat, repo)), cells


# 두뇌④ 채널 라벨 → baseline_margin.csv 컬럼 (ESM 라벨만 다름)
_BCOL = {"스마트스토어": "스마트스토어", "ESM(G마켓·옥션)": "ESM", "식봄": "식봄",
         "캐시노트": "캐시노트", "알리": "알리", "쿠팡": "쿠팡",
         "배민상회": "배민상회", "올웨이즈": "올웨이즈"}
_HELP_LOGIC = """**기준(베이스)** = 이 상품을 가장 잘 파는 핵심 채널들(이익 누적 85%)의 순이익 가중평균 마진.

**채널마다 (위에서부터 먼저 걸리는 것):**
- 이익 비중 <1% & 핵심 아님 → **가격 테스트** ⚪ (작으니 한번 찔러봄)
- 핵심 채널인데 마진<기준 → **올림**(기준까지 절반) 🟢 / 마진 적정 → **유지** 🟢
- 적게 파는데 비쌈(마진>기준) → **내림**(기준까지 절반) 🟡
- 많이 파는데(≥10%) 쌈(마진<기준) → **올림**(기준까지 절반) 🟡
- 그 외(적게 파는데 쌈) → **낮게 유지** 🔴 (가격 문제 아님)

**덮어쓰기:**
- 📉**매출목표 미달**(가장 잘 파는 채널도 월 50만 미만) → 더 팔리게 **나들 마진까지 절반 내림** 🟡 / 이미 쌈이면 🔴
- ⚪**재고 안 빠짐**(소진예측 180일 초과) → **재고정리 내림**(360일=−2%p까지·재고 풀리면 자동 원복) 🟡 / 더 못 내리면 🔴

**마무리:** 변동 <0.3%p → 유지 🟢 · 변동 >3%p → 🔴 · 신호 약함(판매 6개월 미만/마진 거의 안 흔듦) → 🟡

**🟢 수락**=확신 높고 변동 작음(일괄) · **🟡 검토**=보통 변동(한번 보고) · **🔴 필수**=큰 변동·쌈+안팔림·가격 외 요인·충돌(사람 판단).

*항상: 절반씩만 · 역마진 금지·**3% 절대 하한**(3%까지만 내리고 그 밑은 막음 · 이미 3% 이하라 더 못 내리면 작업목록서 제외) · 전부 사람 저장 · 측정 루프가 다음 사이클에 검증(악화→되돌림) · 선물세트·나들·미관리 채널 제외.*"""
_BASELINE_PATH = "reference/baseline_margin.csv"
_APP_REPO = "OURPROJECTDAO/work-automation-app"


def _app_pat() -> str:
    return st.secrets.get("GITHUB_PAT", "")


@st.cache_data(ttl=600, show_spinner=False)
def load_baseline(pat_app: str) -> dict:
    """baseline_margin.csv → {(관리코드, 채널컬럼): 분수}. 현 cmm 기준마진율(타깃)."""
    if not pat_app:
        return {}
    import base64 as _b64, io as _io, json as _js, urllib.parse as _uq, urllib.request as _ur
    url = "https://api.github.com/repos/%s/contents/%s" % (
        _APP_REPO, "/".join(_uq.quote(s) for s in _BASELINE_PATH.split("/")))
    try:
        req = _ur.Request(url, headers={"Authorization": "Bearer " + pat_app,
                          "Accept": "application/vnd.github+json", "User-Agent": "mo"})
        with _ur.urlopen(req) as r:
            meta = _js.load(r)
        df = pd.read_csv(_io.StringIO(_b64.b64decode(meta["content"]).decode("utf-8-sig")), dtype=str)
        cc0 = df.columns[0]
        out = {}
        for _, row in df.iterrows():
            code = str(row[cc0])
            for col in df.columns[1:]:
                val = row[col]
                if pd.notna(val) and str(val).strip():
                    try:
                        out[(code, col)] = float(val)
                    except (TypeError, ValueError):
                        pass
        return out
    except Exception:
        return {}


def _apply_baseline(updates):
    """updates=[(관리코드, baseline컬럼, Δ분수)]. 기존 타깃 + Δ 로 갱신(절대 덮어쓰기 아님) → cmm 반영.
    되돌림은 Δ에 역부호(타깃 + (before−applied))를 넘기면 원복. return ((applied, miss_row), None) | (None, err)."""
    import base64 as _b64, io as _io, json as _js, urllib.parse as _uq, urllib.request as _ur
    pat_app = _app_pat()
    if not pat_app:
        return None, "GITHUB_PAT(앱 repo 쓰기) secrets 없음"
    url = "https://api.github.com/repos/%s/contents/%s" % (
        _APP_REPO, "/".join(_uq.quote(s) for s in _BASELINE_PATH.split("/")))
    h = {"Authorization": "Bearer " + pat_app, "Accept": "application/vnd.github+json",
         "User-Agent": "mo"}
    with _ur.urlopen(_ur.Request(url, headers=h)) as r:
        meta = _js.load(r)
    bdf = pd.read_csv(_io.StringIO(_b64.b64decode(meta["content"]).decode("utf-8-sig")), dtype=str)
    code_col = bdf.columns[0]
    bdf = bdf.set_index(code_col)
    idx = set(bdf.index.astype(str))
    applied, miss_row = [], []
    for code, col, dlt in updates:
        if col not in bdf.columns or str(code) not in idx:
            miss_row.append(code)
            continue
        try:
            cur = float(bdf.loc[str(code), col])
        except (TypeError, ValueError):
            miss_row.append(code)
            continue
        new = max(0.0, cur + dlt)
        bdf.loc[str(code), col] = (f"{new:.4f}").rstrip("0").rstrip(".")
        applied.append((code, col))
    buf = _io.StringIO()
    bdf.reset_index().to_csv(buf, index=False, lineterminator="\r\n")
    body = {"message": "data(baseline): 두뇌④ Δ가산/원복 적용(%d건)" % len(applied),
            "content": _b64.b64encode(("\ufeff" + buf.getvalue()).encode("utf-8")).decode(),
            "sha": meta["sha"]}
    with _ur.urlopen(_ur.Request(url, method="PUT", data=_js.dumps(body).encode(), headers=h)) as r:
        _js.load(r)
    return (applied, miss_row), None


pat, repo = _data_secret()
if not pat:
    st.warning("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    st.stop()

unit = 2700.0  # 실택배비 표준(daily_margin.DEFAULT_FLAT·채널마진모니터와 동일)
st.session_state.setdefault("mo_recent", set())   # 이번 세션 방금 기록한 (관리코드,채널) — 즉시 숨김(원장 복제 지연 대비)
st.session_state.setdefault("mo_tblver", 0)        # 작업목록 표 버전(기록 후 선택 초기화용)

tab_wl, tab_ms = st.tabs(["📋 작업목록", "📈 측정 결과 (Gate3)"])

# ════════════════════════════════════════════════════════════════════
# 탭 1 — 작업목록 (권장 기준마진율)
# ════════════════════════════════════════════════════════════════════
with tab_wl:
    with st.expander("ⓘ 권장 기준마진율은 어떻게 정해지나 — 한눈에 보기"):
        st.markdown(_HELP_LOGIC)
    wl, cells = build_worklist(pat, repo, unit)
    if wl.empty:
        st.info("적재된 온라인 매출 데이터가 없습니다. (거래처 그룹에 '온라인' 지정 필요)")
    else:
        ACT = [mo.A_UP, mo.A_DOWN, mo.A_HOLDLOW, mo.A_TURN]
        act = wl[wl["액션"].isin(ACT)].copy()
        queue = wl[wl["액션"] == mo.A_QUEUE]

        # 45일 억제: 최근 결정한 (관리코드,채널)은 측정창 동안 숨김
        _SUPPRESS_DAYS = 45
        _sup = set()
        try:
            _led = dl.read_all(pat, repo)
            if not _led.empty and "ts" in _led.columns:
                _cut = (_dt.date.today() - _dt.timedelta(days=_SUPPRESS_DAYS)).isoformat()
                _r = _led[(_led["ts"].astype(str) >= _cut)
                          & (_led["status"].astype(str) == "pending")]
                _sup = set(zip(_r["관리코드"].astype(str), _r["채널"].astype(str)))
        except Exception:
            _sup = set()
        _sup |= st.session_state.get("mo_recent", set())  # 방금 기록분 즉시 숨김
        act["_k"] = list(zip(act["관리코드"].astype(str), act["채널"].astype(str)))
        _n_sup_shown = int(act["_k"].isin(_sup).sum()) if _sup else 0
        act = act[~act["_k"].isin(_sup)].drop(columns="_k").copy()
        # 가격 제한(마진 민감) 상품 = 못 건드림 → 작업목록서 빼고 🔒 별도 묶음(baseline 미적용)
        _lock = load_locked()
        act["_nfc"] = act["관리코드"].astype(str).map(_nfc)
        act_lock = act[act["_nfc"].isin(_lock)].drop(columns="_nfc").copy()
        act = act[~act["_nfc"].isin(_lock)].drop(columns="_nfc").copy()
        # 변화 없음(권장=현재 · 주로 '낮게 유지' 🔴 = 가격 외 요인) → 별도 묶음
        act_none = act[act["Δ"].abs() < 0.05].copy()
        act = act[act["Δ"].abs() >= 0.05].copy()

        # ── 필터 ───────────────────────────────────────────
        chans = sorted(act["채널"].unique())
        # 채널 — 칩(pills) 다중선택. 전부 선택(기본)=전체. 칩 없는 구버전이면 멀티셀렉트 폴백.
        if hasattr(st, "pills"):
            _sel = st.pills("채널", chans, selection_mode="multi", default=chans, key="mo_ch_pills")
            pick_ch = list(_sel) if _sel else []
        else:
            pick_ch = st.multiselect("채널", chans, default=[], key="mo_ch", placeholder="전체")
        if pick_ch and len(pick_ch) >= len(chans):
            pick_ch = []  # 전부 선택 = 필터 없음(전체)
        f2, f3 = st.columns([1.2, 2.8])
        with f2:
            flags = ["🟢", "🟡", "🔴"]
            pick_fl = st.pills("플래그", flags, selection_mode="multi", default=flags,
                               key="mo_fl_pills") if hasattr(st, "pills") else \
                st.multiselect("플래그", flags, default=flags, key="mo_fl")
            if not pick_fl:
                pick_fl = flags  # 전부 해제 = 전체
        with f3:
            q = st.text_input("상품명 / 관리코드 검색", key="mo_q", placeholder="예: 스팸 · 15-04").strip()

        def _apply_chq(df):
            d = df
            if pick_ch:
                d = d[d["채널"].isin(pick_ch)]
            if q:
                qn = _nfc(q)
                d = d[d["상품명"].astype(str).map(_nfc).str.contains(qn, case=False, na=False)
                      | d["관리코드"].astype(str).str.contains(qn, case=False, na=False)]
            return d
        scoped = _apply_chq(act)        # KPI 기준 = 채널+검색(플래그 무관)
        none_v = _apply_chq(act_none)   # 변화없음 묶음도 채널+검색 따라감
        lock_v = _apply_chq(act_lock)   # 가격제한 묶음도 채널+검색 따라감
        if len(lock_v):
            lock_v = lock_v.copy()
            lock_v["제한내용"] = lock_v["관리코드"].astype(str).map(_nfc).map(_lock)
        v = scoped.copy()
        if pick_fl:
            v = v[v["플래그"].isin(pick_fl)]
        v = v.reset_index(drop=True)
        _rev = cells.copy()
        _rev["월매출"] = (_rev["매출"] / _rev["개월"].clip(lower=1)).round().astype("int64")
        v = v.merge(_rev[["관리코드", "채널", "월매출"]], on=["관리코드", "채널"], how="left")
        v["월매출"] = v["월매출"].fillna(0).astype("int64")
        # 현 기준마진율(cmm 타깃) — baseline_margin.csv 채널별 조회
        _bl = load_baseline(_app_pat())
        v["기준마진율"] = [
            (round(_bl[(str(c), _BCOL[ch])] * 100, 2)
             if (ch in _BCOL and (str(c), _BCOL[ch]) in _bl) else float("nan"))
            for c, ch in zip(v["관리코드"], v["채널"])]

        # ── KPI ────────────────────────────────────────────
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("손볼 것", f"{len(scoped):,}")
        k2.metric("🟢 수락", f"{(scoped['플래그']=='🟢').sum():,}")
        k3.metric("🟡 검토", f"{(scoped['플래그']=='🟡').sum():,}")
        k4.metric("🔴 필수", f"{(scoped['플래그']=='🔴').sum():,}")
        st.caption(f"숫자판=채널·검색 필터 반영 · 🔒가격제한 {len(lock_v):,} · 변화없음(검토만) {len(none_v):,} · 실험큐(관망) {len(queue):,} · "
                   f"측정중(45일 숨김) {_n_sup_shown:,} · 📉매출목표 미달 {int((scoped['목표']=='📉미달').sum()):,} · "
                   f"선물세트 {len(load_giftset_codes()):,}종 제외(⑧) — 임팩트(월순이익)순. 행 선택 → 아래에서 기록/적용.")

        if len(lock_v):
            with st.expander(f"🔒 가격 제한(마진 민감) {len(lock_v):,}건 — 가격 못 움직임(채널마진모니터 제한상품) · 기준마진율 변경 대상 아님"):
                _lshow = lock_v.sort_values("월순이익", ascending=False)[
                    ["관리코드", "상품명", "채널", "현재마진", "베이스", "월순이익", "제한내용"]]
                st.dataframe(_lshow, width="stretch", hide_index=True,
                             column_config={
                                 "현재마진": st.column_config.NumberColumn("현재(%)", format="%.1f"),
                                 "베이스": st.column_config.NumberColumn("베이스(%)", format="%.1f"),
                                 "월순이익": st.column_config.NumberColumn("월순이익", format="localized"),
                                 "제한내용": st.column_config.TextColumn("제한내용", width="medium")},
                             height=min(400, 80 + 36 * min(len(_lshow), 10)))
                st.caption("margin_floor.csv에 '마진율 민감 상품'으로 등록된 건(단가유지·배송비포함가 등). 가격을 못 바꾸므로 추천·baseline 변경 안 함 — 등록/해제는 채널마진모니터 쪽 margin_floor 관리.")

        if len(none_v):
            with st.expander(f"🔴 변화 없음 · 가격 외 요인 {len(none_v):,}건 — 이미 싼데 안 팔림(노출·상세·광고/단종 판단)"):
                _nshow = none_v.sort_values("월순이익", ascending=False)[
                    ["플래그", "목표", "관리코드", "상품명", "채널", "현재마진", "베이스", "월순이익", "사유"]]
                st.dataframe(_nshow, width="stretch", hide_index=True,
                             column_config={
                                 "현재마진": st.column_config.NumberColumn("현재(%)", format="%.1f"),
                                 "베이스": st.column_config.NumberColumn("베이스(%)", format="%.1f"),
                                 "월순이익": st.column_config.NumberColumn("월순이익", format="localized"),
                                 "사유": st.column_config.TextColumn(width="large")},
                             height=min(400, 80 + 36 * min(len(_nshow), 10)))
                st.caption("가격을 더 내려도 효과 적다고 본 행(권장=현재) → 기준마진율 변경 대상 아님. 노출/상세/광고/단종 등 가격 외 레버로 판단.")

        # ── 작업목록 ────────────────────────────────────────
        disp = v.copy()
        disp["변화"] = disp["Δ"].map(
            lambda x: f"▲ +{x:.1f}" if x > 0 else (f"▼ {x:.1f}" if x < 0 else "—"))
        show = disp[["플래그", "목표", "관리코드", "상품명", "채널", "현재마진", "베이스", "권장마진",
                     "변화", "기준마진율", "월매출", "월순이익", "월볼륨", "액션", "사유"]]

        def _color_dir(s):
            if isinstance(s, str) and s.startswith("▲"):
                return "color:#cf222e;font-weight:700"
            if isinstance(s, str) and s.startswith("▼"):
                return "color:#0969da;font-weight:700"
            return "color:#8c959f"

        styled = show.style.map(_color_dir, subset=["변화"])
        ev = st.dataframe(
            styled, width="stretch", hide_index=True,
            on_select="rerun", selection_mode="multi-row",
            column_config={
                "플래그": st.column_config.TextColumn(width="small"),
                "목표": st.column_config.TextColumn(width="small", help="best 관리채널 월매출<50만 = 매출목표 미달(⑦) → 나들 마진까지 절반스텝 인하·저마진은 🔴"),
                "관리코드": st.column_config.TextColumn(width="small"),
                "상품명": st.column_config.TextColumn(width="medium"),
                "현재마진": st.column_config.NumberColumn("현재(%)", format="%.1f"),
                "베이스": st.column_config.NumberColumn("베이스(%)", format="%.1f"),
                "권장마진": st.column_config.NumberColumn("권장(%)", format="%.1f"),
                "변화": st.column_config.TextColumn("권장변화(%p)", width="small"),
                "기준마진율": st.column_config.NumberColumn("기준(현)%", format="%.2f", help="현 cmm 기준마진율(타깃). 변경 시 여기에 변화량(Δ)을 더함"),
                "월매출": st.column_config.NumberColumn("월매출", format="localized"),
                "월순이익": st.column_config.NumberColumn("월순이익", format="localized"),
                "월볼륨": st.column_config.NumberColumn("월볼륨", format="localized"),
                "사유": st.column_config.TextColumn(width="large"),
            },
            height=min(620, 80 + 36 * min(len(show), 15)),
            key=f"mo_tbl_{st.session_state['mo_tblver']}")

        sel_rows = []
        try:
            sel_rows = [i for i in ev.selection.rows if 0 <= i < len(v)]
        except Exception:
            sel_rows = []

        sel_df = v.iloc[sel_rows] if sel_rows else v.iloc[0:0]
        chg = True  # 기록 시 항상 기준마진율(baseline) 함께 변경 → cmm 반영(토글 제거)
        if len(sel_rows):
            prv = sel_df[["관리코드", "상품명", "채널", "기준마진율"]].copy()
            prv["변화(%p)"] = (sel_df["권장마진"] - sel_df["현재마진"]).round(1)
            prv["새기준(%)"] = (prv["기준마진율"] + prv["변화(%p)"]).round(2)
            prv["반영"] = ["✔ cmm" if (ch in _BCOL and pd.notna(bm)) else "기록만"
                          for ch, bm in zip(sel_df["채널"], sel_df["기준마진율"])]
            st.caption("↓ 현 기준마진율(타깃)에 **변화량(Δ)**을 더해 갱신합니다(기준값 없는 행은 기록만). 확인 후 버튼.")
            st.dataframe(prv, hide_index=True, width="stretch",
                         column_config={"기준마진율": st.column_config.NumberColumn("기준(현)%", format="%.2f"),
                                        "새기준(%)": st.column_config.NumberColumn("새기준%", format="%.2f")})

        cA, cB = st.columns([1, 1])
        with cA:
            st.download_button("작업목록 CSV", v.to_csv(index=False).encode("utf-8-sig"),
                               file_name="기준마진율_권장.csv", mime="text/csv", key="mo_dl")
        with cB:
            if st.button(f"📝 기록 + 기준마진율 변경 ({len(sel_rows)}건)", disabled=len(sel_rows) == 0, key="mo_rec"):
                recs = []
                for i in sel_rows:
                    r = v.iloc[i]
                    _bef = float(r["현재마진"]) / 100
                    _rec = float(r["권장마진"]) / 100
                    recs.append(dict(
                        관리코드=r["관리코드"], 채널=r["채널"], 액션=r["액션"],
                        마진_before=_bef, 마진_권장=_rec, 마진_적용=(_rec if chg else _bef),
                        베이스=float(r["베이스"]) / 100, 플래그=r["플래그"], 사유=r["사유"],
                        측정전_월볼륨=int(r["월볼륨"]), 측정전_월순이익=int(r["월순이익"])))
                try:
                    n = dl.append(recs, pat, repo)
                    msg = f"결정 원장 {n}건 기록(45일 숨김)."
                    if chg:
                        ups, skip = [], []
                        for _, r in sel_df.iterrows():
                            col = _BCOL.get(r["채널"])
                            if col is None:
                                skip.append(r["채널"])
                                continue
                            ups.append((r["관리코드"], col, (float(r["권장마진"]) - float(r["현재마진"])) / 100))
                        res, err = _apply_baseline(ups)
                        if err:
                            st.warning("기준마진율 변경 실패: " + err + " (기록은 완료)")
                        else:
                            applied, miss_row = res
                            msg += f" 기준마진율 {len(applied)}건 변경(타깃+Δ) → cmm 반영(최대 10분 내)."
                            if miss_row:
                                msg += f" · 기준값 없음 {len(miss_row)}건(기록만)."
                            if skip:
                                msg += f" · 매핑외 채널 건너뜀: {', '.join(sorted(set(skip)))}."
                    # 방금 처리분 즉시 숨김 + 표 선택 초기화 + 새로고침(필터는 유지) → 이어서 진행
                    for _i in sel_rows:
                        _rr = v.iloc[_i]
                        st.session_state["mo_recent"].add((str(_rr["관리코드"]), str(_rr["채널"])))
                    st.session_state["mo_tblver"] += 1
                    if chg:
                        load_baseline.clear()  # 기준마진율 컬럼 최신화
                    st.toast("✅ " + msg + " 이어서 진행하세요.")
                    st.rerun()
                except Exception as e:
                    st.error(f"실패: {e}")

        st.divider()
        st.caption(
            "ⓘ **마진=순이익÷정산액**(택배 실배분 차감·기준마진율 시트 정의). **베이스**=순이익 누적 85% proven 채널 "
            "순이익가중평균. **↑/↓ 절반스텝**=베이스까지 거리의 절반만(반응 보고 또 절반/되돌림). **hold-low**=싼데 안 "
            "팔림→안 올림. **🔴**=|Δ|>3%p·신호 약함·hold-low(사람 판단). 권장변화 ▲올림(빨강)·▼내림(파랑)·한국식. "
            "v0 한계: ⑧ 시즌(명절세트) 미보정으로 월순이익 상위에 세트류 부풀려질 수 있음 · 나들=하한 참조(별도) · "
            "기록하면 baseline_margin.csv에 Δ(권장−현재)를 가산해 cmm 반영(기준값 없는 낱개/소분류 행은 기록만)·같은 결정은 45일 측정창 동안 숨김. "
            "**↓ 회전**=장기소진(소진예측>180일·두뇌②) 청산 마크다운(−2%p 캡·재고풀리면 자동복귀). **📉목표**=best 채널 월매출<50만(⑦·나들 마진까지). 택배비 2,700원 고정·나들 제외·선물세트 제외.")

# ════════════════════════════════════════════════════════════════════
# 탭 2 — 측정 결과 (Gate 3: 결정 후 실적으로 효과 측정 → 유지/되돌림)
# ════════════════════════════════════════════════════════════════════
with tab_ms:
    st.caption(f"기록한 결정의 **결정일 이후 실적**으로 효과를 측정합니다. **적재된 매출이 결정일 기준 {mo.MEASURE_MIN_DAYS}일분** "
               "쌓이면 측정 대상(월별 적재 갭에 안 휘둘리도록 벽시계 아닌 **데이터 커버리지** 기준). "
               "**개선=유지 · 악화=되돌림(기준마진율 원복)**. 판정은 제안 — 사람이 최종 (Gate 3: 잘못된 결정 리뷰).")
    led = dl.read_all(pat, repo)
    if led is None or led.empty:
        st.info("기록된 결정이 없습니다. **작업목록** 탭에서 결정을 기록하면 여기서 측정합니다.")
    else:
        prod, _ = build_prod(pat, repo, unit)
        pend = led[led["status"].astype(str) == "pending"].copy()
        meas = led[led["status"].astype(str) == "measured"].copy()
        closed_n = int((led["status"].astype(str).isin(["closed", "reverted"])).sum())

        if prod.empty:
            st.info("온라인 매출 데이터가 없어 측정할 수 없습니다.")
        else:
            mdf = mo.measure(prod, pend)  # pending 비면 빈 결과
            ready = mdf[mdf["ready"]].reset_index(drop=True) if not mdf.empty else mdf
            wait = mdf[~mdf["ready"]].reset_index(drop=True) if not mdf.empty else mdf

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("측정 가능", f"{len(ready):,}")
            m2.metric("개선", f"{int((ready['결과'] == '개선').sum()) if len(ready) else 0:,}")
            m3.metric("악화", f"{int((ready['결과'] == '악화').sum()) if len(ready) else 0:,}")
            m4.metric("대기", f"{len(wait):,}")
            st.caption(f"측정완료(미조치) {len(meas):,} · 종료(유지·되돌림) {closed_n:,} · 전체 결정 {len(led):,}")

            # ── 측정 가능(pending·ready) → 확정 ─────────────────
            st.markdown("##### 측정 가능 — 확정하면 원장에 측정후·결과 기록")
            if len(ready):
                rshow = ready[["결과", "제안", "관리코드", "상품명", "채널", "액션",
                               "측정전_월순이익", "측정후_월순이익", "측정전_월볼륨",
                               "측정후_월볼륨", "측정후마진", "측정일수", "post_개월", "플래그"]].copy()
                rev = st.dataframe(
                    rshow, width="stretch", hide_index=True,
                    on_select="rerun", selection_mode="multi-row",
                    column_config={
                        "결과": st.column_config.TextColumn(width="small"),
                        "제안": st.column_config.TextColumn(width="small"),
                        "관리코드": st.column_config.TextColumn(width="small"),
                        "측정전_월순이익": st.column_config.NumberColumn("전·월순이익", format="localized"),
                        "측정후_월순이익": st.column_config.NumberColumn("후·월순이익", format="localized"),
                        "측정전_월볼륨": st.column_config.NumberColumn("전·월볼륨", format="localized"),
                        "측정후_월볼륨": st.column_config.NumberColumn("후·월볼륨", format="localized"),
                        "측정후마진": st.column_config.NumberColumn("후마진%", format="%.1f"),
                        "측정일수": st.column_config.NumberColumn("측정일수", format="%d", help="run-rate에 쓴 적재 post 기간(일). 30.4일=1개월로 일수 정규화"),
                    },
                    height=min(560, 80 + 36 * min(len(rshow), 12)), key="ms_ready")
                try:
                    rsel = [i for i in rev.selection.rows if 0 <= i < len(ready)]
                except Exception:
                    rsel = []
                if st.button(f"📌 측정 확정 ({len(rsel)}건)", disabled=len(rsel) == 0, key="ms_confirm"):
                    recs = []
                    today = _dt.date.today().isoformat()
                    for i in rsel:
                        rr = ready.iloc[i]
                        recs.append(dict(
                            decision_id=rr["decision_id"], 측정일=today,
                            측정후_월볼륨=(int(rr["측정후_월볼륨"]) if pd.notna(rr["측정후_월볼륨"]) else None),
                            측정후_월순이익=(int(rr["측정후_월순이익"]) if pd.notna(rr["측정후_월순이익"]) else None),
                            결과=rr["결과"], status="measured"))
                    try:
                        n = dl.update(recs, pat, repo)
                        st.success(f"{n}건 측정 확정 — 아래 '측정 완료'에서 유지/되돌림 결정.")
                    except Exception as e:
                        st.error(f"실패: {e}")
            else:
                st.caption("측정 가능한 결정이 아직 없습니다.")

            if len(wait):
                with st.expander(f"측정 대기 {len(wait)}건 — 적재 post 기간 부족(<{mo.MEASURE_MIN_DAYS}일)"):
                    st.caption("측정일수 = 적재 최신거래일 − 결정일. 다음 월 매출이 적재되면 자동으로 측정 가능으로 넘어옵니다.")
                    st.dataframe(
                        wait[["관리코드", "상품명", "채널", "액션", "측정일수", "경과일", "post_개월", "플래그"]],
                        hide_index=True, width="stretch",
                        column_config={"측정일수": st.column_config.NumberColumn("측정일수(적재)", format="%d")})

            st.divider()

            # ── 측정 완료(measured) → 사람 조치: 유지 / 되돌림 ──
            st.markdown("##### 측정 완료 — 유지로 닫거나 되돌림(기준마진율 원복)")
            if meas.empty:
                st.caption("측정 완료 후 미조치 결정이 없습니다.")
            else:
                mv = meas.copy()
                for c in ["측정전_월순이익", "측정후_월순이익", "측정전_월볼륨", "측정후_월볼륨"]:
                    if c in mv.columns:
                        mv[c] = pd.to_numeric(mv[c], errors="coerce")
                mv["_적용여부"] = [
                    "✔ 변경됨" if (pd.notna(a) and pd.notna(b) and abs(float(a) - float(b)) > 1e-9) else "기록만"
                    for a, b in zip(mv["마진_적용"], mv["마진_before"])]
                mshow = mv[["결과", "관리코드", "채널", "액션", "_적용여부",
                            "측정전_월순이익", "측정후_월순이익", "측정일", "ts"]].copy()
                mev = st.dataframe(
                    mshow, width="stretch", hide_index=True,
                    on_select="rerun", selection_mode="multi-row",
                    column_config={
                        "결과": st.column_config.TextColumn(width="small"),
                        "관리코드": st.column_config.TextColumn(width="small"),
                        "_적용여부": st.column_config.TextColumn("기준변경", width="small"),
                        "측정전_월순이익": st.column_config.NumberColumn("전·월순이익", format="localized"),
                        "측정후_월순이익": st.column_config.NumberColumn("후·월순이익", format="localized"),
                        "ts": st.column_config.TextColumn("결정일", width="small"),
                    },
                    height=min(480, 80 + 36 * min(len(mshow), 10)), key="ms_meas")
                try:
                    msel = [i for i in mev.selection.rows if 0 <= i < len(mv)]
                except Exception:
                    msel = []

                # 되돌림 미리보기(원복 payload) — 실제 변경됐던 행만
                if msel:
                    _bl2 = load_baseline(_app_pat())
                    prev_rows = []
                    for i in msel:
                        rr = mv.iloc[i]
                        col = _BCOL.get(rr["채널"])
                        bef = float(rr["마진_before"] or 0)
                        app = float(rr["마진_적용"] or 0)
                        if col is None or abs(app - bef) <= 1e-9:
                            continue  # 기록만 → 원복 대상 아님
                        cur = _bl2.get((str(rr["관리코드"]), col))
                        rev_delta = bef - app  # 역가산
                        prev_rows.append(dict(
                            관리코드=rr["관리코드"], 채널=rr["채널"],
                            현기준=round(cur * 100, 2) if cur is not None else float("nan"),
                            원복후=round((cur + rev_delta) * 100, 2) if cur is not None else float("nan")))
                    if prev_rows:
                        st.caption("↩ 되돌림 대상(기준마진율 원복) — 현 타깃에 **−Δ**(원래 변화 역가산):")
                        st.dataframe(pd.DataFrame(prev_rows), hide_index=True, width="stretch",
                                     column_config={"현기준": st.column_config.NumberColumn("현 기준%", format="%.2f"),
                                                    "원복후": st.column_config.NumberColumn("원복 후%", format="%.2f")})

                ck, cr = st.columns(2)
                with ck:
                    if st.button(f"✅ 유지로 닫기 ({len(msel)}건)", disabled=len(msel) == 0, key="ms_keep"):
                        recs = [dict(decision_id=mv.iloc[i]["decision_id"], status="closed") for i in msel]
                        try:
                            n = dl.update(recs, pat, repo)
                            st.success(f"{n}건 유지로 닫힘.")
                        except Exception as e:
                            st.error(f"실패: {e}")
                with cr:
                    if st.button(f"↩ 되돌림 ({len(msel)}건)", disabled=len(msel) == 0, key="ms_revert"):
                        ups = []
                        for i in msel:
                            rr = mv.iloc[i]
                            col = _BCOL.get(rr["채널"])
                            bef = float(rr["마진_before"] or 0)
                            app = float(rr["마진_적용"] or 0)
                            if col is not None and abs(app - bef) > 1e-9:
                                ups.append((rr["관리코드"], col, (bef - app)))  # 역가산 원복
                        recs = [dict(decision_id=mv.iloc[i]["decision_id"], status="reverted") for i in msel]
                        try:
                            err = None
                            applied = []
                            if ups:
                                res, err = _apply_baseline(ups)
                                if not err:
                                    applied, _miss = res
                            n = dl.update(recs, pat, repo)
                            if err:
                                st.warning(f"원장 {n}건 reverted 기록 · 기준마진율 원복 실패: {err}")
                            else:
                                msg = f"되돌림 {n}건 — 기준마진율 {len(applied)}건 원복(−Δ) → cmm 반영(최대 10분 내)."
                                if not ups:
                                    msg = f"되돌림 {n}건 — 기록만 결정이라 원복할 기준변경 없음(원장만 reverted)."
                                st.success(msg)
                        except Exception as e:
                            st.error(f"실패: {e}")

    st.divider()
    st.caption(
        "ⓘ 측정후 = **결정일(ts) 이후** 거래만 월 run-rate(동일 마진정의). **일수 정규화**(합÷측정일수/30.4 — 달력월 개수 아님) + "
        f"**커버리지 게이트**(적재 post 기간 ≥{mo.MEASURE_MIN_DAYS}일이면 측정; 월별 적재 갭에 안 휘둘림). "
        f"개선=후월순이익 > 전×{1 + mo.RESP_BAND:.0%} · 악화 < 전×{1 - mo.RESP_BAND:.0%} · 그 사이=무변화(관찰). "
        "행사·시즌·품절 교란은 미보정(관찰 한계) — 판정은 제안. **되돌림**은 적용했던 변화(Δ)를 baseline에서 "
        "역가산 원복(기록만 결정은 원복 없이 종료). 원장 status: pending→measured→closed/reverted (forward·비-PII).")
