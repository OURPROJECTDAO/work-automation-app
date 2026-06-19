"""기준마진율 최적화 (두뇌④) — P×C마다 권장 기준마진율 작업목록.

ADR 0026 / workflows/margin-optimizer.md. 두뇌③(가격 A/B) 로더 재사용.
- 마진 = 순이익/정산액(택배 실배분 차감) — 채널마진모니터 기준마진율 시트와 동일 정의.
- 베이스 = 순이익 누적 85% proven 채널 순이익가중평균 / 볼륨×(마진vs베이스) 4분면 / 절반스텝.
- 게이트 🟢수락·🟡검토·🔴필수. 출력=권장값(저장은 사용자가 채널마진모니터에서). 선택→결정 원장(Gate3) 기록.
core/intelligence/margin_optimizer.py(순수함수) + decision_log.py 사용. v0: 권장 산출·원장 기록까지(cmm 편집창 직접 prefill은 후속).
"""
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

_REF = Path(__file__).parent.parent.parent / "reference"

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


@st.cache_data(ttl=1800, show_spinner="권장 기준마진율 계산 중...")
def build_worklist(pat: str, repo: str, unit: float) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = load_sales_min(pat, repo)
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()
    gmap = load_group_map(pat, repo)
    online = {s for s in df["상호명"].astype(str).unique() if gmap.get(_nfc(s)) == "온라인"}
    if not online:
        return pd.DataFrame(), pd.DataFrame()
    dmax = df["거래일자"].max()
    ts0 = pd.Timestamp((pd.Timestamp(dmax) - pd.DateOffset(months=18)).date())
    view = df[(df["거래일자"] >= ts0) & (df["상호명"].astype(str).isin(online))].copy()
    if view.empty:
        return pd.DataFrame(), pd.DataFrame()
    ship = load_ship_rate(pat, repo)
    prod = cc.compute_online_margin(view, ship, unit, use_actual=True)
    prod["채널"] = prod["상호명"].astype(str).map(cc.label)
    prod = prod[~prod["채널"].astype(str).str.contains("나들", na=False)]  # 나들=데이터 확인용·미관리 채널 → 제외
    cells = mo.cell_stats(prod)
    return mo.worklist(cells), cells


pat, repo = _data_secret()
if not pat:
    st.warning("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    st.stop()

unit = 2700.0  # 실택배비 표준(daily_margin.DEFAULT_FLAT·채널마진모니터와 동일)

wl, cells = build_worklist(pat, repo, unit)
if wl.empty:
    st.info("적재된 온라인 매출 데이터가 없습니다. (거래처 그룹에 '온라인' 지정 필요)")
    st.stop()

ACT = ["↑ 절반스텝", "↓ 절반스텝", "hold-low"]
act = wl[wl["액션"].isin(ACT)].copy()
queue = wl[wl["액션"] == "실험큐"]

# 45일 억제: 최근 결정한 (관리코드,채널)은 측정창 동안 숨김(데이터 갱신마다 또 뜨는 것 방지)
import datetime as _dt
_SUPPRESS_DAYS = 45
_sup = set()
try:
    _led = dl.read_all(pat, repo)
    if not _led.empty and "ts" in _led.columns:
        _cut = (_dt.date.today() - _dt.timedelta(days=_SUPPRESS_DAYS)).isoformat()
        _r = _led[_led["ts"].astype(str) >= _cut]
        _sup = set(zip(_r["관리코드"].astype(str), _r["채널"].astype(str)))
except Exception:
    _sup = set()
act["_k"] = list(zip(act["관리코드"].astype(str), act["채널"].astype(str)))
_n_sup_shown = int(act["_k"].isin(_sup).sum()) if _sup else 0
act = act[~act["_k"].isin(_sup)].drop(columns="_k").copy()

# ── 필터 ───────────────────────────────────────────────
f1, f2, f3 = st.columns([1.4, 1.4, 2])
with f1:
    chans = sorted(act["채널"].unique())
    pick_ch = st.multiselect("채널", chans, default=[], key="mo_ch",
                             placeholder="전체")
with f2:
    flags = ["🟢", "🟡", "🔴"]
    pick_fl = st.multiselect("플래그", flags, default=flags, key="mo_fl")
with f3:
    q = st.text_input("상품명 / 관리코드 검색", key="mo_q", placeholder="예: 스팸 · 15-04").strip()

v = act.copy()
if pick_ch:
    v = v[v["채널"].isin(pick_ch)]
if pick_fl:
    v = v[v["플래그"].isin(pick_fl)]
if q:
    qn = _nfc(q)
    v = v[v["상품명"].astype(str).map(_nfc).str.contains(qn, case=False, na=False)
          | v["관리코드"].astype(str).str.contains(qn, case=False, na=False)]
v = v.reset_index(drop=True)
_rev = cells.copy()
_rev["월매출"] = (_rev["매출"] / _rev["개월"].clip(lower=1)).round().astype("int64")
v = v.merge(_rev[["관리코드", "채널", "월매출"]], on=["관리코드", "채널"], how="left")
v["월매출"] = v["월매출"].fillna(0).astype("int64")

# ── KPI ────────────────────────────────────────────────
k1, k2, k3, k4 = st.columns(4)
k1.metric("손볼 것", f"{len(act):,}")
k2.metric("🟢 수락", f"{(act['플래그']=='🟢').sum():,}")
k3.metric("🟡 검토", f"{(act['플래그']=='🟡').sum():,}")
k4.metric("🔴 필수", f"{(act['플래그']=='🔴').sum():,}")
st.caption(f"실험큐(관망) {len(queue):,} · 측정중(45일 숨김) {_n_sup_shown:,} · 전체 셀 {len(wl):,} "
           "— 임팩트(월순이익)순. 행 선택 → 아래에서 기록/적용.")

# ── 작업목록 ────────────────────────────────────────────
disp = v.copy()
disp["변화"] = disp["Δ"].map(
    lambda x: f"▲ +{x:.1f}" if x > 0 else (f"▼ {x:.1f}" if x < 0 else "—"))
show = disp[["플래그", "관리코드", "상품명", "채널", "현재마진", "베이스", "권장마진",
             "변화", "월매출", "월순이익", "월볼륨", "액션", "사유"]]


def _color_dir(s):
    # 한국식: 올림 ▲ 빨강 · 내림 ▼ 파랑 · 유지 회색
    if isinstance(s, str) and s.startswith("▲"):
        return "color:#cf222e;font-weight:700"
    if isinstance(s, str) and s.startswith("▼"):
        return "color:#0969da;font-weight:700"
    return "color:#8c959f"


styled = show.style.map(_color_dir, subset=["변화"])
ev = st.dataframe(
    styled, use_container_width=True, hide_index=True,
    on_select="rerun", selection_mode="multi-row",
    column_config={
        "플래그": st.column_config.TextColumn(width="small"),
        "관리코드": st.column_config.TextColumn(width="small"),
        "상품명": st.column_config.TextColumn(width="medium"),
        "현재마진": st.column_config.NumberColumn("현재(%)", format="%.1f"),
        "베이스": st.column_config.NumberColumn("베이스(%)", format="%.1f"),
        "권장마진": st.column_config.NumberColumn("권장(%)", format="%.1f"),
        "변화": st.column_config.TextColumn("권장변화(%p)", width="small"),
        "월매출": st.column_config.NumberColumn("월매출", format="localized"),
        "월순이익": st.column_config.NumberColumn("월순이익", format="localized"),
        "월볼륨": st.column_config.NumberColumn("월볼륨", format="localized"),
        "사유": st.column_config.TextColumn(width="large"),
    },
    height=min(620, 80 + 36 * min(len(show), 15)), key="mo_tbl")

sel_rows = []
try:
    sel_rows = [i for i in ev.selection.rows if 0 <= i < len(v)]  # rerun 후 인덱스 클램프
except Exception:
    sel_rows = []

# 두뇌④ 채널 라벨 → baseline_margin.csv 컬럼 (ESM 라벨만 다름)
_BCOL = {"스마트스토어": "스마트스토어", "ESM(G마켓·옥션)": "ESM", "식봄": "식봄",
         "캐시노트": "캐시노트", "알리": "알리", "쿠팡": "쿠팡",
         "배민상회": "배민상회", "올웨이즈": "올웨이즈"}
_BASELINE_PATH = "reference/baseline_margin.csv"
_APP_REPO = "OURPROJECTDAO/work-automation-app"


def _apply_baseline(updates):
    """updates=[(관리코드, baseline컬럼, 분수)]. baseline_margin.csv 갱신 → cmm 반영.
    return ((applied, miss_row), None) | (None, err). 앱 repo 쓰기 = GITHUB_PAT."""
    import base64 as _b64, io as _io, json as _js, urllib.parse as _uq, urllib.request as _ur
    pat_app = st.secrets.get("GITHUB_PAT", "")
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
    for code, col, frac in updates:
        if col not in bdf.columns:
            continue
        if str(code) not in idx:
            miss_row.append(code)
            continue
        bdf.loc[str(code), col] = (f"{frac:.4f}").rstrip("0").rstrip(".")
        applied.append((code, col))
    buf = _io.StringIO()
    bdf.reset_index().to_csv(buf, index=False, lineterminator="\r\n")
    body = {"message": "data(baseline): 두뇌④ 권장 적용(%d건)" % len(applied),
            "content": _b64.b64encode(("\ufeff" + buf.getvalue()).encode("utf-8")).decode(),
            "sha": meta["sha"]}
    with _ur.urlopen(_ur.Request(url, method="PUT", data=_js.dumps(body).encode(), headers=h)) as r:
        _js.load(r)
    return (applied, miss_row), None


sel_df = v.iloc[sel_rows] if sel_rows else v.iloc[0:0]
chg = st.toggle("기준마진율도 함께 변경 (채널마진모니터 반영)", value=False, key="mo_chg",
                help="켜면 선택 행의 권장마진을 baseline_margin.csv에 기록 → cmm 기준마진율에 반영. 끄면 결정만 기록.")
if chg and len(sel_rows):
    prv = sel_df[["관리코드", "상품명", "채널", "현재마진", "권장마진"]].copy()
    prv["반영"] = prv["채널"].map(lambda c: "✔ cmm" if c in _BCOL else "기록만(매핑외)")
    st.caption("↓ 이 행들의 기준마진율이 baseline_margin.csv에 **현재→권장**으로 기록됩니다. 확인 후 버튼.")
    st.dataframe(prv, hide_index=True, use_container_width=True,
                 column_config={"현재마진": st.column_config.NumberColumn("현재(%)", format="%.1f"),
                                "권장마진": st.column_config.NumberColumn("권장(%)", format="%.1f")})

cA, cB = st.columns([1, 1])
with cA:
    st.download_button("작업목록 CSV", v.to_csv(index=False).encode("utf-8-sig"),
                       file_name="기준마진율_권장.csv", mime="text/csv", key="mo_dl")
with cB:
    _lbl = "📝 기록 + 기준마진율 변경" if chg else "📝 결정 원장에 기록"
    if st.button(f"{_lbl} ({len(sel_rows)}건)", disabled=len(sel_rows) == 0, key="mo_rec"):
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
                ups = [(r["관리코드"], _BCOL[r["채널"]], float(r["권장마진"]) / 100)
                       for _, r in sel_df.iterrows() if r["채널"] in _BCOL]
                skip = sorted({r["채널"] for _, r in sel_df.iterrows() if r["채널"] not in _BCOL})
                res, err = _apply_baseline(ups)
                if err:
                    st.warning("기준마진율 변경 실패: " + err + " (기록은 완료)")
                else:
                    applied, miss_row = res
                    msg += f" 기준마진율 {len(applied)}건 변경 → cmm 반영(최대 10분 내)."
                    if miss_row:
                        msg += f" · baseline 행 없음 {len(miss_row)}건(기록만)."
                    if skip:
                        msg += f" · 매핑외 채널 건너뜀: {', '.join(skip)}."
            st.success(msg + " 다음 사이클에 반응 측정 → 유지/되돌림.")
        except Exception as e:
            st.error(f"실패: {e}")

st.divider()
st.caption(
    "ⓘ **마진=순이익÷정산액**(택배 실배분 차감·기준마진율 시트 정의). **베이스**=순이익 누적 85% proven 채널 "
    "순이익가중평균. **↑/↓ 절반스텝**=베이스까지 거리의 절반만(반응 보고 또 절반/되돌림). **hold-low**=싼데 안 "
    "팔림→안 올림. **🔴**=|Δ|>3%p·신호 약함·hold-low(사람 판단). 권장변화 ▲올림(빨강)·▼내림(파랑)·한국식. "
    "v0 한계: ⑧ 시즌(명절세트) 미보정으로 월순이익 상위에 세트류 부풀려질 수 있음 · 나들=하한 참조(별도) · "
    "'기준마진율도 함께 변경' 켜면 baseline_margin.csv에 써서 cmm 반영(끄면 기록만)·같은 결정은 45일 측정창 동안 숨김. 택배비 2,700원 고정·나들 제외."
)
