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
st.caption(f"실험큐(관망·저우선) {len(queue):,}건 · 전체 셀 {len(wl):,} — 임팩트(월순이익)순. "
           "표에서 행 선택 → 아래 [결정 원장에 기록].")

# ── 작업목록 ────────────────────────────────────────────
show = v[["플래그", "관리코드", "상품명", "채널", "현재마진", "베이스", "권장마진", "Δ",
          "월매출", "월순이익", "월볼륨", "액션", "사유"]]


def _color_delta(val):
    if val > 0:
        return "color:#1a7f37;font-weight:700"   # ↑ 상향 = 초록
    if val < 0:
        return "color:#cf222e;font-weight:700"   # ↓ 인하 = 빨강
    return "color:#8c959f"                       # 유지/hold = 회색


styled = show.style.map(_color_delta, subset=["Δ"])
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
        "Δ": st.column_config.NumberColumn("Δ(%p)", format="%+.1f"),
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

cA, cB = st.columns([1, 1])
with cA:
    st.download_button("작업목록 CSV", v.to_csv(index=False).encode("utf-8-sig"),
                       file_name="기준마진율_권장.csv", mime="text/csv", key="mo_dl")
with cB:
    disabled = len(sel_rows) == 0
    if st.button(f"📝 선택 {len(sel_rows)}건 결정 원장에 기록", disabled=disabled, key="mo_rec"):
        recs = []
        for i in sel_rows:
            r = v.iloc[i]
            recs.append(dict(
                관리코드=r["관리코드"], 채널=r["채널"], 액션=r["액션"],
                마진_before=float(r["현재마진"]) / 100, 마진_권장=float(r["권장마진"]) / 100,
                베이스=float(r["베이스"]) / 100, 플래그=r["플래그"], 사유=r["사유"],
                측정전_월볼륨=int(r["월볼륨"]), 측정전_월순이익=int(r["월순이익"])))
        try:
            n = dl.append(recs, pat, repo)
            st.success(f"결정 원장에 {n}건 기록(status=pending). 다음 사이클에 반응 측정 → 유지/되돌림.")
        except Exception as e:
            st.error(f"기록 실패: {e}")

st.divider()
st.caption(
    "ⓘ **마진=순이익÷정산액**(택배 실배분 차감·기준마진율 시트 정의). **베이스**=순이익 누적 85% proven 채널 "
    "순이익가중평균. **↑/↓ 절반스텝**=베이스까지 거리의 절반만(반응 보고 또 절반/되돌림). **hold-low**=싼데 안 "
    "팔림→안 올림. **🔴**=|Δ|>3%p·신호 약함·hold-low(사람 판단). "
    "v0 한계: ⑧ 시즌(명절세트) 미보정으로 월순이익 상위에 세트류 부풀려질 수 있음 · 나들=하한 참조(별도) · "
    "cmm 편집창 직접 prefill은 후속(현재는 권장값·결정 기록까지). 택배비 2,700원 고정(실택배비 표준)·나들 제외(데이터 확인용 채널)."
)
