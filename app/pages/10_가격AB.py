"""가격 A/B (채널 비교) — 관리코드 1개를 골라 채널별 실현마진·볼륨·가격을 가로로 비교.

두뇌③ A/B v1(서술). intelligence-layer.md §6 ③. 탄력성 예측은 후속(이 페이지 아래 섹션 예정).
- 실현마진 = 매출자료(정산) − 매입가(항등식) − 택배비(ship_alloc 실측·합포 교정). 분모=매입가.
- 노출가 = EasyAdmin 판매가(gross·판매단위 기준). 정산단가 = 매출/판매량(net·낱개).
- 채널 정규화 = ship_alloc.SANGHO_TO_EA. 온라인 거래처 = [거래처 그룹] '온라인'.
core/intelligence/channel_compare.py 의 순수함수 사용. page-only(core 무변경, Reboot 불요).
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
from core.intelligence import orders as _orders
from core.intelligence import ship_alloc

_REF = Path(__file__).parent.parent.parent / "reference"

st.title("🧪 가격 A/B — 채널 비교")
st.caption("관리코드 하나를 골라 **채널별 마진율·판매량·매출·가격** 을 비교합니다. "
           "“어디서 많이 파나 / 어디서 비싸게(실수령) 파나” — 채널 가격 전략의 출발점.")


def _nfc(s) -> str:
    return unicodedata.normalize("NFC", str(s)).strip()


def _won(v) -> str:
    try:
        return f"₩{int(round(float(v))):,}"
    except (TypeError, ValueError):
        return "—"


def _data_secret() -> tuple[str, str]:
    repo = "OURPROJECTDAO/work-automation-data"
    try:
        d = st.secrets["data"]
        return d["pat"], d.get("repo", repo)
    except Exception:
        return st.secrets.get("GITHUB_PAT", ""), repo


@st.cache_data(ttl=3600, show_spinner="매출 데이터 불러오는 중...")
def load_sales_min(pat: str, repo: str) -> pd.DataFrame:
    """매출 master + 합포수량·박스내품(택배 추정용). 채널 비교에 필요한 최소 enrich."""
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
def load_ea_agg(pat: str, repo: str) -> pd.DataFrame:
    return cc.build_ea_price_agg(load_orders(pat, repo))


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


def _default_start(dmin, dmax):
    """기본 = 최근 3개월(시작)."""
    s = (pd.Timestamp(dmax) - pd.DateOffset(months=3)).date()
    return max(s, dmin)


pat, repo = _data_secret()
if not pat:
    st.warning("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    st.stop()

df = load_sales_min(pat, repo)
if df.empty:
    st.info("적재된 매출 데이터가 없습니다. [대시보드 ▸ 데이터 추가]에서 영업이익현황을 올려주세요.")
    st.stop()

gmap = load_group_map(pat, repo)
online = sorted({s for s in df["상호명"].astype(str).unique() if gmap.get(_nfc(s)) == "온라인"})
if not online:
    st.info("거래처 그룹에 '온라인'으로 지정된 거래처가 없습니다. "
            "[대시보드 ▸ 거래처 그룹] 탭에서 온라인 채널을 '온라인' 그룹으로 지정해 주세요.")
    st.stop()

# ── 컨트롤 ─────────────────────────────────────────────
dmin, dmax = df["거래일자"].min().date(), df["거래일자"].max().date()
c1, c2, c3 = st.columns([2, 1, 1])
with c1:
    dr = st.date_input("기간", value=(_default_start(dmin, dmax), dmax),
                       min_value=dmin, max_value=dmax, format="YYYY-MM-DD", key="ab_date")
with c2:
    fee_label = st.radio("택배비 단가", ["3,000원", "2,500원"], horizontal=True, key="ab_fee")
with c3:
    use_actual = st.toggle("실측 송장 (EA, 권장)", value=True, key="ab_actual",
                           help="EasyAdmin 송장그룹 실배분 + 00-12 정합. 끄면 추정.")
unit = 2500.0 if "2,500" in fee_label else 3000.0

if isinstance(dr, (list, tuple)):
    d_start, d_end = (dr[0], dr[-1]) if dr else (dmin, dmax)
else:
    d_start = d_end = dr
ts0, ts1 = pd.Timestamp(d_start), pd.Timestamp(d_end) + pd.Timedelta(days=1)
view = df[(df["거래일자"] >= ts0) & (df["거래일자"] < ts1)
          & (df["상호명"].astype(str).isin(online))].copy()
if view.empty:
    st.info("선택한 기간에 온라인 매출 데이터가 없습니다.")
    st.stop()

# ── 관리코드 선택 (검색 → selectbox) ───────────────────
prodv = view[view["관리코드"].astype(str) != "00-12"]
opt = (prodv.groupby("관리코드", observed=True)
       .agg(매출=("판매금액", "sum"),
            상품명=("상품명", lambda s: s.iloc[0] if len(s) else ""),
            채널수=("상호명", "nunique"))
       .reset_index().sort_values("매출", ascending=False))
q = st.text_input("관리코드 / 상품명 검색", key="ab_q", placeholder="예: 코카콜라 · 31-01-04").strip()
if q:
    qn = _nfc(q)
    m = (opt["관리코드"].astype(str).str.contains(qn, case=False, na=False)
         | opt["상품명"].astype(str).map(_nfc).str.contains(qn, case=False, na=False))
    opt = opt[m]
if opt.empty:
    st.info("검색 결과가 없습니다.")
    st.stop()
opt = opt.head(300)
labels = {f"{r['관리코드']} · {r['상품명']}  (채널 {int(r['채널수'])} · 매출 {_won(r['매출'])})": r["관리코드"]
          for _, r in opt.iterrows()}
pick = st.selectbox(f"관리코드 ({len(opt)}개)", list(labels.keys()), key="ab_code")
code = labels[pick]

# ── 채널 비교 계산 ─────────────────────────────────────
ship = load_ship_rate(pat, repo) if use_actual else None
prod = cc.compute_online_margin(view, ship, unit, use_actual=use_actual)
ea = load_ea_agg(pat, repo)
months = cc.months_in_range(d_start, d_end)
ep = cc.ea_price_lookup(ea, code, months)
bd = cc.channel_breakdown(prod, code, ep)
if bd.empty:
    st.info("선택한 관리코드의 온라인 판매 데이터가 없습니다.")
    st.stop()

nm = opt.loc[opt["관리코드"] == code, "상품명"].iloc[0]
st.subheader(f"📦 {code} · {nm}")
tot = bd[["매출", "판매량"]].sum()
k1, k2, k3 = st.columns(3)
k1.metric("총 매출 (온라인)", _won(tot["매출"]))
k2.metric("총 판매량 (낱개)", f"{int(tot['판매량']):,}")
k3.metric("판매 채널 수", f"{len(bd)}")

show = bd.drop(columns=["_상호명"]).copy()
st.dataframe(
    show, use_container_width=True, hide_index=True,
    column_config={
        "채널": st.column_config.TextColumn(width="medium"),
        "마진율(%)": st.column_config.NumberColumn("마진율(%)", format="%.2f",
                                                help="순이익 ÷ 매입가 (택배비 실배분 반영)"),
        "낱개이익": st.column_config.NumberColumn("낱개이익(원)", format="%.1f",
                                              help="순이익 ÷ 판매량(낱개)"),
        "매출": st.column_config.NumberColumn(format="localized", help="정산금액(수수료 차감 후)"),
        "판매량": st.column_config.NumberColumn("판매량(낱개)", format="localized"),
        "정산단가": st.column_config.NumberColumn("정산단가(낱개,net)", format="localized",
                                             help="매출 ÷ 판매량 — 실수령 단가(낱개 단위·일관)"),
        "노출가": st.column_config.NumberColumn("노출가(EA,판매단위)", format="localized",
                                            help="EasyAdmin 판매가(gross). ★판매단위 기준(낱개 아님)·일부 코드만"),
        "택배": st.column_config.TextColumn(width="small", help="실측=EA송장 / 추정=EA미경유 fallback"),
    },
    height=min(560, 80 + 36 * min(len(show), 14)))

cL, cR = st.columns(2)
with cL:
    st.caption("매출 (원)")
    st.bar_chart(bd.set_index("채널")["매출"], height=240)
with cR:
    st.caption("마진율 (%)")
    st.bar_chart(bd.set_index("채널")["마진율(%)"], height=240)

st.download_button("표 CSV 내려받기", show.to_csv(index=False).encode("utf-8-sig"),
                   file_name=f"채널비교_{code}.csv", mime="text/csv", key="ab_dl")

st.caption(
    "ⓘ **마진율 = 순이익÷매입가**(분모 고정 → 채널 비교 정직). "
    "**정산단가**=실수령 낱개단가(매출÷판매량·일관), **노출가**=EA 소비자가(판매단위 기준·낱개 아님·세트/번들은 EA 과소). "
    "**택배**: 실측=EA 송장그룹 실배분(합포 ceil(팩/3) 교정), 추정=EA 미경유 채널 fallback. "
    "탄력성(가격↓→매출↑) 예측은 후속 — 지금은 *현재 채널 비교(서술)* 단계."
)
