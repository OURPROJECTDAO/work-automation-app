"""대시보드 (Phase 4) — 최소 버전.

KPI = 매출만. 구분(식품/음료/선물세트)·거래처·상품·관리코드별 매출 집계.
(차트 콤보·물류량·이익률·거래처그룹·업로더는 추후 추가.)
데이터: work-automation-data 월별 parquet 파티션(core.dashboard.store.load_master).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # repo root

import pandas as pd
import streamlit as st

from core.dashboard import store
from core.dashboard.sales_data import make_classifier

_REF = Path(__file__).parent.parent.parent / "reference"

st.title("📊 영업이익현황 대시보드")
st.caption("매출 기준 · 구분/거래처/상품/관리코드별 집계 (최소 버전 — 차트·물류량·업로더 등은 추후 추가)")


def _data_secret() -> tuple[str, str]:
    """st.secrets에서 (pat, repo). [data] 우선, 없으면 GITHUB_PAT 폴백."""
    repo = "OURPROJECTDAO/work-automation-data"
    try:
        d = st.secrets["data"]
        return d["pat"], d.get("repo", repo)
    except Exception:
        return st.secrets.get("GITHUB_PAT", ""), repo


@st.cache_data(ttl=3600, show_spinner="매출 데이터 불러오는 중...")
def load_sales(pat: str, repo: str) -> pd.DataFrame:
    df = store.load_master(pat, repo)
    if df.empty:
        return df
    # 구분 분류 (1차 분류표 → 2차 product_master 중분류 fallback)
    cls = pd.read_csv(_REF / "logistics_classification.csv", dtype=str, encoding="utf-8-sig")
    pm = pd.read_csv(_REF / "product_master.csv", dtype=str, encoding="utf-8-sig")
    classify = make_classifier(cls, pm)
    df["구분"] = df["관리코드"].map(classify)
    df["연도"] = df["거래일자"].dt.year
    return df


pat, repo = _data_secret()
if not pat:
    st.error("데이터 저장소 접근 정보(st.secrets `[data] pat`)가 설정되지 않았습니다.")
    st.stop()

df = load_sales(pat, repo)
if df.empty:
    st.warning("적재된 매출 데이터가 없습니다. (work-automation-data 파티션 없음)")
    st.stop()

# ── 필터 ──────────────────────────────────────────
years = sorted(df["연도"].unique())
gubuns = ["음료", "식품", "선물세트", "미분류"]
c1, c2 = st.columns(2)
with c1:
    sel_years = st.multiselect("연도", years, default=years)
with c2:
    sel_gubun = st.multiselect("구분", gubuns, default=gubuns)

view = df[df["연도"].isin(sel_years) & df["구분"].isin(sel_gubun)]
if view.empty:
    st.info("선택한 조건에 해당하는 데이터가 없습니다.")
    st.stop()

# ── KPI: 매출 ─────────────────────────────────────
total = view["판매금액"].sum()
k1, k2, k3 = st.columns(3)
k1.metric("총 매출", f"{total/1e8:,.1f}억")
k2.metric("거래 건수", f"{len(view):,}건")
k3.metric("기간", f"{view['거래일자'].min():%Y-%m} ~ {view['거래일자'].max():%Y-%m}")

st.divider()

# ── 구분/거래처/상품/관리코드별 집계 ────────────────
DIMS = {"구분": "구분", "거래처": "상호명", "상품": "상품명", "관리코드": "관리코드"}
dim_label = st.selectbox("집계 기준", list(DIMS), index=0)
dim = DIMS[dim_label]

agg = (view.groupby(dim, observed=True)["판매금액"].sum()
       .sort_values(ascending=False).reset_index())
agg.columns = [dim_label, "매출"]
agg["비중(%)"] = (agg["매출"] / total * 100).round(1)
agg.insert(0, "순위", range(1, len(agg) + 1))

st.subheader(f"{dim_label}별 매출 — 총 {len(agg):,}개")
disp = agg.copy()
disp["매출"] = disp["매출"].map(lambda v: f"{v:,.0f}")
st.dataframe(disp, use_container_width=True, hide_index=True,
             height=min(560, 80 + 36 * min(len(disp), 30)))
st.download_button(
    "표 CSV 내려받기",
    agg.to_csv(index=False).encode("utf-8-sig"),
    file_name=f"매출_{dim_label}별.csv",
    mime="text/csv",
)
