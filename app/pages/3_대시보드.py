"""대시보드 (Phase 4) — 매출 집계 + 데이터 추가(증분 적재).

[대시보드] 탭: 매출 KPI + 구분/거래처/상품/관리코드별 집계.
[데이터 추가] 탭: 영업이익현황 .xlsx 업로드 → 날짜구간 교체로 월 파티션 누적.
(차트·물류량·이익률·거래처그룹은 추후 추가.)
"""
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # repo root

import pandas as pd
import streamlit as st

from core.dashboard import store
from core.dashboard.sales_data import make_classifier, parse_sales

_REF = Path(__file__).parent.parent.parent / "reference"

st.title("📊 영업이익현황 대시보드")


def _data_secret() -> tuple[str, str]:
    """(pat, repo). secrets [data] 우선, 없으면 GITHUB_PAT 폴백."""
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
    cls = pd.read_csv(_REF / "logistics_classification.csv", dtype=str, encoding="utf-8-sig")
    pm = pd.read_csv(_REF / "product_master.csv", dtype=str, encoding="utf-8-sig")
    classify = make_classifier(cls, pm)
    df["구분"] = df["관리코드"].map(classify)
    df["연도"] = df["거래일자"].dt.year
    return df


pat, repo = _data_secret()
tab_dash, tab_add = st.tabs(["📊 대시보드", "➕ 데이터 추가"])

# ── [데이터 추가] 탭 (st.stop 없음 — 코드상 먼저 배치) ────────────
with tab_add:
    st.subheader("영업이익현황 파일 추가")
    st.caption("일/주/월 다운로드한 영업이익현황(.xlsx)을 올리면 월별로 누적됩니다. "
               "같은 날짜구간을 다시 올리면 최신 파일로 교체됩니다(중복 안 쌓임).")
    if not pat:
        st.error("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    else:
        up = st.file_uploader("영업이익현황 .xlsx", type=["xlsx"], key="ingest_up")
        if up is not None:
            try:
                preview = parse_sales(io.BytesIO(up.getvalue()))
            except Exception as e:  # noqa: BLE001
                st.error(f"파일을 읽지 못했습니다: {e}")
                preview = None
            if preview is not None and len(preview):
                dmin, dmax = preview["거래일자"].min(), preview["거래일자"].max()
                months = sorted(preview["거래일자"].dt.strftime("%Y-%m").unique())
                st.info(f"행 **{len(preview):,}** · 기간 **{dmin:%Y-%m-%d} ~ {dmax:%Y-%m-%d}** · "
                        f"영향 월: {', '.join(months)}")
                st.caption("⚠ 위 날짜구간에 해당하는 기존 데이터는 이 파일 내용으로 교체됩니다.")
                if st.button("이 파일로 적재", type="primary", key="ingest_btn"):
                    with st.spinner("적재 중..."):
                        summary = store.ingest(pat, repo, io.BytesIO(up.getvalue()))
                    load_sales.clear()
                    st.success(f"적재 완료 · {summary['rows']:,}행 · 월 {len(summary['months'])}개 갱신 "
                               f"({summary['date_range'][0]} ~ {summary['date_range'][1]})")
                    st.rerun()
            elif preview is not None:
                st.warning("유효한 거래 데이터가 없습니다(합계행만 있거나 빈 파일).")

# ── [대시보드] 탭 ──────────────────────────────────────────────
with tab_dash:
    if not pat:
        st.warning("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    else:
        df = load_sales(pat, repo)
        if df.empty:
            st.info("적재된 매출 데이터가 없습니다. [➕ 데이터 추가] 탭에서 파일을 올려주세요.")
        else:
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
            else:
                total = view["판매금액"].sum()
                k1, k2, k3 = st.columns(3)
                k1.metric("총 매출", f"{total/1e8:,.1f}억")
                k2.metric("거래 건수", f"{len(view):,}건")
                k3.metric("기간", f"{view['거래일자'].min():%Y-%m} ~ {view['거래일자'].max():%Y-%m}")
                st.divider()

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
