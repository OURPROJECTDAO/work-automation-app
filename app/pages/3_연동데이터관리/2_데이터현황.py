"""연동데이터관리 — data backbone 적재 현황(통합 데이터 관리).

시계열로 누적되는 자료(매출·주문·가격이력·재고 스냅샷·매입현황·발주자료)의 적재
범위·갭을 한눈에. GitHub 디렉토리 목록만 사용(파일 내용 read 0회). 이력 엔진(ADR 0018).
1단계=현황(읽기 전용). 업로드는 다음 단계(직접 적립=매출·주문·가격이력·매입현황).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))  # repo root

import pandas as pd
import streamlit as st

from core.intelligence import coverage as cov_mod


def _data_secret() -> tuple[str, str]:
    repo = "OURPROJECTDAO/work-automation-data"
    try:
        d = st.secrets["data"]
        return d["pat"], d.get("repo", repo)
    except Exception:
        return st.secrets.get("GITHUB_PAT", ""), repo


@st.cache_data(ttl=600, show_spinner="적재 현황 불러오는 중...")
def _load(repo: str, pat: str):
    return cov_mod.coverage(pat, repo)


st.title("📚 데이터 적재 현황")
st.caption("시계열로 누적되는 자료의 적재 범위·갭을 한눈에 봅니다. (work-automation-data)")

pat, repo = _data_secret()
if not pat:
    st.warning("data repo 시크릿([data] 또는 GITHUB_PAT)이 설정되지 않았습니다.")
    st.stop()

if st.button("🔄 새로고침"):
    _load.clear()
    st.rerun()

cov = _load(repo, pat)

ok = [r for r in cov if r["status"] == "ok"]
if ok:
    cols = st.columns(len(ok))
    for col, r in zip(cols, ok):
        with col:
            if r["kind"] == "monthly":
                st.metric(r["label"], f"{r['files']}개월")
                if r["gaps"]:
                    st.caption(f"⚠️ 갭 {len(r['gaps'])}: {', '.join(r['gaps'])}")
                else:
                    st.caption(f"{r['first']} ~ {r['last']}")
            else:
                st.metric(r["label"], f"{r['size_kb']}KB")
                st.caption(r["note"])

rows = []
for r in cov:
    if r["first"]:
        rows.append({"데이터": r["label"], "start": pd.Timestamp(r["first"] + "-01"),
                     "end": pd.Timestamp(cov_mod.next_month(r["last"]) + "-01"), "구분": "적재됨"})
        for g in r["gaps"]:
            rows.append({"데이터": r["label"], "start": pd.Timestamp(g + "-01"),
                         "end": pd.Timestamp(cov_mod.next_month(g) + "-01"), "구분": "갭(누락)"})
if rows:
    import plotly.express as px
    tdf = pd.DataFrame(rows)
    order = [r["label"] for r in cov if r["first"]]
    fig = px.timeline(tdf, x_start="start", x_end="end", y="데이터", color="구분",
                      color_discrete_map={"적재됨": "#378ADD", "갭(누락)": "#E24B4A"},
                      category_orders={"데이터": order})
    fig.update_yaxes(autorange="reversed", title=None)
    fig.update_xaxes(title=None, tickformat="%Y-%m")
    fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10),
                      legend_title_text="", legend=dict(orientation="h", y=1.12))
    st.plotly_chart(fig, use_container_width=True)

UP = {"direct": "직접 업로드", "auto": "자동(부산물)", "planned": "예정"}
tbl = pd.DataFrame([{
    "데이터": r["label"],
    "범위": (f"{r['first']} ~ {r['last']}" if r["first"]
             else ("단일파일" if r["kind"] == "single" and r["status"] == "ok" else "—")),
    "개월/파일": r["files"],
    "갭": (", ".join(r["gaps"]) if r["gaps"] else "—"),
    "크기": f"{r['size_kb']}KB",
    "적립 방식": UP[r["upload"]],
    "설명": r["note"],
} for r in cov])
st.dataframe(tbl, use_container_width=True, hide_index=True)

st.divider()
st.caption("업로드는 다음 단계에서 추가됩니다 — 직접 적립 대상: 매출·주문·가격이력·매입현황. "
           "재고 스냅샷은 상품관리 업로드 시 자동 적립됩니다.")
