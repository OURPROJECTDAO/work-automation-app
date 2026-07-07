"""연동데이터관리 — data backbone 적재 현황(통합 데이터 관리).

시계열로 누적되는 자료(매출·주문·가격이력·재고 스냅샷·매입현황·발주자료)의 적재
범위·갭을 한눈에. GitHub 디렉토리 목록만 사용(파일 내용 read 0회). 이력 엔진(ADR 0018).
현황(읽기 전용) + 직접 적립 업로드(매출·주문·가격이력·매입현황).
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
                    _f = r.get("first_day") or r["first"]
                    _l = r.get("last_day") or r["last"]
                    st.caption(f"{_f} ~ {_l}")
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
    st.plotly_chart(fig, width="stretch")

UP = {"direct": "직접 업로드", "auto": "자동(부산물)", "planned": "예정"}
tbl = pd.DataFrame([{
    "데이터": r["label"],
    "범위": (f"{r.get('first_day') or r['first']} ~ {r.get('last_day') or r['last']}" if r["first"]
             else ("단일파일" if r["kind"] == "single" and r["status"] == "ok" else "—")),
    "개월/파일": r["files"],
    "갭": (", ".join(r["gaps"]) if r["gaps"] else "—"),
    "크기": f"{r['size_kb']}KB",
    "적립 방식": UP[r["upload"]],
    "설명": r["note"],
} for r in cov])
st.dataframe(tbl, width="stretch", hide_index=True)

st.divider()
st.subheader("📤 직접 적립")
st.caption("매출·주문·가격이력·매입현황을 여기서 업로드하면 바로 적재됩니다. "
           "재고 스냅샷은 상품관리 업로드 시 자동 적립됩니다.")

_SOURCES = {
    "매출 — 천년경영 영업이익현황": "sales",
    "주문 — EasyAdmin 확장주문검색": "orders",
    "가격이력 — 상품수정삭제로그": "price",
    "매입현황 — 유형별매입현황": "purchases",
}
_EXT = {"sales": ["xlsx"], "orders": ["xls"], "price": ["xlsx"], "purchases": ["xlsx"]}
_HELP = {
    "sales": "Exp______영업이익현황_YYYYMMDD-YYYYMMDD.xlsx",
    "orders": "확장주문검색_YYYYMMDDHHMMSS_.xls (HTML 위장 파일 — 정상)",
    "price": "Exp______상품수정삭제로그_YYYYMMDD-YYYYMMDD.xlsx",
    "purchases": "Exp______유형별매입현황_YYYYMMDD-YYYYMMDD.xlsx (통파일·개별 둘 다 가능)",
}
_DATE_COL = {"sales": "거래일자", "orders": "기준일", "price": "수정일자", "purchases": "기준일"}


def _salt() -> str:
    try:
        return st.secrets["data"].get("customer_key_salt", "")
    except Exception:
        return ""


choice = st.selectbox("자료 종류", list(_SOURCES.keys()))
kind = _SOURCES[choice]

up = st.file_uploader(f"{choice} 파일 업로드", type=_EXT[kind], help=_HELP[kind], key=f"up_{kind}")

if up:
    raw = up.read()
    try:
        if kind == "sales":
            from core.dashboard import sales_data as _mod, store as _store
            new = _mod.parse_sales(raw)
        elif kind == "orders":
            from core.intelligence import orders as _mod
            salt = _salt()
            if not salt:
                st.warning("⚠️ customer_key_salt 시크릿이 없어 고객키·합포박스키는 빈값으로 적재됩니다.")
            new = _mod.parse_orders(raw, salt=salt or None)
        elif kind == "price":
            from core.intelligence import price_history as _mod
            new = _mod.parse_price_log(raw)
        else:  # purchases
            from core.intelligence import purchases as _mod
            new = _mod.parse_purchases(raw)

        st.success(f"✅ {len(new)}행 인식됨")
        dc = _DATE_COL[kind]
        if len(new) and dc in new.columns:
            st.caption(f"기간: {new[dc].min()} ~ {new[dc].max()}")
        st.dataframe(new.head(8), width="stretch", height=200)

        if st.button("📤 적재", type="primary", width="stretch", key=f"ingest_{kind}"):
            with st.spinner("GitHub에 적재 중..."):
                if kind == "sales":
                    result = _store.ingest(pat, repo, raw)
                else:
                    result = _mod.ingest(new, pat, repo)
            st.success(f"✅ 적재 완료: {result}")
            _load.clear()
            st.rerun()
    except KeyError as e:
        st.error(f"필수 컬럼을 찾을 수 없습니다: {e}. 헤더가 깨졌을 수 있어요 — 원본 파일 1행을 확인해주세요.")
    except Exception as e:
        st.error(f"처리 오류: {e}")
        st.exception(e)
