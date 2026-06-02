"""업무 자동화 시스템 — Streamlit 진입점."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import streamlit as st

st.set_page_config(
    page_title="업무 자동화",
    page_icon="⚙️",
    layout="wide",
)

_P = Path(__file__).parent / "pages"
_M = _P / "2_기준데이터관리"
_N = _P / "3_연동데이터관리"

pg = st.navigation({
    "": [
        st.Page(_P / "1_파일처리.py", title="파일처리", icon="📂"),
    ],
    "기준데이터관리": [
        st.Page(_M / "1_오픈마켓합포도서산간확인.py", title="오픈마켓합포도서산간확인", icon="📋"),
        st.Page(_M / "2_온누리양식_발주서.py",        title="온누리양식_발주서",        icon="💰"),
        st.Page(_M / "3_발주서출력업무.py",           title="발주서출력업무",           icon="🚚"),
    ],
    "연동데이터관리": [
        st.Page(_N / "1_상품관리.py", title="상품관리", icon="🔗"),
    ],
    " ": [
        st.Page(_P / "3_대시보드.py", title="대시보드", icon="📊"),
    ],
})
pg.run()
