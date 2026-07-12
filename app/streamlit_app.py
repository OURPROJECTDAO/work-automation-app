"""업무 자동화 시스템 — Streamlit 진입점."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import streamlit as st
import pandas as pd
# pandas 3.0 + pyarrow arrow-backed 문자열(read_csv)이 특정 네이티브 조합(Cloud 최신 stack)에서
# SIGSEGV → 문자열 저장소를 python 백엔드로 고정해 크래시 경로 회피 (pitfalls 2026-07-13).
try:
    pd.set_option("mode.string_storage", "python")
except Exception:
    pass
from core.ui import inject_css

st.set_page_config(
    page_title="업무 자동화",
    page_icon="⚙️",
    layout="wide",
)

# 전역 UI 폴리시 (전 페이지 적용 — 엔트리는 매 로드마다 먼저 실행)
inject_css()

# 사이드바 브랜드 (네비 위)
with st.sidebar:
    st.markdown(
        '<div class="ui-brand"><span class="logo">⚙️</span> 업무 자동화</div>',
        unsafe_allow_html=True,
    )

_P = Path(__file__).parent / "pages"
_M = _P / "2_기준데이터관리"
_N = _P / "3_연동데이터관리"

pg = st.navigation({
    "": [
        st.Page(_P / "0_지도로드맵.py", title="지도·로드맵", icon="🗺️"),
        st.Page(_P / "0b_데일리대시보드.py", title="데일리 대시보드", icon="📅"),
        st.Page(_P / "1_파일처리.py", title="파일처리", icon="📂"),
        st.Page(_P / "5_송장처리.py", title="송장처리", icon="🏷️"),
    ],
    "기준데이터관리": [
        st.Page(_M / "1_오픈마켓합포도서산간확인.py", title="오픈마켓합포도서산간확인", icon="📋"),
        st.Page(_M / "2_온누리양식_발주서.py",        title="온누리양식_발주서",        icon="💰"),
        st.Page(_M / "3_발주서출력업무.py",           title="발주서출력업무",           icon="🚚"),
        st.Page(_M / "4_천년경영업로드.py",         title="천년경영업로드",         icon="🏪"),
    ],
    "연동데이터관리": [
        st.Page(_N / "1_상품관리.py", title="상품관리", icon="🔗"),
        st.Page(_N / "2_데이터현황.py", title="데이터현황", icon="📚"),
    ],
    "분석·지능": [
        st.Page(_P / "3_대시보드.py", title="대시보드", icon="📊"),
        st.Page(_P / "6_채널마진모니터.py", title="채널마진모니터", icon="💹"),
        st.Page(_P / "7_업로드감시.py", title="업로드감시", icon="📦"),
        st.Page(_P / "8_마진침식.py", title="마진침식", icon="🩸"),
        st.Page(_P / "9_재고지능.py", title="재고지능", icon="🔮"),
        st.Page(_P / "10_가격AB.py", title="가격 A/B", icon="🧪"),
        st.Page(_P / "11_상품360.py", title="상품 360", icon="🪪"),
        st.Page(_P / "12_시장가매칭.py", title="시장가 매칭", icon="🛒"),
        st.Page(_P / "13_기준마진율최적화.py", title="기준마진율 최적화", icon="🎯"),
    ],
})
pg.run()
