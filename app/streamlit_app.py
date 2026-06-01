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

st.title("⚙️ 업무 자동화 시스템")
st.markdown("""
왼쪽 메뉴에서 기능을 선택하세요.

| 메뉴 | 설명 |
|---|---|
| 📂 파일 처리 | 발주 엑셀 업로드 → 워크플로우 실행 → 결과 다운로드 |
| 🗂 기준 데이터 관리 | 도서산간 / 필터링 / 미배송 리스트 조회 · 수정 |
| 📊 대시보드 | 처리 결과 데이터 뷰 (Phase 4 예정) |
""")
