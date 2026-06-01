"""파일 처리 페이지: 발주 파일 업로드 → 워크플로우 실행 → 결과 다운로드."""
import sys, tempfile
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import streamlit as st
import pandas as pd
import core.workflows.openmarket_merge  # noqa: F401  ← @register 트리거
from core.workflows.registry import list_workflows, get_workflow

st.title("📂 파일 처리")
st.caption("마켓플레이스 발주 파일(.xls)을 업로드하면 합포·도서산간·필터링 결과를 한 번에 만들어드립니다.")

col1, col2 = st.columns([2, 1])
with col1:
    uploaded = st.file_uploader(
        "발주 파일 업로드",
        type=["xls", "xlsx", "xlsm"],
        help="스마트스토어·쿠팡·G마켓 등에서 다운로드한 .xls 발주 파일",
    )
with col2:
    workflows = list_workflows()
    workflow_name = st.selectbox("워크플로우", workflows) if workflows else None
    if not workflows:
        st.warning("등록된 워크플로우 없음")

if uploaded and workflow_name:
    if st.button("▶ 실행", type="primary", use_container_width=True):
        with st.spinner("처리 중..."):
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    tmp_path = Path(tmp)
                    input_path = tmp_path / uploaded.name
                    input_path.write_bytes(uploaded.getvalue())
                    output_dir = tmp_path / "output"
                    output_dir.mkdir()
                    result = get_workflow(workflow_name).run(input_path, output_dir)
                    result_bytes = result.read_bytes()
                    result_name  = result.name

                st.success("✅ 처리 완료!")

                # 결과 요약 지표
                import io as _io
                sheets = pd.read_excel(_io.BytesIO(result_bytes), sheet_name=None, dtype=str)
                labels = {"합포확인": "합포", "지역확인": "도서산간",
                          "필터링확인": "필터링", "미배송지역확인": "미배송", "송장출력": "전체 송장"}
                cols = st.columns(len(sheets))
                for col, (sname, df) in zip(cols, sheets.items()):
                    col.metric(labels.get(sname, sname), f"{len(df)}건")

                st.download_button(
                    label="📥 결과 파일 다운로드",
                    data=result_bytes,
                    file_name=result_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"오류: {e}")
                st.exception(e)
else:
    st.info("파일을 업로드하고 워크플로우를 선택하세요.")
