"""파일 처리 페이지: 업로드 → 워크플로우 선택 → 실행 → 다운로드."""
import tempfile
from pathlib import Path
import streamlit as st

# 워크플로우 임포트 (모듈 추가 시 자동 등록)
import core.workflows.openmarket_merge  # noqa: F401
from core.workflows.registry import list_workflows, get_workflow

st.title("📂 파일 처리")

uploaded = st.file_uploader(
    "입력 Excel 파일 업로드", type=["xlsx", "xlsm"],
    help="오픈마켓에서 다운로드한 발주 엑셀 파일을 올려주세요."
)
workflows = list_workflows()
workflow_name = st.selectbox("워크플로우 선택", workflows) if workflows else None

if not workflows:
    st.warning("등록된 워크플로우가 없습니다. core/workflows/ 에 모듈을 추가하세요.")

if uploaded and workflow_name:
    if st.button("▶ 실행", type="primary"):
        with tempfile.TemporaryDirectory() as tmp:
            input_path = Path(tmp) / uploaded.name
            input_path.write_bytes(uploaded.getvalue())
            output_dir = Path(tmp) / "output"
            output_dir.mkdir()

            with st.spinner("처리 중..."):
                try:
                    result = get_workflow(workflow_name).run(input_path, output_dir)
                    st.success("✅ 완료!")
                    st.download_button(
                        label="📥 결과 파일 다운로드",
                        data=result.read_bytes(),
                        file_name=result.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except NotImplementedError as e:
                    st.error(f"미구현 단계: {e}")
                except Exception as e:
                    st.error(f"오류: {e}")
