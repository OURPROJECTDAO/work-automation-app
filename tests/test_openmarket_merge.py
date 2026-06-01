"""
오픈마켓합포도서산간확인 골든 파일 테스트.

사용법 (Phase 1 완료 후):
  1. tests/fixtures/ 에 입력 샘플 xlsx 배치
  2. tests/golden/  에 VBA 출력 결과물(정답) xlsx 배치
  3. pytest tests/ 실행
"""
from pathlib import Path
import tempfile
import pytest
import pandas as pd

FIXTURE_DIR = Path(__file__).parent / "fixtures"
GOLDEN_DIR  = Path(__file__).parent / "golden"


@pytest.mark.skip(reason="Phase 1: 골든 파일 미배치. 준비 후 skip 제거.")
def test_matches_golden():
    from core.workflows.registry import get_workflow

    input_file  = FIXTURE_DIR / "sample_input.xlsx"
    golden_file = GOLDEN_DIR  / "sample_output.xlsx"

    assert input_file.exists(),  "fixtures/sample_input.xlsx 를 배치하세요"
    assert golden_file.exists(), "golden/sample_output.xlsx 를 배치하세요"

    with tempfile.TemporaryDirectory() as tmp:
        result = get_workflow("오픈마켓_합포도서산간확인").run(input_file, Path(tmp))

        golden  = pd.read_excel(golden_file, sheet_name=None)
        output  = pd.read_excel(result,      sheet_name=None)

        for sheet in golden:
            assert sheet in output, f"시트 누락: {sheet}"
            pd.testing.assert_frame_equal(
                output[sheet].reset_index(drop=True),
                golden[sheet].reset_index(drop=True),
                check_like=True,
                obj=f"시트 '{sheet}'",
            )
