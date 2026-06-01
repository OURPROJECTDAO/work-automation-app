"""
오픈마켓합포도서산간확인 골든 파일 대조 테스트.

비교 전 두 가지 정규화 적용:
  1. normalize_sheet(): 골든 xlsm의 ,_x000D_\n (OOXML CRLF) → ', '
     HTML-xls 원본은 `, ` 구분자 — 같은 데이터, 표현 형식만 차이.
  2. 송장번호 기준 정렬: VBA xlPinYin ≠ Python Unicode 정렬 → 순서 독립 비교.
"""
from pathlib import Path
import tempfile
import pytest
import pandas as pd

FIXTURE_DIR = Path(__file__).parent / "fixtures"
GOLDEN_DIR  = Path(__file__).parent / "golden"
INPUT_FILE  = FIXTURE_DIR / "input_01.xls"
GOLDEN_FILE = GOLDEN_DIR  / "golden_01.xlsm"

SHEETS = ["송장출력", "합포확인", "지역확인", "미배송지역확인", "필터링확인"]


def normalize_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """골든 xlsx OOXML CRLF 인코딩 → HTML-xls 형식으로 정규화."""
    if "상품명" in df.columns:
        df = df.copy()
        df["상품명"] = df["상품명"].str.replace(",_x000D_\n", ", ", regex=False)
    return df


@pytest.mark.skipif(
    not INPUT_FILE.exists() or not GOLDEN_FILE.exists(),
    reason="fixtures/input_01.xls 또는 golden/golden_01.xlsm 미배치"
)
def test_matches_golden():
    from core.workflows.registry import get_workflow

    with tempfile.TemporaryDirectory() as tmp:
        result = get_workflow("오픈마켓_합포도서산간확인").run(INPUT_FILE, Path(tmp))
        # 임시 디렉토리 삭제 전에 읽기
        output = pd.read_excel(result, sheet_name=None, dtype=str)

    for sname in SHEETS:
        assert sname in output, f"출력 시트 누락: {sname}"
        gold = normalize_sheet(
            pd.read_excel(GOLDEN_FILE, sheet_name=sname, dtype=str)
        )
        out = output[sname]

        assert len(out) == len(gold), (
            f"[{sname}] 행수 불일치: 결과={len(out)} 골든={len(gold)}"
        )

        # 송장번호 기준 정렬 (VBA xlPinYin vs Python Unicode 정렬 차이 회피)
        sort_col = "송장번호"
        pd.testing.assert_frame_equal(
            out.sort_values(sort_col).reset_index(drop=True),
            gold.sort_values(sort_col).reset_index(drop=True),
            check_dtype=False,
            obj=f"시트 '{sname}'",
        )
