"""
워크플로우 베이스 인터페이스.
모든 처리 로직은 Workflow / Step 을 상속해 구현한다.
"""
import unicodedata
from pathlib import Path
from typing import List
import pandas as pd


def normalize_kr(s) -> str:
    """한글 유니코드 NFC 정규화. 주소/상품명 매칭 전 반드시 적용."""
    if not isinstance(s, str):
        return s
    return unicodedata.normalize("NFC", s)


def clean_cell(v):
    """셀 하나의 개행/탭/양끝공백 제거 + NFC 정규화. 문자열이 아니면 그대로."""
    if not isinstance(v, str):
        return v
    s = v.replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()
    return unicodedata.normalize("NFC", s)


def sanitize_ref_df(df: pd.DataFrame, key_col: str | None = None) -> pd.DataFrame:
    """참조 CSV를 GitHub에 저장하기 직전의 셀 위생 처리.

    왜 필요한가 (2026-08-04 사고):
      엑셀에서 셀 범위를 복사해 st.data_editor 에 붙여넣으면 셀 끝 개행(\\n)이
      값에 그대로 딸려 들어온다. CSV 는 그 값을 따옴표로 감싸 보존하므로
      파일 자체는 정상으로 보이지만, 이후 모든 lookup 이 조용히 실패한다
      ('PC005982\\n' != 'PC005982'). 게다가 마지막 열에 개행이 남으면
      재읽기 때 유령 빈 행이 생겨 저장할 때마다 행수가 늘어난다(122→123→...).

    처리:
      - 모든 문자열 셀: 개행/탭 → 공백, 양끝 strip, NFC 정규화
      - 전 컬럼이 빈 행 제거 (유령 행 청소)
      - key_col 지정 시 그 값이 빈 행도 제거
    """
    out = df.copy()
    # pandas 3 는 dtype=str 로 읽은 열이 object 가 아니라 str dtype 이라
    # dtype 검사로 거르면 조용히 통과한다 → 전 컬럼에 걸고 clean_cell 이
    # 비문자열은 그대로 돌려주게 한다.
    for c in out.columns:
        out[c] = out[c].map(clean_cell)

    blank = out.apply(
        lambda r: all(v is None or (isinstance(v, float) and v != v)
                      or str(v).strip() == "" for v in r), axis=1)
    out = out[~blank]

    if key_col and key_col in out.columns:
        out = out[out[key_col].astype(str).str.strip() != ""]

    return out.reset_index(drop=True)


class WorkflowContext:
    """단계 간 데이터를 담는 컨텍스트."""
    def __init__(self, input_path: Path, output_dir: Path):
        self.input_path = input_path
        self.output_dir = output_dir
        self.sheets: dict[str, pd.DataFrame] = {}  # 시트명 → DataFrame
        self.meta: dict = {}                        # 단계 간 공유 메타


class Step:
    """단일 처리 단계 인터페이스. 서브클래스에서 name과 run() 을 구현."""
    name: str = ""

    def run(self, ctx: WorkflowContext) -> None:
        raise NotImplementedError(f"{self.__class__.__name__}.run() 미구현")


class Workflow:
    """워크플로우 베이스.
    서브클래스에서 name, steps, output_sheets 를 정의.
    """
    name: str = ""
    steps: List[Step] = []
    output_sheets: List[str] = []  # 출력 xlsx에 포함할 시트 (순서 유지)

    def run(self, input_path: Path, output_dir: Path) -> Path:
        ctx = WorkflowContext(input_path, output_dir)
        ctx.sheets = self._load(input_path)
        for step in self.steps:
            step.run(ctx)
        return self._save(ctx)

    def _load(self, path: Path) -> dict:
        from core.io_excel import load_sheets
        return load_sheets(path)

    def _save(self, ctx: WorkflowContext) -> Path:
        from core.io_excel import save_sheets
        out = ctx.output_dir / f"{self.name}_결과.xlsx"
        save_sheets(out, {k: ctx.sheets[k] for k in self.output_sheets if k in ctx.sheets})
        return out
