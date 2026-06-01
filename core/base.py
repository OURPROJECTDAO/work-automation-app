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
