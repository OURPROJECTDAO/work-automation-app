"""
오픈마켓합포도서산간확인V7 → Python 재구현.
Phase 1 에서 각 Step 클래스를 채울 것.

버튼 순서(= 처리 단계):
  0   ProcessMergedCells       병합 셀 해제 + 상품명(H) 줄바꿈 합치기
  0.5 SaveSheetToNewFile       정제 송장 별도 출력
  1   CopyDuplicatesToSummary  주소(G) 중복 → 합포확인
  2   SortColumnBDescending    합포확인 주소 정렬
  3   HighlightColumnC         합포확인 색상 구분
  4   FilterAndCopyRows        상품명(H) ↔ 필터링리스트 → 필터링확인
  -   FasterCopyRows           주소(G) ↔ 도서산간리스트 → 지역확인
  -   mbCopyRows               주소(G) ↔ 미배송지리스트 → 미배송지역확인

참조 데이터(reference/):
  dosan_list.csv       도서산간리스트 (10,551행)
  filter_list.csv      필터링리스트   (120건)
  undelivered_list.csv 미배송지리스트 (30건)
"""
from pathlib import Path

from core.base import Workflow, Step, WorkflowContext, normalize_kr
from core.workflows.registry import register

REF_DIR = Path(__file__).parent.parent.parent / "reference"


# ── Phase 1: 아래 Step 클래스들을 구현 ────────────────────────────────

class StepProcessMergedCells(Step):
    """0. 병합 셀 해제 + 상품명 줄바꿈 합치기. (VBA: ProcessMergedCells)"""
    name = "병합셀_정제"
    def run(self, ctx: WorkflowContext) -> None:
        raise NotImplementedError("Phase 1 구현 예정")


class StepCopyDuplicates(Step):
    """1. 주소(G열) 중복 → 합포확인. (VBA: CopyDuplicatesToSummary)"""
    name = "합포_찾기"
    def run(self, ctx: WorkflowContext) -> None:
        raise NotImplementedError("Phase 1 구현 예정")


class StepSortByAddress(Step):
    """2. 합포확인 주소 정렬. (VBA: SortColumnBDescending)"""
    name = "합포_정렬"
    def run(self, ctx: WorkflowContext) -> None:
        raise NotImplementedError("Phase 1 구현 예정")


class StepColorGroups(Step):
    """3. 합포확인 색상 구분. (VBA: HighlightColumnC) — xlsx 색상으로 재현."""
    name = "합포_색상"
    def run(self, ctx: WorkflowContext) -> None:
        raise NotImplementedError("Phase 1 구현 예정")


class StepFilterProducts(Step):
    """4. 상품명 ↔ 필터링리스트 → 필터링확인. (VBA: FilterAndCopyRows)"""
    name = "필터링_확인"
    def run(self, ctx: WorkflowContext) -> None:
        raise NotImplementedError("Phase 1 구현 예정")


class StepDosanCheck(Step):
    """도서산간리스트 주소 매칭 → 지역확인. (VBA: FasterCopyRows)"""
    name = "도서산간_확인"
    def run(self, ctx: WorkflowContext) -> None:
        raise NotImplementedError("Phase 1 구현 예정")


class StepUndeliveredCheck(Step):
    """미배송지리스트 주소 매칭 → 미배송지역확인. (VBA: mbCopyRows)"""
    name = "미배송지_확인"
    def run(self, ctx: WorkflowContext) -> None:
        raise NotImplementedError("Phase 1 구현 예정")


# ── 워크플로우 등록 ────────────────────────────────────────────────────

@register
class OpenmarketMergeWorkflow(Workflow):
    name = "오픈마켓_합포도서산간확인"
    steps = []  # Phase 1: [StepProcessMergedCells(), StepCopyDuplicates(), ...]
    output_sheets = [
        "송장출력", "합포확인", "지역확인", "미배송지역확인", "필터링확인"
    ]
