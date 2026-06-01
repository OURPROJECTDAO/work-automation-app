"""
오픈마켓합포도서산간확인V7 → Python 재구현 (Phase 1)

[VBA 재현 주의사항]
- FasterCopyRows: ListSheet.Range("B1:B") — B1(헤더='주소')부터 읽으므로
  '주소' 자체가 키워드로 포함됨. '(상세주소 없음)' 등에 매칭 → 재현 필수.
- FasterCopyRows: 도서산간아님 예외 로직 없음 (해당 시트는 미사용).
- SortColumnBDescending: xlPinYin 정렬 → Python sort(ascending=False)로 근사.
"""
from pathlib import Path
import pandas as pd
from core.base import Workflow, Step, WorkflowContext, normalize_kr
from core.workflows.registry import register

REF_DIR = Path(__file__).parent.parent.parent / "reference"


class StepLoadInput(Step):
    """0. 입력(HTML-xls) 로드 + 기본 정제."""
    name = "입력_로드"
    def run(self, ctx: WorkflowContext) -> None:
        df = ctx.sheets['송장출력'].copy()
        df['송장번호'] = df['송장번호'].astype(str).str.replace(r'\.0$', '', regex=True)
        df['주소'] = df['주소'].apply(normalize_kr)
        df['상품명'] = df['상품명'].apply(normalize_kr)
        ctx.sheets['송장출력'] = df


class StepCopyDuplicates(Step):
    """1. 주소 중복 → 합포확인 (VBA: CopyDuplicatesToSummary)."""
    name = "합포_찾기"
    def run(self, ctx: WorkflowContext) -> None:
        df = ctx.sheets['송장출력']
        addr_counts = df['주소'].value_counts()
        dup_addrs = set(addr_counts[addr_counts > 1].index)
        df_dup = df[df['주소'].isin(dup_addrs)].copy()
        df_hapo = df_dup[['판매처', '수령자', '주소', '상품명', '송장번호']].rename(
            columns={'수령자': '수취인명'}
        )
        ctx.sheets['합포확인'] = df_hapo


class StepSortByAddress(Step):
    """2. 합포확인 주소 내림차순 정렬 (VBA: SortColumnBDescending, xlPinYin)."""
    name = "합포_정렬"
    def run(self, ctx: WorkflowContext) -> None:
        df = ctx.sheets['합포확인']
        df_sorted = df.sort_values('주소', ascending=False).reset_index(drop=True)
        ctx.sheets['합포확인'] = df_sorted


class StepColorGroups(Step):
    """3. 합포확인 색상 구분 (VBA: HighlightColumnC, color 36↔35).
    DataFrame 내용은 변경 없음 — ctx.meta에 색상 정보만 저장.
    """
    name = "합포_색상"
    COLORS = [36, 35]  # VBA colorIndex: 36=연노랑, 35=연주황

    def run(self, ctx: WorkflowContext) -> None:
        df = ctx.sheets['합포확인']
        colors = {}
        color_idx = 0
        prev_addr = None
        for i, row in df.iterrows():
            addr = row['주소']
            if addr != prev_addr:
                if prev_addr is not None:
                    color_idx = 1 - color_idx
                prev_addr = addr
            colors[i] = self.COLORS[color_idx]
        ctx.meta['hapo_colors'] = colors


class StepFilterProducts(Step):
    """4. 상품명 ↔ 필터링리스트 → 필터링확인 (VBA: FilterAndCopyRows)."""
    name = "필터링_확인"
    def run(self, ctx: WorkflowContext) -> None:
        from core.io_excel import load_csv_ref
        df = ctx.sheets['송장출력']
        fl = load_csv_ref(REF_DIR / 'filter_list.csv')
        keywords = [normalize_kr(k) for k in fl['상품명'].tolist() if k.strip()]
        mask = df['상품명'].apply(
            lambda p: any(kw in normalize_kr(str(p)) for kw in keywords)
        )
        ctx.sheets['필터링확인'] = df[mask].copy().reset_index(drop=True)


class StepDosanCheck(Step):
    """도서산간리스트 주소 매칭 → 지역확인 (VBA: FasterCopyRows).

    VBA 재현 포인트:
      ListData = ListSheet.Range("B1:B" & LastRowList)  ← 헤더(B1='주소')부터 읽음
      → '주소' 키워드가 포함되어 '상세주소 없음', '주소아과' 등에 매칭됨.
      도서산간아님 예외 처리 없음 (VBA에 미구현).
    """
    name = "도서산간_확인"
    def run(self, ctx: WorkflowContext) -> None:
        from core.io_excel import load_csv_ref
        import openpyxl

        df = ctx.sheets['송장출력']

        # VBA 재현: B1(헤더='주소')부터 읽으므로 '주소'도 키워드에 포함
        ds = load_csv_ref(REF_DIR / 'dosan_list.csv')
        # '주소' 헤더 포함 (VBA B1 read 재현)
        ds_kw = ['주소'] + [normalize_kr(k) for k in ds['주소'].tolist() if k.strip()]

        def is_dosan(addr):
            a = normalize_kr(str(addr))
            return any(kw in a for kw in ds_kw)

        ctx.sheets['지역확인'] = df[df['주소'].apply(is_dosan)].copy().reset_index(drop=True)


class StepUndeliveredCheck(Step):
    """미배송지리스트 주소 매칭 → 미배송지역확인 (VBA: mbCopyRows)."""
    name = "미배송지_확인"
    def run(self, ctx: WorkflowContext) -> None:
        from core.io_excel import load_csv_ref
        df = ctx.sheets['송장출력']
        mb = load_csv_ref(REF_DIR / 'undelivered_list.csv')
        mb_kw = [normalize_kr(k) for k in mb['미배송지 주소 리스트'].tolist() if k.strip()]

        def is_mb(addr):
            a = normalize_kr(str(addr))
            return any(kw in a for kw in mb_kw)

        ctx.sheets['미배송지역확인'] = df[df['주소'].apply(is_mb)].copy().reset_index(drop=True)


@register
class OpenmarketMergeWorkflow(Workflow):
    name = "오픈마켓_합포도서산간확인"
    steps = [
        StepLoadInput(),
        StepCopyDuplicates(),
        StepSortByAddress(),
        StepColorGroups(),
        StepFilterProducts(),
        StepDosanCheck(),
        StepUndeliveredCheck(),
    ]
    output_sheets = ["송장출력", "합포확인", "지역확인", "미배송지역확인", "필터링확인"]

    def _load(self, path: Path) -> dict:
        from core.io_excel import detect_and_load_input
        df = detect_and_load_input(path)
        return {'송장출력': df}
