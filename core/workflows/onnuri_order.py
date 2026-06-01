"""
온누리양식_발주서 워크플로우.

입력: 발주서 xlsx (관리코드, 총 주문 수량 등 포함)
처리: SKU 참조 → 합계:판매가(부가세 포함) 계산
출력: 원본파일명(확인).xlsx

수식: 합계 = 공급가(VAT포함) × 수량 + ceil(수량 / 최대합포수량) × 배송비
참조: reference/sku_list.csv
"""
import math
import shutil
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

from core.base import Workflow, Step, WorkflowContext
from core.workflows.registry import register

_REF = Path(__file__).parent.parent.parent / "reference"
_COL_TOTAL = "합계 : 판매가(부가세 포함)"
_COL_CODE = "관리코드"
_COL_QTY = "총 주문 수량"


class LoadSKU(Step):
    """SKU 참조 CSV 로드 → ctx.meta['sku'] (관리코드 인덱스)."""
    name = "load_sku"

    def run(self, ctx: WorkflowContext) -> None:
        sku = pd.read_csv(
            _REF / "sku_list.csv",
            encoding="utf-8-sig",
            dtype={"관리코드": str, "원코드": str},
        )
        sku.drop_duplicates("관리코드", keep="first", inplace=True)
        ctx.meta["sku"] = sku.set_index("관리코드")


class CalcTotal(Step):
    """합계:판매가 = 공급가 × 수량 + ceil(수량/최대합포) × 배송비."""
    name = "calc_total"

    def run(self, ctx: WorkflowContext) -> None:
        df = ctx.sheets["발주서"]
        sku = ctx.meta["sku"]

        def _calc(row):
            code = str(row[_COL_CODE]) if row[_COL_CODE] is not None else ""
            if not code or code not in sku.index:
                return None
            qty = int(row[_COL_QTY])
            s = sku.loc[code]
            price = int(s["공급가(VAT 포함)"])
            max_bundle = int(s["배송비 부과 규칙 (규격 기준)"])
            ship = int(s["배송비"])
            n = math.ceil(qty / max_bundle)
            return price * qty + n * ship

        df[_COL_TOTAL] = df.apply(_calc, axis=1)


@register
class OnnuriOrderWorkflow(Workflow):
    """온누리양식 발주서 처리 워크플로우."""

    name = "온누리양식_발주서"
    steps = [LoadSKU(), CalcTotal()]
    output_sheets = ["발주서"]

    def _load(self, path: Path) -> dict:
        """openpyxl로 읽어 원본 셀 값(문자열 우편번호 등) 보존."""
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows:
            return {"발주서": pd.DataFrame()}
        return {"발주서": pd.DataFrame(rows[1:], columns=rows[0])}

    def _save(self, ctx: WorkflowContext) -> Path:
        """원본 xlsx 복사 후 합계 컬럼만 덮어쓰기 → 원본파일명(확인).xlsx."""
        stem = ctx.input_path.stem
        out = ctx.output_dir / f"{stem}(확인).xlsx"
        shutil.copy2(ctx.input_path, out)

        df = ctx.sheets["발주서"]
        wb = load_workbook(out)
        ws = wb.active

        header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        try:
            col_idx = header.index(_COL_TOTAL) + 1  # 1-based
        except ValueError:
            wb.close()
            raise ValueError(f"'{_COL_TOTAL}' 컬럼을 발주서 시트에서 찾지 못했습니다.")

        for i, val in enumerate(df[_COL_TOTAL].tolist()):
            if val is not None:
                cell = ws.cell(row=i + 2, column=col_idx)
                cell.value = int(val)
                cell.alignment = Alignment(horizontal="right")

        wb.save(out)
        wb.close()
        return out
