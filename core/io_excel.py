"""
Excel/CSV 읽기·쓰기 헬퍼.
한글 인코딩 규칙:
  - xlsx/xlsm: openpyxl → 내부 UTF-8이라 문제없음
  - csv: UTF-8-sig(BOM) → Excel에서 직접 열어도 안 깨짐
"""
from pathlib import Path
import pandas as pd
import openpyxl


def load_sheets(path: Path) -> dict[str, pd.DataFrame]:
    """xlsm/xlsx 의 모든 시트를 DataFrame dict 로 로드. 한글 시트명 OK."""
    wb = openpyxl.load_workbook(path, read_only=True, keep_vba=True, data_only=True)
    result = {}
    for name in wb.sheetnames:
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            result[name] = pd.DataFrame()
        else:
            result[name] = pd.DataFrame(rows[1:], columns=rows[0])
    wb.close()
    return result


def save_sheets(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    """sheets dict 를 xlsx 로 저장. 빈 DataFrame 도 시트로 포함."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)


def load_csv_ref(path: Path) -> pd.DataFrame:
    """UTF-8-sig(BOM) csv 로드. 참조 데이터(도서산간/필터링/미배송) 전용."""
    return pd.read_csv(path, encoding="utf-8-sig", dtype=str).fillna("")
