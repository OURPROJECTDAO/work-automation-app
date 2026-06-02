"""
발주서 출력 업무 워크플로우.

입력 : 판매처상품매출통계 .xls (HTML), 상품관리 .xlsx (일일 재고)
처리 : Phase1(정제+분류) → Phase2(재고대조+출력)
출력 : 발주자료 아카이브 xlsx, 최종결과물 xlsx (물류팀+품절목록)
참조 : reference/logistics_classification.csv
       reference/unit_list.csv
       reference/spec_master.csv
       reference/product_master.csv  ← 연동데이터관리에서 매일 갱신
"""
import io
import math
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter

_REF = Path(__file__).parent.parent.parent / "reference"

# ───────────────────────────────────────────────
# Reference 로딩
# ───────────────────────────────────────────────

def load_classification() -> pd.DataFrame:
    return pd.read_csv(_REF / "logistics_classification.csv",
                       encoding="utf-8-sig", dtype=str)

def load_unit_list() -> pd.DataFrame:
    return pd.read_csv(_REF / "unit_list.csv",
                       encoding="utf-8-sig", dtype={"관리코드": str, "원코드": str})

def load_spec_master() -> pd.DataFrame:
    return pd.read_csv(_REF / "spec_master.csv",
                       encoding="utf-8-sig", dtype={"관리코드": str})

def load_product_master() -> pd.DataFrame:
    path = _REF / "product_master.csv"
    if not path.exists():
        return pd.DataFrame()
    return pd.read_csv(path, encoding="utf-8-sig", dtype=str)

def get_product_master_updated() -> str:
    meta = _REF / "product_master_updated.txt"
    if meta.exists():
        return meta.read_text(encoding="utf-8").strip()
    return ""

# ───────────────────────────────────────────────
# Phase 1 — 정제 + 분류
# ───────────────────────────────────────────────

_COLS = ["erp관리코드", "어드민옵션", "총수량", "평균단가",
         "정산금액", "판매처그룹", "선결제택배비", "옵션추가항목1"]


def parse_sales_report(file_bytes: bytes) -> pd.DataFrame:
    """HTML-xls 매출통계 파싱. 노이즈 행 제거 후 실제 데이터 반환."""
    html = file_bytes.decode("utf-8").replace("\ufeff", "")
    df = pd.read_html(io.StringIO(html))[0]

    # 헤더 행 찾기
    header_idx = None
    for i, row in df.iterrows():
        if str(row.iloc[0]).strip() == "erp관리코드":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("헤더 행(erp관리코드)을 찾을 수 없습니다")

    data = df.iloc[header_idx + 1:].copy()
    data.columns = _COLS
    data = data.reset_index(drop=True)

    for col in ["총수량", "평균단가", "정산금액", "선결제택배비"]:
        data[col] = pd.to_numeric(data[col], errors="coerce")

    return data


def fill_management_code(df: pd.DataFrame) -> pd.DataFrame:
    """Step 0 : erp관리코드 공백 → 옵션추가항목1에서 채움."""
    df = df.copy()
    mask = df["erp관리코드"].isna() & df["옵션추가항목1"].notna()
    df.loc[mask, "erp관리코드"] = df.loc[mask, "옵션추가항목1"].astype(str)
    return df


def split_merged_cells(df: pd.DataFrame) -> pd.DataFrame:
    """Step 1 : 총수량 NaN 행(병합셀 두 번째 행) → 위 행의 총수량·정산금액 ÷ 2."""
    df = df.copy().reset_index(drop=True)
    i = 1
    while i < len(df):
        if pd.isna(df.at[i, "총수량"]) and not pd.isna(df.at[i - 1, "총수량"]):
            for col in ["총수량", "정산금액"]:
                half = (df.at[i - 1, col] or 0) / 2
                df.at[i - 1, col] = half
                df.at[i, col] = half
            for col in ["판매처그룹", "선결제택배비", "평균단가"]:
                df.at[i, col] = df.at[i - 1, col]
        i += 1
    return df


def enrich_classification(df: pd.DataFrame,
                          cls_df: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """Step 2 : 구분 매핑. 미분류 코드 목록 반환."""
    cls_map = dict(zip(cls_df["관리코드"].astype(str), cls_df["구분"]))
    df = df.copy()
    df["구분"] = df["erp관리코드"].astype(str).map(cls_map)
    unmatched = (
        df[df["구분"].isna() & df["erp관리코드"].notna()]
        [["erp관리코드", "어드민옵션"]]
        .drop_duplicates()
        .to_dict("records")
    )
    return df, unmatched


def deduplicate(df: pd.DataFrame) -> pd.DataFrame:
    """Step 3 : 관리코드 기준 중복 제거 + 총수량·정산금액·선결제비 합산."""
    agg = {
        "어드민옵션": "first",
        "총수량": "sum",
        "평균단가": "first",
        "정산금액": "sum",
        "판매처그룹": "first",
        "선결제택배비": "sum",
        "옵션추가항목1": "first",
        "구분": "first",
    }
    return (
        df.groupby("erp관리코드", sort=False)
        .agg(agg)
        .reset_index()
    )


def enrich_spec(df: pd.DataFrame, spec_df: pd.DataFrame) -> pd.DataFrame:
    """Step 4 : 규격 추가."""
    spec_map = dict(zip(spec_df["관리코드"].astype(str), spec_df["규격"]))
    df = df.copy()
    df["규격"] = df["erp관리코드"].astype(str).map(spec_map).fillna("")
    return df


def run_phase1(sales_bytes: bytes, cls_df=None, spec_df=None):
    """
    Phase 1 전체 실행.
    반환 : (result_df, unmatched, pre_cls_df, archive_df)
      - archive_df : 정제+셀나누기까지만 된 8열 스냅샷 (중복제거 전, 발주자료 아카이브용)
      - unmatched 비어있으면 result_df 완성, pre_cls_df=None
      - unmatched 있으면 result_df=None, pre_cls_df로 GATE A 처리 후 재실행
    """
    if cls_df is None:
        cls_df = load_classification()
    if spec_df is None:
        spec_df = load_spec_master()

    df = parse_sales_report(sales_bytes)
    df = fill_management_code(df)
    df = split_merged_cells(df)
    # 유효 행만 (총수량, erp관리코드 모두 있는 것)
    df = df[df["총수량"].notna() & df["erp관리코드"].notna()].copy()

    # 발주자료 아카이브 스냅샷 (중복제거·구분·규격 전, 원본 8열)
    archive_df = df[_COLS].copy()

    df2, unmatched = enrich_classification(df, cls_df)
    if unmatched:
        return None, unmatched, df, archive_df   # pre_cls_df + archive

    df2 = deduplicate(df2)
    df2 = enrich_spec(df2, spec_df)
    out_cols = ["구분", "규격", "erp관리코드", "어드민옵션", "총수량",
                "평균단가", "정산금액", "판매처그룹", "선결제택배비", "옵션추가항목1"]
    df2 = df2[[c for c in out_cols if c in df2.columns]]
    return df2, [], None, archive_df


def resume_phase1_after_gate(pre_cls_df: pd.DataFrame,
                             cls_df=None, spec_df=None):
    """GATE A 통과 후 Phase 1 재시작 (parse 단계 생략)."""
    if cls_df is None:
        cls_df = load_classification()
    if spec_df is None:
        spec_df = load_spec_master()

    df, unmatched = enrich_classification(pre_cls_df, cls_df)
    if unmatched:
        return None, unmatched, pre_cls_df

    df = deduplicate(df)
    df = enrich_spec(df, spec_df)
    out_cols = ["구분", "규격", "erp관리코드", "어드민옵션", "총수량",
                "평균단가", "정산금액", "판매처그룹", "선결제택배비", "옵션추가항목1"]
    df = df[[c for c in out_cols if c in df.columns]]
    return df, [], None


# ───────────────────────────────────────────────
# Phase 2 — 재고 대조 + 출력
# ───────────────────────────────────────────────

def reconcile_stock(df: pd.DataFrame,
                    pm_df: pd.DataFrame,
                    unit_df: pd.DataFrame):
    """
    재고 대조.
    pm_df : 상품관리 (col4=관리코드, col14=박스재고)
    unit_df : 낱개처리목록
    반환 : (df_with_stock, unmatched_units)
    """
    # 상품관리 맵 (관리코드 → 박스재고)
    stock_map = {}
    if not pm_df.empty:
        codes = pm_df.iloc[:, 4].astype(str).str.strip()
        stocks = pd.to_numeric(pm_df.iloc[:, 14], errors="coerce").fillna(0)
        stock_map = dict(zip(codes, stocks))

    # 낱개 맵 (낱개코드 → {원코드, 입력값})
    unit_map = {}
    for _, r in unit_df.iterrows():
        if pd.notna(r["관리코드"]) and pd.notna(r["원코드"]):
            unit_map[str(r["관리코드"]).strip()] = {
                "원코드": str(r["원코드"]).strip(),
                "입력값": float(r["입력값"]) if pd.notna(r["입력값"]) else 1.0,
            }

    df = df.copy()
    df["_재고박스"] = 0.0
    df["_필요수량"] = pd.to_numeric(df["총수량"], errors="coerce").fillna(0).astype(float)
    df["_낱개"] = False
    df["_원코드미매칭"] = False

    for idx, row in df.iterrows():
        code = str(row["erp관리코드"]).strip()
        qty = float(row["총수량"]) if pd.notna(row["총수량"]) else 0

        if code in unit_map:
            원코드 = unit_map[code]["원코드"]
            배수 = unit_map[code]["입력값"]
            재고 = stock_map.get(원코드, None)
            if 재고 is None:
                df.at[idx, "_원코드미매칭"] = True
                재고 = 0
            df.at[idx, "_재고박스"] = 재고
            df.at[idx, "_필요수량"] = qty * 배수
            df.at[idx, "_낱개"] = True
        else:
            df.at[idx, "_재고박스"] = stock_map.get(code, 0)

    df["재고"] = (df["_재고박스"] - df["_필요수량"]).round(0).astype(int)
    df["총수량표시"] = df.apply(
        lambda r: (f"낱{int(r['총수량'])}" if r["_낱개"]
                   else str(int(r["총수량"])) if pd.notna(r["총수량"])
                   else ""),
        axis=1,
    )

    unmatched_units = (
        df[df["_원코드미매칭"]][["erp관리코드", "어드민옵션"]]
        .drop_duplicates()
        .to_dict("records")
    )
    return df, unmatched_units


def run_phase2(phase1_df: pd.DataFrame, pm_df=None, unit_df=None):
    """
    Phase 2 전체 실행.
    반환 : (logistics_df, unmatched_units, stockout_df)
    """
    if pm_df is None:
        pm_df = load_product_master()
    if unit_df is None:
        unit_df = load_unit_list()

    sections = []
    for 구분 in ["선물세트", "식품", "음료"]:
        sec = phase1_df[phase1_df["구분"] == 구분].copy()
        sec = sec.sort_values("규격", ascending=False)
        sections.append(sec)

    combined = pd.concat(sections, ignore_index=True)
    combined, unmatched_units = reconcile_stock(combined, pm_df, unit_df)

    if unmatched_units:
        return None, unmatched_units, combined

    stockout = (
        combined[combined["재고"] < 0]
        [["erp관리코드", "어드민옵션", "총수량표시", "재고"]]
        .copy()
    )
    stockout.columns = ["관리코드", "상품명", "발주수량", "현재고"]
    stockout = stockout.reset_index(drop=True)
    return combined, [], stockout


# ───────────────────────────────────────────────
# Excel 생성
# ───────────────────────────────────────────────

_THIN = Side(border_style="thin", color="000000")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_RED_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _write_border(ws, min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                             min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = _BORDER


def _merge_consecutive(ws, col: int, start_row: int, end_row: int):
    """동일 값이 연속되는 구간을 병합."""
    vals = [ws.cell(row=r, column=col).value for r in range(start_row, end_row + 1)]
    i = 0
    while i < len(vals):
        j = i + 1
        while j < len(vals) and vals[j] == vals[i]:
            j += 1
        if j - i > 1:
            r_s = start_row + i
            r_e = start_row + j - 1
            ws.merge_cells(
                start_row=r_s, start_column=col,
                end_row=r_e,   end_column=col
            )
            ws.cell(row=r_s, column=col).alignment = _CENTER
        i = j


def generate_archive_xlsx(archive_df: pd.DataFrame) -> bytes:
    """발주자료 아카이브 xlsx (복사붙여넣기 시트, 원본 8열·중복제거 전)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "복사붙여넣기"

    # 실제 워크플로우 헤더 (어드민 옵션 = 띄어쓰기 포함)
    display_header = ["erp관리코드", "어드민 옵션", "총수량", "평균단가",
                      "정산금액", "판매처그룹", "선결제택배비", "옵션추가항목1"]
    ws.append(display_header)
    for _, row in archive_df.iterrows():
        ws.append([row.get(c, "") for c in _COLS])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_result_xlsx(logistics_df: pd.DataFrame,
                         stockout_df: pd.DataFrame) -> bytes:
    """최종결과물 xlsx : 물류팀 + 품절목록."""
    wb = Workbook()

    # ── 물류팀 시트 ──────────────────────────────
    ws = wb.active
    ws.title = "물류팀"

    today = datetime.now().strftime("%Y-%m-%d")

    # Row 1 : 타이틀
    ws.append(["", "멸치+오픈마켓", "", today, "", "재고"])

    current_row = 2
    for 구분 in ["선물세트", "식품", "음료"]:
        section = logistics_df[logistics_df["구분"] == 구분]
        if section.empty:
            continue

        # 섹션 헤더
        ws.append(["구분", "규격", "erp관리코드", "어드민 옵션", "총수량", "재고"])
        current_row += 1
        section_data_start = current_row

        for _, r in section.iterrows():
            재고_val = int(r["재고"])
            ws.append([
                r.get("구분", ""),
                r.get("규격", ""),
                r.get("erp관리코드", ""),
                r.get("어드민옵션", ""),
                r.get("총수량표시", ""),
                재고_val,
            ])
            if 재고_val < 0:
                ws.cell(row=current_row, column=6).fill = _RED_FILL
            # F열 밑줄
            ws.cell(row=current_row, column=6).font = Font(underline="single")
            current_row += 1

        section_data_end = current_row - 1

        if section_data_end >= section_data_start:
            # A열 : 섹션 전체 병합 (구분 동일)
            if section_data_end > section_data_start:
                ws.merge_cells(
                    start_row=section_data_start, start_column=1,
                    end_row=section_data_end,     end_column=1
                )
                ws.cell(row=section_data_start, column=1).alignment = _CENTER
            else:
                ws.cell(row=section_data_start, column=1).alignment = _CENTER

            # B열 : 연속 동일 규격 병합
            _merge_consecutive(ws, col=2,
                               start_row=section_data_start,
                               end_row=section_data_end)

    # 전체 테두리 + C열 숨김
    last_row = ws.max_row
    _write_border(ws, 1, last_row, 1, 6)
    ws.column_dimensions["C"].hidden = True

    # 컬럼 너비
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 8

    # ── 품절목록 시트 ─────────────────────────────
    ws2 = wb.create_sheet("품절목록")
    ws2.append(["관리코드", "상품명", "발주수량", "현재고"])
    for _, r in stockout_df.iterrows():
        ws2.append([r["관리코드"], r["상품명"], r["발주수량"], int(r["현재고"])])

    _write_border(ws2, 1, ws2.max_row, 1, 4)
    ws2.column_dimensions["B"].width = 35
    ws2.column_dimensions["C"].width = 10
    ws2.column_dimensions["D"].width = 10

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
