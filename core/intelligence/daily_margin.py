"""당일 마진 점검 (두뇌① 탭D) — 당일 천년경영업로드 output + 송장출력 + master로
대략 실현마진 이상치를 즉시 탐지. intelligence-layer.md §6 ① 당일 렌즈 (사용자 설계 2026-06-15).

마진 3조각 (조인키 = 관리코드):
- 매출 = 천년경영 output **실제기입단가 × 수량** (= net·채널 수수료 적용·매출자료 proximate).
        raw '정산금액'(gross)이 아님 — 쿠팡 정산×0.88·식봄 ×0.93·스마트 정산+배송 → 실제기입단가가 net.
- 원가 = product_master **매입단가(낱개) × 낱개수량**.
- 택배 = **채널 flat 택배단가 × 실제 박스수(송장출력 distinct 송장번호) × 물류량 share**
        (물류량 = 낱개수량 ÷ 박스내품 — 대시보드/P2와 동일 가중). 합포 박스도 distinct 송장번호로 1건.
- 이상 = 역마진(마진<0) OR 마진율 < 채널 baseline − buffer (탭C와 동일 기준).

★ 당일 raw 기반 **조기 트립와이어**. 정산 진실 = 매출자료 월정산(탭C). EasyAdmin/erp 정산은 raw(§2A).
★ PII: 송장출력에서 송장번호는 카운트만(원본 미저장)·수령자/주소/주문번호/상품명 미사용.
"""
from __future__ import annotations

import io
import unicodedata

import pandas as pd

from core.intelligence.margin_erosion import EA_TO_CMM  # 송장 판매처(raw) → cmm 채널


def _nfc(v):
    return unicodedata.normalize("NFC", str(v)).strip() if v is not None and pd.notna(v) else ""


# 천년경영 output 시트 접두(전체/낱개 제거 후) → cmm 채널. 미매핑(멸치·11번가·셀러허브·제이티)=오프라인성 제외.
SHEET_TO_CMM = {
    "스마트스토어": "스마트스토어",
    "ESM": "ESM",
    "식봄": "식봄",
    "배민상회": "배민상회", "대용량": "배민상회",   # 대용량 = 배민대용량장보기
    "올웨이즈": "올웨이즈",
    "캐시노트": "캐시노트",
    "쿠팡": "쿠팡",
    "알리": "알리",
}

DEFAULT_FLAT = 2700  # 채널 flat 택배단가 기본값(사용자 인앱 편집). 문서화된 실택배비 표준.

# 천년경영 전체 시트 컬럼: 0일자 1관리코드 2상품명 3주문수량 4평균단가 5정산금액 6선결제택배비 7실제기입단가
# 천년경영 낱개 시트 컬럼: ...3코드단위주문수량 ...7실제기입단가 8개당수량 9★기입수량 10★낱개단가
_C_CODE, _C_QTY, _C_REAL = 1, 3, 7
_N_PIECES, _N_UNIT = 9, 10


def _wb(file):
    from openpyxl import load_workbook
    if isinstance(file, (bytes, bytearray)):
        src = io.BytesIO(file)
    elif hasattr(file, "read"):
        src = io.BytesIO(file.read())
    else:
        src = file
    return load_workbook(src, data_only=True)   # ★ read_only 금지(천년경영 dim 오독, KB pitfalls)


def parse_invoice_boxes(file) -> dict:
    """송장출력(.xlsx) → {cmm채널: 실제 박스수(distinct 송장번호)}.

    판매처[3] → EA_TO_CMM, 송장번호[9] distinct 카운트. 송장번호 원본 미저장(그룹 카운트만, PII).
    합포(같은 송장번호 여러 라인)도 박스 1건으로 정확히 셈.
    """
    wb = _wb(file)
    ws = wb[wb.sheetnames[0]]
    boxes: dict = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r is None or len(r) < 10:
            continue
        ch = EA_TO_CMM.get(_nfc(r[3]))
        inv = r[9]
        if not ch or inv is None or _nfc(inv) == "":
            continue
        boxes.setdefault(ch, set()).add(_nfc(inv))
    return {ch: len(s) for ch, s in boxes.items()}


def parse_cheonnyeon_sales(file, box_lookup) -> pd.DataFrame:
    """천년경영업로드 output(.xlsx, 27시트) → (채널, 관리코드)별 net 매출·낱개수량·물류량.

    box_lookup = 관리코드(NFC) → 박스내품 dict(결측·0 → 1.0).
    전체: 매출=실제기입단가×주문수량, 낱개수량=주문수량×박스내품.
    낱개: 매출=★낱개단가×★기입수량, 낱개수량=★기입수량.
    물류량 = 낱개수량 ÷ 박스내품.
    """
    wb = _wb(file)
    agg: dict = {}
    for sn in wb.sheetnames:
        base = sn.replace("전체", "").replace("낱개", "")
        ch = SHEET_TO_CMM.get(base)
        if ch is None:
            continue
        nat = sn.endswith("낱개")
        for r in wb[sn].iter_rows(min_row=2, values_only=True):
            if r is None or r[_C_CODE] is None:
                continue
            code = _nfc(r[_C_CODE])
            if not code:
                continue
            bn = float(box_lookup.get(code, 1.0) or 1.0) or 1.0
            if nat:
                pieces = float(r[_N_PIECES] or 0)
                net = float(r[_N_UNIT] or 0) * pieces
            else:
                qty = float(r[_C_QTY] or 0)
                pieces = qty * bn
                net = float(r[_C_REAL] or 0) * qty
            a = agg.setdefault((ch, code), {"매출": 0.0, "낱개수량": 0.0})
            a["매출"] += net
            a["낱개수량"] += pieces
    rows = []
    for (ch, code), a in agg.items():
        bn = float(box_lookup.get(code, 1.0) or 1.0) or 1.0
        rows.append({"채널": ch, "관리코드": code, "매출": a["매출"],
                     "낱개수량": a["낱개수량"], "물류량": a["낱개수량"] / bn})
    return pd.DataFrame(rows, columns=["채널", "관리코드", "매출", "낱개수량", "물류량"])


def compute_daily_margin(sales_df: pd.DataFrame, box_counts: dict, master_buy: dict,
                         name_lookup: dict, baseline_dict: dict,
                         flat_by_channel=None, buffer: float = 0.02) -> pd.DataFrame:
    """당일 (채널, 관리코드) 마진 + 이상 flag.

    master_buy = 관리코드→매입단가(낱개). baseline_dict = 관리코드→{채널: 기준마진}.
    flat_by_channel = 채널→flat 택배단가(dict, 결측 채널 = DEFAULT_FLAT).
    return 전체 행 DataFrame(이상 행만 거르는 건 호출측). 정렬 = 마진 오름차순(가장 손해 먼저).
    """
    if sales_df is None or sales_df.empty:
        return pd.DataFrame(columns=["채널", "관리코드", "상품명", "매출", "낱개수량",
                                     "원가", "택배", "마진", "마진율", "기준마진", "역마진", "미달"])
    flat_by_channel = flat_by_channel or {}
    df = sales_df.copy()
    # 택배 = 채널 실제 박스수 × flat × 물류량 share
    df["택배"] = 0.0
    for ch, g in df.groupby("채널"):
        tot = g["물류량"].sum()
        boxes = box_counts.get(ch, 0)
        flat = float(flat_by_channel.get(ch, DEFAULT_FLAT))
        if tot > 0 and boxes > 0:
            df.loc[g.index, "택배"] = boxes * flat * g["물류량"] / tot
    df["원가"] = [float(master_buy.get(c, 0.0)) * q for c, q in zip(df["관리코드"], df["낱개수량"])]
    df["마진"] = df["매출"] - df["원가"] - df["택배"]
    df["마진율"] = df["마진"] / df["매출"].where(df["매출"] > 0)
    df["상품명"] = df["관리코드"].map(lambda c: name_lookup.get(c, ""))
    df["기준마진"] = [(baseline_dict.get(c, {}) or {}).get(ch)
                   for c, ch in zip(df["관리코드"], df["채널"])]
    df["역마진"] = df["마진"] < 0
    df["미달"] = [(bm is not None) and pd.notna(mr) and (mr < bm - buffer)
                for bm, mr in zip(df["기준마진"], df["마진율"])]
    cols = ["채널", "관리코드", "상품명", "매출", "낱개수량", "원가", "택배",
            "마진", "마진율", "기준마진", "역마진", "미달"]
    return df[cols].sort_values("마진").reset_index(drop=True)
