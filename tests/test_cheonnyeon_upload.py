"""천년경영업로드자동화 골든 테스트.

fixtures (PII 없음 — 파이프라인 사용 컬럼만 추출):
  baeju_20260604.xlsx        발주자료 7열 (상품수준)
  baemin_ship_20260604.xlsx  배민 배송비 (Z=관리코드, AL=배송비)
  sss_ship_20260604.xlsx     스스 배송비 슬림·평문 (AJ/AL/AO, r2헤더)
  classification_snapshot.csv / bm_commission.csv / sub_list.csv  기준데이터 스냅샷
golden: golden_260604.xlsx (마켓 전체/낱개 시트만)

기준데이터는 fixture로 명시 주입 → 라이브 reference 편집과 무관하게 결정적.
"""
import datetime
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from core.workflows import cheonnyeon_upload as cy

FX = Path(__file__).parent / "fixtures" / "cheonnyeon"
GOLD = Path(__file__).parent / "golden" / "cheonnyeon" / "golden_260604.xlsx"
RUN_DATE = datetime.date(2026, 6, 4)

FULL_COLS = [(1, "B"), (2, "C"), (3, "D"), (4, "E"), (5, "F"), (6, "G"), (7, "H")]
UNIT_COLS = FULL_COLS + [(8, "I"), (9, "J"), (10, "K")]


def _norm(v):
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return str(v).strip()


def _eq(a, b, tol=0.5):
    a, b = _norm(a), _norm(b)
    if a is None and b is None:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= tol
    return a == b


@pytest.fixture(scope="module")
def result():
    cls_map = cy.load_classification(
        pd.read_csv(FX / "classification_snapshot.csv", encoding="utf-8-sig", dtype=str))
    comm_map = cy.load_commission(
        pd.read_csv(FX / "bm_commission.csv", encoding="utf-8-sig", dtype=str))
    sub_map = cy.load_sub_list(
        pd.read_csv(FX / "sub_list.csv", encoding="utf-8-sig", dtype=str))
    baeju = cy.parse_baeju((FX / "baeju_20260604.xlsx").read_bytes())
    baemin_ws = openpyxl.load_workbook(FX / "baemin_ship_20260604.xlsx", data_only=True).active
    sss_ws = openpyxl.load_workbook(FX / "sss_ship_20260604.xlsx", data_only=True).active
    return cy.process(baeju, baemin_ws, sss_ws, cls_map, comm_map, sub_map, RUN_DATE)


@pytest.fixture(scope="module")
def golden():
    wb = openpyxl.load_workbook(GOLD, data_only=True)
    out = {}
    for name in cy.ALL_SHEETS + cy.UNIT_SHEETS:
        nc = 11 if name in cy.UNIT_SHEETS else 8
        out[name] = [r for r in wb[name].iter_rows(min_row=2, max_col=nc, values_only=True)
                     if r[1] not in (None, "")]
    wb.close()
    return out


@pytest.mark.parametrize("name", cy.ALL_SHEETS)
def test_full_sheets(result, golden, name):
    sheets, _ = result
    mine, gold = sheets[name], golden[name]
    assert len(mine) == len(gold), f"{name} 행수 {len(mine)} != {len(gold)}"
    for i, (m, g) in enumerate(zip(mine, gold), 2):
        for ci, key in FULL_COLS:
            assert _eq(m.get(key), g[ci]), f"{name} r{i} {key}: {m.get(key)!r} != {g[ci]!r}"


@pytest.mark.parametrize("name", cy.UNIT_SHEETS)
def test_unit_sheets(result, golden, name):
    _, units = result
    mine, gold = units[name], golden[name]
    assert len(mine) == len(gold), f"{name} 행수 {len(mine)} != {len(gold)}"
    for i, (m, g) in enumerate(zip(mine, gold), 2):
        for ci, key in UNIT_COLS:
            assert _eq(m.get(key), g[ci]), f"{name} r{i} {key}: {m.get(key)!r} != {g[ci]!r}"


def test_smartstore_shipping_join(result):
    """배송비 조인 검증: 스마트스토어 G = 배송비 AL*0.964 (대표 1건)."""
    sheets, _ = result
    row = next(r for r in sheets["스마트스토어전체"] if r["B"] == "20-90-11")
    assert abs(row["G"] - 2892) <= 1          # 3000 * 0.964
    assert abs(row["H"] - (row["F"] + row["G"]) / row["D"]) <= 0.5


def test_unit_split_formula(result):
    """낱개 J=D*개당수량, K=H*D/J 검증."""
    _, units = result
    for r in units["스마트스토어낱개"]:
        assert abs(r["J"] - r["D"] * r["I"]) <= 1e-6
        assert abs(r["K"] - r["H"] * r["D"] / r["J"]) <= 0.5
