"""invoice_fill 워크플로우 테스트 (합성 데이터, PII 없음, 자족형)."""
import io
import xlrd
import xlwt
import pytest

from core.workflows.invoice_fill import (
    parse_template_xls, build_master_lookup, vlookup_fill,
    find_consolidation_candidates, apply_decisions, finalize, write_template_xls,
)

TEMPLATE_HEADER = ["상품주문ID", "상품주문번호", "주문번호", "결제일시", "주문일시", "구매자명",
                   "구매자연락처", "수취인명(받는사람)", "수취인연락처", "상품명", "수량",
                   "배송지", "배송방법", "택배사", "송장번호"]
GUIDE = ["수정/삭제 불가"] * 12 + ["택배배송", "택배사 입력", "송장번호 입력"]

ADDR_X, ADDR_Y, ADDR_Z, ADDR_W = (f"가상시 가상구 가상로 {n}" for n in (1, 2, 3, 4))


def _make_template_xls(rows):
    """rows: (상품주문번호, 수취인, 상품명, 배송지). 송장 빈칸 .xls 바이트."""
    wb = xlwt.Workbook(encoding="utf-8")
    sh = wb.add_sheet("송장번호 일괄입력 템플릿")
    for c, v in enumerate(TEMPLATE_HEADER):
        sh.write(0, c, v)
    for c, v in enumerate(GUIDE):
        sh.write(1, c, v)
    for i, (ordno, recv, prod, addr) in enumerate(rows):
        r = i + 2
        sh.write(r, 0, 1000 + i)          # 상품주문ID (숫자)
        sh.write(r, 1, ordno)             # 상품주문번호 (키)
        sh.write(r, 2, "M" + ordno)       # 주문번호 (미끼)
        sh.write(r, 7, recv); sh.write(r, 9, prod); sh.write(r, 11, addr)
        sh.write(r, 12, "택배배송"); sh.write(r, 13, "한진택배"); sh.write(r, 14, "")
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def _master_row(order, recv, addr, prod, song):
    return {"상태": "송장", "관리번호": "0", "발주일": "2026-06-04", "판매처": "식봄(마켓보로)",
            "주문번호": order, "수령자": recv, "주소": addr, "상품명": prod,
            "택배사": "한진택배", "송장번호": song}


@pytest.fixture
def scenario():
    before = _make_template_xls([
        ("A001", "수취갑", "콜라", ADDR_X),
        ("A002", "수취갑", "사이다", ADDR_X),
        ("A003", "수취갑", "환타", ADDR_X),   # N/A → 합포후보(박스 2개)
        ("B001", "수취을", "물", ADDR_Y),
        ("B002", "수취을", "차", ADDR_Y),     # N/A → 합포후보(박스 1개)
        ("C001", "수취병", "빵", ADDR_Z),     # N/A 독립 → 삭제
        ("D001", "수취정", "김", ADDR_W),
    ])
    master = [
        _master_row("A001", "수취갑", ADDR_X, "콜라", "T100"),
        _master_row("A001", "수취갑", ADDR_X, "콜라", "T101"),   # 분할 → 첫매칭만
        _master_row("A002", "수취갑", ADDR_X, "사이다", "T200"),
        _master_row("B001", "수취을", ADDR_Y, "물", "T300"),
        _master_row("D001", "수취정", ADDR_W, "김", "T400"),
        _master_row("ZZZ", "무관", "어딘가", "기타", "T999"),
    ]
    return before, master


def test_first_match(scenario):
    _, master = scenario
    lk = build_master_lookup(master, "주문번호")
    assert lk["A001"] == ("한진택배", "T100")   # 분할배송 첫 박스만


def test_vlookup_and_candidates(scenario):
    before, master = scenario
    parsed = parse_template_xls(before)
    lk = build_master_lookup(master, "주문번호")
    rows = vlookup_fill(parsed["rows"], lk, "식봄")
    assert [r["_status"] for r in rows] == \
        ["matched", "matched", "na", "matched", "na", "na", "matched"]
    cands, indep = find_consolidation_candidates(rows)
    cmap = {c["na_row"]["상품주문번호"]: c for c in cands}
    assert set(cmap) == {"A003", "B002"}
    assert {b["송장"] for b in cmap["A003"]["boxes"]} == {"T100", "T200"}  # 다중박스
    assert len(cmap["B002"]["boxes"]) == 1                                  # 단일박스
    assert [rows[i]["상품주문번호"] for i in indep] == ["C001"]            # 독립


def test_decisions_finalize_and_xls(scenario):
    before, master = scenario
    parsed = parse_template_xls(before)
    lk = build_master_lookup(master, "주문번호")
    rows = vlookup_fill(parsed["rows"], lk, "식봄")
    cands, _ = find_consolidation_candidates(rows)
    cmap = {c["na_row"]["상품주문번호"]: c for c in cands}
    apply_decisions(rows, {cmap["A003"]["na_index"]: "T200",
                           cmap["B002"]["na_index"]: "T300"})  # C001 미결정→유지
    keep, na_count, na_rows = finalize(rows)
    assert na_count == 1 and na_rows[0]["상품주문번호"] == "C001"
    assert [rows[i]["상품주문번호"] for i in keep] == \
        ["A001", "A002", "A003", "B001", "B002", "D001"]
    # 원본 .xls 양식 라운드트립
    out = write_template_xls(parsed, rows, keep)
    g = xlrd.open_workbook(file_contents=out).sheet_by_index(0)
    assert g.name == "송장번호 일괄입력 템플릿"
    assert [g.cell_value(0, c) for c in range(g.ncols)] == TEMPLATE_HEADER
    assert {g.cell_value(r, 1): g.cell_value(r, 14) for r in range(2, g.nrows)} == \
        {"A001": "T100", "A002": "T200", "A003": "T200",
         "B001": "T300", "B002": "T300", "D001": "T400"}
    assert g.cell_type(2, 0) == xlrd.XL_CELL_NUMBER   # 상품주문ID 숫자 보존
