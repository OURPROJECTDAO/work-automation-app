"""발주서출력업무(logistics_order) 골든 대조 테스트.

입력(매출통계+상품관리+기준데이터) → Phase1/2 → 물류팀·품절목록을
골든 결과물(물류팀프로그램v5_2 최종결과물.xlsm 추출)과 1:1 대조.
"""
import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))
import core.workflows.logistics_order as lo

FIX = Path(__file__).parent / "fixtures" / "logistics"


@pytest.fixture(scope="module")
def pipeline_result():
    sales = (FIX / "sales_input.xls").read_bytes()
    cls   = pd.read_csv(FIX / "classification.csv", encoding="utf-8-sig", dtype=str)
    spec  = pd.read_csv(FIX / "spec_master.csv",    encoding="utf-8-sig", dtype={"관리코드": str})
    unit  = pd.read_csv(FIX / "unit_list.csv",      encoding="utf-8-sig",
                        dtype={"관리코드": str, "원코드": str})
    pm    = pd.read_csv(FIX / "product_master.csv", encoding="utf-8-sig", dtype=str)

    p1, unmatched_cls, _, archive = lo.run_phase1(sales, cls_df=cls, spec_df=spec)
    assert not unmatched_cls, f"미분류 코드 발생: {unmatched_cls}"

    p2, unmatched_unit, stockout = lo.run_phase2(p1, pm_df=pm, unit_df=unit)
    assert not unmatched_unit, f"낱개 미매칭 발생: {unmatched_unit}"

    return {"phase1": p1, "logistics": p2, "stockout": stockout, "archive": archive}


def test_archive_is_8col_prededup(pipeline_result):
    """발주자료 아카이브 = 원본 8열, 중복제거 전."""
    archive = pipeline_result["archive"]
    assert list(archive.columns) == lo._COLS
    # 중복제거 전이므로 phase1(중복제거 후)보다 행 수가 많아야 함
    assert len(archive) >= len(pipeline_result["phase1"])


def test_logistics_matches_golden(pipeline_result):
    """물류팀: erp관리코드별 (총수량표시, 재고) 골든 일치."""
    p2 = pipeline_result["logistics"]
    golden = pd.read_csv(FIX / "golden_물류팀.csv", encoding="utf-8-sig",
                         dtype={"erp관리코드": str, "총수량": str})

    result_map = {
        str(r["erp관리코드"]).strip(): (str(r["총수량표시"]), int(r["재고"]))
        for _, r in p2.iterrows()
    }
    assert len(p2) == len(golden), f"행수 불일치: 결과 {len(p2)} vs 골든 {len(golden)}"

    for _, g in golden.iterrows():
        code = str(g["erp관리코드"]).strip()
        assert code in result_map, f"골든 코드 누락: {code}"
        got = result_map[code]
        exp = (str(g["총수량"]), int(g["재고"]))
        assert got == exp, f"{code}: 결과 {got} vs 골든 {exp}"


def test_stockout_matches_golden(pipeline_result):
    """품절목록: 관리코드별 (발주수량, 현재고) 골든 일치."""
    so = pipeline_result["stockout"]
    golden = pd.read_csv(FIX / "golden_품절목록.csv", encoding="utf-8-sig",
                         dtype={"관리코드": str, "발주수량": str})

    so_map = {
        str(r["관리코드"]).strip(): (str(r["발주수량"]), int(r["현재고"]))
        for _, r in so.iterrows()
    }
    assert len(so) == len(golden), f"품절 행수 불일치: 결과 {len(so)} vs 골든 {len(golden)}"

    for _, g in golden.iterrows():
        code = str(g["관리코드"]).strip()
        assert code in so_map, f"골든 품절코드 누락: {code}"
        got = so_map[code]
        exp = (str(g["발주수량"]), int(g["현재고"]))
        assert got == exp, f"{code}: 결과 {got} vs 골든 {exp}"


def test_split_merged_cells():
    """셀나누기: 총수량 NaN 행 → 위 행과 ÷2 분할."""
    df = pd.DataFrame({
        "erp관리코드": ["A", "B"],
        "어드민옵션": ["x", "y"],
        "총수량": [2.0, None],
        "평균단가": [100, 100],
        "정산금액": [400.0, None],
        "판매처그룹": ["ESM", None],
        "선결제택배비": [0, None],
        "옵션추가항목1": [None, None],
    })
    out = lo.split_merged_cells(df)
    assert out.loc[0, "총수량"] == 1.0
    assert out.loc[1, "총수량"] == 1.0
    assert out.loc[0, "정산금액"] == 200.0
    assert out.loc[1, "정산금액"] == 200.0
    assert out.loc[1, "판매처그룹"] == "ESM"
