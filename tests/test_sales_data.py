"""Phase 4 대시보드 데이터 계층 단위테스트.

검증: 합계행 제외 · 컬럼선별/타입 · 월 분할 · 날짜구간 교체 · 구분 2단 분류 ·
박스내품 결측 fallback · 물류량.
"""
import os

import pandas as pd
import pytest

from core.dashboard.sales_data import (
    parse_sales, split_by_month, date_range_replace,
    make_classifier, make_box_lookup, apply_categories, KEEP_COLS,
)

FIX = os.path.join(os.path.dirname(__file__), "fixtures", "dashboard")


@pytest.fixture
def sales():
    return parse_sales(os.path.join(FIX, "sales_sample.xlsx"))


@pytest.fixture
def cls_df():
    return pd.read_csv(os.path.join(FIX, "classification.csv"), dtype=str)


@pytest.fixture
def pm_df():
    return pd.read_csv(os.path.join(FIX, "product_master.csv"), dtype=str)


def test_parse_drops_total_row(sales):
    # 합성 fixture는 데이터 5행 + 합계행 1행. 합계행 제외 → 5행.
    assert len(sales) == 5
    assert sales["거래일자"].notna().all()


def test_parse_keeps_cols_and_types(sales):
    assert list(sales.columns) == KEEP_COLS + ["ym"]
    assert pd.api.types.is_datetime64_any_dtype(sales["거래일자"])
    assert sales["ym"].iloc[0] == "2026-05"


def test_split_by_month(sales):
    parts = split_by_month(sales)
    assert set(parts) == {"2026-05"}
    assert "ym" not in parts["2026-05"].columns
    assert len(parts["2026-05"]) == 5


def test_date_range_replace_preserves_outside_and_replaces_inside(sales):
    master = sales.copy()
    # 5/10~5/20 구간을 금액 +1000 한 새 파일로 교체
    new = master[(master["거래일자"] >= "2026-05-10")
                 & (master["거래일자"] <= "2026-05-20")].copy()
    in_range = len(new)
    new["판매금액"] = new["판매금액"] + 1000
    merged = date_range_replace(master, new)
    # 같은 데이터 기반 → 총행수 보존
    assert len(merged) == len(master)
    # 구간 행수 × 1000 만큼 증가
    assert merged["판매금액"].sum() - master["판매금액"].sum() == in_range * 1000
    # 구간 밖(5/01, 5/31) 행은 그대로
    outside = merged[(merged["거래일자"] < "2026-05-10")
                     | (merged["거래일자"] > "2026-05-20")]
    assert len(outside) == len(master) - in_range


def test_date_range_replace_empty_master(sales):
    out = date_range_replace(None, sales)
    assert len(out) == len(sales)
    out2 = date_range_replace(sales.iloc[0:0], sales)
    assert len(out2) == len(sales)


def test_classifier_two_tier(cls_df, pm_df):
    classify = make_classifier(cls_df, pm_df)
    assert classify("11-11-11") == "음료"      # 1차 분류표
    assert classify("22-22-22") == "식품"      # 1차 분류표
    assert classify("44-44-44") == "음료"      # 2차 fallback(음료-B동)
    assert classify("33-33-33") == "미분류"    # 잡화-S동 → 미분류
    assert classify("99-99-99") == "미분류"    # 미존재 → 미분류


def test_box_lookup_missing_falls_back_to_one(pm_df):
    boxq = make_box_lookup(pm_df)
    assert boxq("11-11-11") == 20.0
    assert boxq("55-55-55") == 1.0   # 박스내품 0 → 1
    assert boxq("99-99-99") == 1.0   # 미존재 → 1


def test_apply_categories_mulryangg(sales, cls_df, pm_df):
    out = apply_categories(sales, cls_df, pm_df)
    assert {"구분", "박스내품", "물류량"} <= set(out.columns)
    r11 = out[out["관리코드"] == "11-11-11"].iloc[0]
    assert r11["구분"] == "음료" and r11["물류량"] == 100 / 20
    r55 = out[out["관리코드"] == "55-55-55"].iloc[0]
    assert r55["물류량"] == 10 / 1  # 박스내품 결측 → 수량 그대로
