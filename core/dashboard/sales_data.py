"""영업이익현황(거래처별 매출) 데이터 계층 — 프레임워크 무관(core).

흐름: 영업이익현황 xlsx 업로드 → parse_sales(합계행 제외·컬럼선별·타입)
      → split_by_month(월별 분할) → 월별 parquet 파티션 누적
      → date_range_replace(중복 처리: 날짜구간 통째 교체)
대시보드 표시 시: apply_categories(구분 2단 분류 + 박스내품 조인 + 물류량).

설계 근거: decisions/0006-dashboard-data-layer.md
실데이터 검증: logs/2026-06/2026-06-08-phase4-dashboard-datalayer.md
"""
from __future__ import annotations

import io
import unicodedata
from datetime import timezone, timedelta

import pandas as pd

_KST = timezone(timedelta(hours=9))  # Streamlit Cloud = UTC. 날짜는 항상 KST.

# 영업이익현황 23컬럼 중 대시보드에 필요한 것만. 나머지(바코드·단위·메이커·
# 비고 등 빈컬럼/잡컬럼)는 버려 parquet 슬림 + PII 표면 축소.
KEEP_COLS = [
    "거래일자", "상호명", "관리코드", "상품명", "규격",
    "상품분류", "수량", "판매금액", "판매이익",
]
_CAT_COLS = ["상호명", "관리코드", "상품분류", "규격"]  # category dtype로 메모리↓

# 구분 2단 분류 — product_master 중분류(창고 동) → 구분 fallback.
# 1차 = logistics_classification.csv(관리코드→구분: 음료/식품/선물세트).
# 2차 = product_master 중분류. 세제외/잡화 등 비식품은 미분류로(사용자 분류 전까지).
_ZONE_TO_GUBUN = {"음료-B동": "음료", "통조림-C동": "식품"}
_UNCLASSIFIED = "미분류"


def _nfc(s) -> str:
    return unicodedata.normalize("NFC", str(s).strip())


def parse_sales(file_or_buf) -> pd.DataFrame:
    """영업이익현황 xlsx → 정제 DataFrame.

    - calamine 엔진(openpyxl 대비 4배 빠름, 50만행 대비 메모리 안전).
    - 맨끝 합계행(거래일자 NaT) 제외: 안 하면 전 수치 2배.
    - KEEP_COLS만 유지, 거래일자=datetime, 관리코드/상호명 NFC+strip, ym 파생.
    """
    if isinstance(file_or_buf, (bytes, bytearray)):
        file_or_buf = io.BytesIO(file_or_buf)
    df = pd.read_excel(file_or_buf, engine="calamine")
    df = df[df["거래일자"].notna()].copy()  # 합계행 제외
    df = df[[c for c in KEEP_COLS if c in df.columns]].copy()
    df["거래일자"] = pd.to_datetime(df["거래일자"])
    df["관리코드"] = df["관리코드"].map(_nfc)
    df["상호명"] = df["상호명"].astype(str).str.strip()
    df["ym"] = df["거래일자"].dt.strftime("%Y-%m")
    return df.reset_index(drop=True)


def as_category(df: pd.DataFrame) -> pd.DataFrame:
    """상주 메모리 절감용 category dtype 적용(집계 전 1회)."""
    df = df.copy()
    for c in _CAT_COLS:
        if c in df.columns:
            df[c] = df[c].astype("category")
    return df


def split_by_month(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """ym(YYYY-MM) → 그 달 DataFrame. 파티션 단위."""
    return {ym: g.drop(columns="ym").reset_index(drop=True)
            for ym, g in df.groupby("ym", observed=True)}


def date_range_replace(master: pd.DataFrame | None, new: pd.DataFrame) -> pd.DataFrame:
    """중복 처리 = 날짜구간 통째 교체(decisions/0006 #4).

    데이터에 전표/라인 고유ID가 없어 행단위 dedup 불가. "파일이 다루는 날짜구간은
    그 파일이 진실"(최근 다운로드 우선). new의 [min, max] 날짜를 master에서 지우고
    new를 삽입. ERP export가 연속 날짜구간이라는 전제.
    """
    if master is None or len(master) == 0:
        return new.sort_values("거래일자").reset_index(drop=True)
    dmin, dmax = new["거래일자"].min(), new["거래일자"].max()
    keep = master[(master["거래일자"] < dmin) | (master["거래일자"] > dmax)]
    out = pd.concat([keep, new], ignore_index=True)
    return out.sort_values("거래일자").reset_index(drop=True)


def make_classifier(cls_df: pd.DataFrame, pm_df: pd.DataFrame):
    """관리코드 → 구분(음료/식품/선물세트/미분류) 분류 함수 생성.

    cls_df: logistics_classification.csv (컬럼: 관리코드, 구분)
    pm_df : product_master.csv (관리코드, 중분류명 사용)
    """
    cls_map = {_nfc(k): v for k, v in zip(cls_df["관리코드"], cls_df["구분"])}
    mid_map = {_nfc(k): v for k, v in zip(pm_df["관리코드"], pm_df["중분류명"])}

    def classify(code) -> str:
        c = _nfc(code)
        g = cls_map.get(c)
        if g:
            return g
        return _ZONE_TO_GUBUN.get(mid_map.get(c), _UNCLASSIFIED)

    return classify


def make_box_lookup(pm_df: pd.DataFrame):
    """관리코드 → 박스내품(float). 결측/0이면 1.0(물류량=수량 그대로)."""
    box_map = {}
    for k, v in zip(pm_df["관리코드"], pm_df["박스내품"]):
        try:
            f = float(str(v).strip())
        except (ValueError, TypeError):
            f = 0.0
        box_map[_nfc(k)] = f

    def boxq(code) -> float:
        f = box_map.get(_nfc(code), 0.0)
        return f if f > 0 else 1.0

    return boxq


def apply_categories(df: pd.DataFrame, cls_df: pd.DataFrame,
                     pm_df: pd.DataFrame) -> pd.DataFrame:
    """표시용 파생: 구분 + 박스내품 + 물류량(=수량÷박스내품).

    parquet에는 저장 안 함(기준데이터가 매일 갱신되므로 표시 시점에 조인).
    """
    df = df.copy()
    classify = make_classifier(cls_df, pm_df)
    boxq = make_box_lookup(pm_df)
    codes = df["관리코드"]
    df["구분"] = codes.map(classify)
    df["박스내품"] = codes.map(boxq)
    df["물류량"] = df["수량"] / df["박스내품"]
    return df
