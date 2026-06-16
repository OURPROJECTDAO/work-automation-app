"""채널 가격·마진 비교 (두뇌③ A/B v1·서술) — 관리코드 1개 × 채널별 실현마진/볼륨/가격.

intelligence-layer.md §6 ③ 채널 가격 A/B 의 *서술(현재 비교)* 절반. 탄력성 예측은 후속.

분석 단위 = **(관리코드, 채널)** 셀. 한 관리코드를 골라 채널들을 가로로 비교 →
"어디서 많이 파나 / 어디서 비싸게(실수령) 파나".

수치 출처(전부 적재됨·재적재 0):
- 매출·판매이익·매입가·판매량 = 매출자료(대시보드 parquet·정산 진실, §2A). 매입가=판매금액−판매이익(항등식).
- 택배비 = ship_alloc 실측(EA 송장그룹 실배분·합포 ceil(팩/3) 교정·00-12 정합). EA 미경유/2026이전=추정 fallback.
- 노출가 = EasyAdmin 주문 판매가[30](gross). ★단위=EA 판매단위(낱개 아님)·고가 세트/번들은 EA 과소(§7).
- 채널 정규화 = ship_alloc.SANGHO_TO_EA(상호명↔EA 판매처. ESM=G마켓+옥션·배민=상회+대용량·자사몰 제외).

마진율 = 순이익/매입가 (분모=매입가 고정 → 채널 비교 정직, 대시보드 온라인마진 탭과 동일 정의).
낱개이익(원/낱개) = 순이익/판매량(낱개) → 마진율만으론 안 보이는 "비싸게 팔 수 있는 지점".
정산단가(원/낱개,net) = 매출/판매량 → 단위 일관(매출자료 수량=낱개). 실수령 단가.
"""
from __future__ import annotations

import unicodedata

import numpy as np
import pandas as pd

from core.intelligence import ship_alloc

# 온라인 상호명(매출자료) → 친근한 채널 라벨. 외(미등록)는 상호명 그대로.
CHANNEL_LABEL = {
    "오픈마켓- 스마트스토어": "스마트스토어",
    "오픈마켓(ESM/옥션.지마켓.G9)": "ESM(G마켓·옥션)",
    "오픈마켓- (주) 마켓보로": "식봄",
    "오픈마켓 (주) 한국신용데이터": "캐시노트",
    "오픈마켓- 알리": "알리",
    "오픈마켓 쿠팡 (윙배송)": "쿠팡",
    "오픈마켓- (주) 우아한형제들": "배민상회",
    "오픈마켓- 올웨이즈": "올웨이즈",
    "제이티유통": "제이티유통",
}
_BOX_LINE = "00-12"


def _nfc(s) -> str:
    if s is None:
        return ""
    t = str(s)
    if t == "nan":
        return ""
    return unicodedata.normalize("NFC", t).strip()


_LABEL_NFC = {_nfc(k): v for k, v in CHANNEL_LABEL.items()}


def label(sangho) -> str:
    s = _nfc(sangho)
    return _LABEL_NFC.get(s, s)


def months_in_range(d_start, d_end) -> list[str]:
    """[d_start, d_end] 가 걸치는 YYYY-MM 목록."""
    p = pd.period_range(pd.Timestamp(d_start), pd.Timestamp(d_end), freq="M")
    return [str(x) for x in p]


def compute_online_margin(view: pd.DataFrame, ship_rate: dict | None,
                          unit: float, use_actual: bool = True) -> pd.DataFrame:
    """온라인 매출 view(날짜·거래처 필터된)에 택배비 실배분 → _택배/_매입/_순/_실측 부여.

    대시보드 _render_online_margin 과 동일 계산(추정 k = 채널 00-12÷추정송장, 실측 우선).
    view = 매출자료(상호명·관리코드·거래일자·수량·판매금액·판매이익·합포수량·박스내품).
    ★ k(추정 fallback)는 view 전 관리코드로 산출돼야 정확 → 코드 필터 전에 호출할 것.
    """
    box = view["관리코드"].astype(str) == _BOX_LINE
    prod = view[~box].copy()
    hap = prod["합포수량"].fillna(1.0)
    hap = hap.where(hap > 0, 1.0)
    boxn = prod["박스내품"].where(prod["박스내품"] > 0, 1.0)
    prod["_송장"] = prod["수량"] / (hap * boxn)
    실제_s = view.loc[box].groupby("상호명", observed=True)["수량"].sum()
    추정_s = prod.groupby("상호명", observed=True)["_송장"].sum()
    k_s = (실제_s.reindex(추정_s.index).fillna(0.0) / 추정_s.where(추정_s != 0)).fillna(1.0)
    prod["_k"] = prod["상호명"].map(k_s).fillna(1.0)
    prod["_택배_추정"] = prod["_송장"] * unit * prod["_k"]
    if use_actual and ship_rate and ship_rate.get("rate"):
        prod["_택배_실측"] = ship_alloc.attach_actual_ship(prod, ship_rate, float(unit), reconcile=True)
    else:
        prod["_택배_실측"] = float("nan")
    prod["_실측"] = prod["_택배_실측"].notna()
    prod["_택배"] = prod["_택배_실측"].where(prod["_실측"], prod["_택배_추정"])
    prod["_매입"] = prod["판매금액"] - prod["판매이익"]
    prod["_순"] = prod["판매이익"] - prod["_택배"]
    return prod


def channel_breakdown(prod: pd.DataFrame, code: str,
                      ea_price: dict | None = None) -> pd.DataFrame:
    """한 관리코드의 채널별(상호명) 실현마진·볼륨·가격 비교 표.

    prod = compute_online_margin 결과. ea_price = {상호명: 평균노출가}(선택).
    return 컬럼: 채널·마진율(%)·낱개이익·매출·판매량·정산단가·노출가·택배·_상호명.
    매출 desc 정렬.
    """
    c = _nfc(code)
    sub = prod[prod["관리코드"].astype(str).map(_nfc) == c]
    cols = ["채널", "마진율(%)", "낱개이익", "매출", "판매량", "정산단가", "노출가", "택배", "_상호명"]
    if sub.empty:
        return pd.DataFrame(columns=cols)
    g = (sub.groupby("상호명", observed=True)
            .agg(매출=("판매금액", "sum"), 판매이익=("판매이익", "sum"),
                 택배비=("_택배", "sum"), 판매량=("수량", "sum"),
                 _실측건=("_실측", "sum"), _건수=("_실측", "size"))
            .reset_index())
    g["매입가"] = g["매출"] - g["판매이익"]
    g["순이익"] = g["판매이익"] - g["택배비"]
    _mi = g["매입가"].astype("float64")
    g["마진율(%)"] = (g["순이익"] / _mi.where(_mi != 0) * 100).round(2)
    _q = g["판매량"].astype("float64")
    g["낱개이익"] = (g["순이익"] / _q.where(_q != 0)).round(1)
    g["정산단가"] = (g["매출"] / _q.where(_q != 0)).round(0)
    g["택배"] = np.where(g["_실측건"] == g["_건수"], "실측",
                       np.where(g["_실측건"] == 0, "추정", "혼합"))
    g["채널"] = g["상호명"].map(label)
    g["_상호명"] = g["상호명"]
    ep = ea_price or {}
    g["노출가"] = g["상호명"].map(lambda s: ep.get(_nfc(s)))
    g = g.sort_values("매출", ascending=False)
    for col in ("매출", "매입가", "순이익", "택배비"):
        g[col] = g[col].round().astype("int64")
    g["판매량"] = g["판매량"].round().astype("int64")
    return g[cols].reset_index(drop=True)


def build_ea_price_agg(orders: pd.DataFrame) -> pd.DataFrame:
    """EA 주문 → (상호명, 관리코드, ym) 가중 판매가 집계. 노출가 lookup 토대.

    판매처→상호명(ship_alloc.ea_to_sangho), 매핑 외(자사몰) 제외. 판매가>0·수량>0만.
    return [상호명, 관리코드, ym, _pw(=Σ판매가×수량), _q(=Σ수량)].
    """
    cols = ["상호명", "관리코드", "ym", "_pw", "_q"]
    if orders is None or orders.empty:
        return pd.DataFrame(columns=cols)
    rev = ship_alloc.ea_to_sangho()
    df = orders.copy()
    df["상호명"] = df["판매처"].map(_nfc).map(rev)
    df = df[df["상호명"].notna()].copy()
    df["관리코드"] = df["erp관리코드"].map(_nfc)
    가 = pd.to_numeric(df["판매가"], errors="coerce")
    q = pd.to_numeric(df["상품수량"], errors="coerce").fillna(0.0)
    keep = 가.notna() & (가 > 0) & (q > 0)
    df = df[keep].copy()
    if df.empty:
        return pd.DataFrame(columns=cols)
    df["ym"] = pd.to_datetime(df["기준일"]).dt.strftime("%Y-%m")
    df["_pw"] = 가[keep] * q[keep]
    df["_q"] = q[keep]
    g = (df.groupby(["상호명", "관리코드", "ym"], observed=True)
           .agg(_pw=("_pw", "sum"), _q=("_q", "sum")).reset_index())
    return g[cols]


def ea_price_lookup(ea_agg: pd.DataFrame, code: str, months) -> dict:
    """(상호명 → 평균 노출가) for 관리코드 code 와 기간(months) 가중평균."""
    if ea_agg is None or ea_agg.empty:
        return {}
    c = _nfc(code)
    mset = set(months)
    sub = ea_agg[(ea_agg["관리코드"] == c) & (ea_agg["ym"].isin(mset))]
    if sub.empty:
        return {}
    g = sub.groupby("상호명", observed=True).agg(pw=("_pw", "sum"), q=("_q", "sum"))
    return {_nfc(s): round(r.pw / r.q) for s, r in g.iterrows() if r.q > 0}
