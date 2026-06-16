"""송장 그룹 → 박스당 실택배비 상품 배분 (P2 실측 마진). intelligence-layer.md §5 Phase 2.

EasyAdmin 주문(orders/easyadmin_YYYY-MM.parquet)의 **송장그룹(=박스)** 을 구성 상품 라인에
물류량(수량÷박스내품) 비중으로 배분 → (상호명, 관리코드, 월) **실제 박스수**.
이를 같은 키의 **매출 낱개수**(매출자료 수량)로 나눠 **택배강도(박스/낱개)** 산출.
대시보드 온라인 상품마진 탭의 추정송장·k 를 이 실측으로 대체.

핵심 단위 정정(2026-06-15 실측):
- 매출자료 '수량' = 낱개(pieces), EA '상품수량' = 판매단위(pack/box). 배율 ≈ 박스내품.
- 따라서 강도 분모 = **매출 낱개수**(EA 수량 아님). 분자 = EA 실제 박스수(송장그룹 배분).
- rate = EA박스 ÷ 매출낱개 → 택배_행 = 매출수량 × rate × 단가. 단위 일치·기간 필터 안전.

설계 (사용자 확정 2026-06-15):
- 매출자료(정산 진실)는 수량·판매금액·이익 그대로 유지, EA는 박스/낱개 비율만 제공
  → EA≠매출 timing/취소 괴리(intelligence-layer §7)가 절대 택배비에 새지 않음.
- 상호명↔EA판매처: ESM=G마켓+옥션, 배민=배민상회+배민대용량, 나머지 1:1.
- EA 미경유 상호명/월 = rate 없음 → 페이지가 기존 추정(추정송장·k) fallback.
- 박스 98.6% 단일라인: 그 박스=그 상품 통째. 합포(관리코드≥2)는 4건뿐 → 물류량 가중.
- rate 상한 1.0(낱개 1개를 2박스로 못 보냄; EA>매출 noise 방어).
"""
from __future__ import annotations

import re
import unicodedata

import numpy as np
import pandas as pd

# 상호명(매출자료) → EA 판매처(들). 사용자 확정 2026-06-15.
SANGHO_TO_EA = {
    "오픈마켓- 스마트스토어": ["스마트스토어"],
    "오픈마켓(ESM/옥션.지마켓.G9)": ["G마켓", "옥션 #2"],
    "오픈마켓- (주) 마켓보로": ["식봄(마켓보로)"],
    "오픈마켓 (주) 한국신용데이터": ["캐시노트"],
    "오픈마켓- 알리": ["알리익스프레스자동"],
    "오픈마켓 쿠팡 (윙배송)": ["쿠팡(자동)"],
    "오픈마켓- (주) 우아한형제들": ["배민상회", "배민대용량장보기"],
    "오픈마켓- 올웨이즈": ["올웨이즈"],
    "제이티유통": ["제이티유통"],
}

_SHIP_STATES = ("배송", "송장")   # 실제 출고만(접수·빈 송장 제외)
_BOX_LINE = "00-12"               # 택배비 라인(상품 아님)


def _nfc(v) -> str:
    if v is None:
        return ""
    s = str(v)
    if s == "nan":
        return ""
    return unicodedata.normalize("NFC", s).strip()


def ea_to_sangho() -> dict:
    """EA 판매처(NFC) → 상호명(NFC) 역매핑."""
    out = {}
    for sangho, chans in SANGHO_TO_EA.items():
        for ch in chans:
            out[_nfc(ch)] = _nfc(sangho)
    return out


def allocate_boxes(orders: pd.DataFrame, box_lookup) -> pd.DataFrame:
    """주문 라인별 박스지분(_박스지분) 부여. 송장그룹 내 물류량(수량÷박스내품) 비중.

    box_lookup = 관리코드→박스내품 callable(또는 dict). 결측/0 → 1.0.
    단일 라인 박스 → 지분 1.0. 빈 송장·비출고 상태 제외.
    """
    df = orders.copy()
    df = df[df["송장그룹"].map(_nfc) != ""]
    if "상태" in df.columns:
        df = df[df["상태"].map(_nfc).isin(_SHIP_STATES)]
    df = df.copy()
    df["_code"] = df["erp관리코드"].map(_nfc)
    if callable(box_lookup):
        boxn = df["_code"].map(box_lookup)
    else:
        boxn = df["_code"].map(lambda c: box_lookup.get(c, 1.0))
    boxn = pd.to_numeric(boxn, errors="coerce").fillna(1.0)
    boxn = boxn.where(boxn > 0, 1.0)
    qty = pd.to_numeric(df["상품수량"], errors="coerce").fillna(0.0)
    df["_vol"] = qty / boxn                                    # 물류량
    tot = df.groupby("송장그룹")["_vol"].transform("sum")       # 박스 총 물류량
    df["_박스지분"] = (df["_vol"] / tot.where(tot != 0)).fillna(0.0)
    return df


def compute_box_counts(orders: pd.DataFrame, box_lookup, hapo_codes=None):
    """(상호명,관리코드,월)·(상호명,월) **실제 박스수**(박스 credit 합) + alloc df.

    return code_box{(상호명,code,ym): boxes}, ch_box{(상호명,ym): boxes}, alloc_df
    매핑 안 되는 EA 판매처(자사몰 등) 라인 제외.

    hapo_codes(=reference/hapo_175_190.csv 관리코드 set) 주면 **175~200 30개입 물리합포** 교정:
    합포 가능 품목을 **합포박스키(같은 수령자·주소·발주일) 그룹**으로 묶어 ceil(팩/3) 물리박스로
    환산(송장 분리돼도 한 박스에 최대 3팩) → credit = 팩비중 × ceil(팩/3). 비합포·박스키없음은
    송장그룹 지분 그대로(기존 동작). hapo_codes 없으면 전부 송장그룹(BEFORE).
    ★ 블랭킷 합포박스키 swap은 B2B 동일주소 다회주문을 1박스로 과대병합(-27%)하므로 금지 —
       합포 가능 품목에만 제한(logs/2026-06-16-ship-alloc-hapo).
    """
    alloc = allocate_boxes(orders, box_lookup)
    alloc["_credit"] = alloc["_박스지분"]                       # 기본=송장그룹 지분
    if hapo_codes and "합포박스키" in alloc.columns:
        hc = {re.sub(r"[.\-]+$", "", _nfc(c)) for c in hapo_codes}
        is_h = alloc["_code"].map(lambda c: re.sub(r"[.\-]+$", "", c) in hc)
        bk = alloc["합포박스키"].map(_nfc)
        h = alloc[is_h & (bk != "")].copy()
        if not h.empty:
            h["_bk"] = bk[h.index]
            h["_packs"] = pd.to_numeric(h["상품수량"], errors="coerce").fillna(0.0)
            tot = h.groupby("_bk")["_packs"].transform("sum")
            pboxes = np.ceil(tot / 3.0)                          # ceil(팩/3) 물리박스
            cr = (h["_packs"] / tot.where(tot != 0) * pboxes)
            alloc.loc[h.index, "_credit"] = cr.fillna(alloc.loc[h.index, "_박스지분"])
    rev = ea_to_sangho()
    alloc["_상호명"] = alloc["판매처"].map(_nfc).map(rev)        # 매핑 외(자사몰) → NaN
    alloc = alloc[alloc["_상호명"].notna()].copy()
    if alloc.empty:
        return {}, {}, alloc
    alloc["_ym"] = pd.to_datetime(alloc["기준일"]).dt.strftime("%Y-%m")
    code = (alloc.groupby(["_상호명", "_code", "_ym"], observed=True)["_credit"]
                 .sum().reset_index())
    code_box = {(r["_상호명"], r["_code"], r["_ym"]): r["_credit"]
                for r in code.to_dict("records")}
    ch = (alloc.groupby(["_상호명", "_ym"], observed=True)["_credit"]
               .sum().reset_index())
    ch_box = {(r["_상호명"], r["_ym"]): r["_credit"] for r in ch.to_dict("records")}
    return code_box, ch_box, alloc


def compute_ship_rate(orders: pd.DataFrame, sales: pd.DataFrame, box_lookup, hapo_codes=None,
                      sangho_col: str = "상호명", code_col: str = "관리코드",
                      date_col: str = "거래일자", qty_col: str = "수량") -> dict:
    """택배강도(박스/낱개) = EA 실제박스 ÷ 매출 낱개수. (상호명,code,월)·(상호명,월).

    sales = 매출 master(전체). 분모=매출 낱개수(00-12 제외, 매핑 상호명만). rate 상한 1.0.
    hapo_codes 주면 175~200 물리합포 ceil(팩/3) 교정(compute_box_counts 참조).
    return {"rate": {(상호명,code,ym): bpp}, "ch_rate": {(상호명,ym): bpp}, "stats": {...}}
    """
    empty = {"rate": {}, "ch_rate": {}, "reconcile": {},
             "stats": {"boxes": 0.0, "codes": 0, "months": [], "covered_pieces": 0.0}}
    if orders is None or orders.empty or sales is None or sales.empty:
        return empty
    code_box, ch_box, alloc = compute_box_counts(orders, box_lookup, hapo_codes)
    if not code_box:
        return empty

    # 매출 낱개수 (분모) — 매핑 상호명·상품행만
    mapped = {_nfc(s) for s in SANGHO_TO_EA}
    s = sales.copy()
    s["_sh"] = s[sangho_col].map(_nfc)
    s["_cd"] = s[code_col].map(_nfc)
    s = s[s["_sh"].isin(mapped) & (s["_cd"] != _BOX_LINE)].copy()
    if s.empty:
        return empty
    s["_ym"] = pd.to_datetime(s[date_col]).dt.strftime("%Y-%m")
    s["_q"] = pd.to_numeric(s[qty_col], errors="coerce").fillna(0.0)
    pieces = (s.groupby(["_sh", "_cd", "_ym"], observed=True)["_q"].sum())
    ch_pieces = (s.groupby(["_sh", "_ym"], observed=True)["_q"].sum())

    def _clamp(x):
        return 1.0 if x > 1.0 else (0.0 if x < 0 else x)

    rate = {}
    for g, boxes in code_box.items():
        p = pieces.get(g, 0.0)
        if p > 0:
            rate[g] = _clamp(boxes / p)
    ch_rate = {}
    for g, boxes in ch_box.items():
        p = ch_pieces.get(g, 0.0)
        if p > 0:
            ch_rate[g] = _clamp(boxes / p)
    covered = float(pieces.reindex(rate.keys()).sum())

    # 00-12 정합 스케일: (상호명,월) 매출 00-12 booked 송장수 ÷ raw 실측박스(Σ rate×pieces)
    #   → EA=상품별 분배 모양, 00-12=총 magnitude. (기존 k의 '총액 일치' 목적 보존)
    s0012 = sales.copy()
    s0012["_sh"] = s0012[sangho_col].map(_nfc)
    s0012 = s0012[s0012["_sh"].isin(mapped) & (s0012[code_col].map(_nfc) == _BOX_LINE)].copy()
    reconcile = {}
    if not s0012.empty:
        s0012["_ym"] = pd.to_datetime(s0012[date_col]).dt.strftime("%Y-%m")
        s0012["_q"] = pd.to_numeric(s0012[qty_col], errors="coerce").fillna(0.0)
        booked = s0012.groupby(["_sh", "_ym"], observed=True)["_q"].sum()
        raw_box = {}                                            # (상호명,월) raw 실측박스 — attach와 동일(ch_rate fallback 포함)
        for (sh, cd, ym), p in pieces.items():
            eff = rate.get((sh, cd, ym))
            if eff is None:
                eff = ch_rate.get((sh, ym))
            if eff:
                raw_box[(sh, ym)] = raw_box.get((sh, ym), 0.0) + eff * p
        for key, rb in raw_box.items():
            bk = booked.get(key, 0.0)
            if rb > 0 and bk > 0:
                reconcile[key] = bk / rb
    return {"rate": rate, "ch_rate": ch_rate, "reconcile": reconcile,
            "stats": {"boxes": float(sum(code_box.values())), "codes": len(rate),
                      "months": sorted({k[2] for k in code_box}),
                      "covered_pieces": covered}}


def attach_actual_ship(prod: pd.DataFrame, rate: dict, unit: float, reconcile: bool = True,
                       sangho_col: str = "상호명", code_col: str = "관리코드",
                       date_col: str = "거래일자", qty_col: str = "수량") -> pd.Series:
    """매출 행(prod)에 실측 택배비 Series 반환 = 수량 × rate × 단가. rate 없으면 NaN.

    code-level rate 우선, 없으면 채널(상호명)-level fallback. 둘 다 없으면 NaN(페이지가 추정).
    reconcile=True → (상호명,월) 00-12 정합 스케일 곱(EA 분배 모양·00-12 총액).
    """
    r_code, r_ch, rec = rate.get("rate", {}), rate.get("ch_rate", {}), rate.get("reconcile", {})
    sh = prod[sangho_col].map(_nfc)
    cd = prod[code_col].map(_nfc)
    ym = pd.to_datetime(prod[date_col]).dt.strftime("%Y-%m")
    qty = pd.to_numeric(prod[qty_col], errors="coerce").fillna(0.0)
    ck = list(zip(sh, cd, ym))
    hk = list(zip(sh, ym))
    rr = pd.Series([r_code.get(k) for k in ck], index=prod.index, dtype="float64")
    rr_ch = pd.Series([r_ch.get(k) for k in hk], index=prod.index, dtype="float64")
    rr = rr.fillna(rr_ch)
    out = (qty * rr * unit).astype("float64")
    if reconcile and rec:
        sc = pd.Series([rec.get(k, 1.0) for k in hk], index=prod.index, dtype="float64")
        out = out * sc
    return out                                          # NaN where no rate
