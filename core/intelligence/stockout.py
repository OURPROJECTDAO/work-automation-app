"""두뇌② 입고·품절 예측 — 현재고 ÷ 소진율 → 소진 예측일·재발주 시점.

intelligence-layer.md §6 ②. 입력(전부 적립 완료·관리코드 직조인):
- 현재고: product_master 최종재고(낱개 = 박스×박스내품+낱개, 일치율 100% 검증).
- 소진율: 매출자료(전채널·정산진실·이미 낱개 분해) 최근 N개월 수량 ÷ 일수.
  ★orders(온라인만)보다 매출자료(전채널)가 물리재고 소진에 정확. base 관리코드 직조인이라
   두뇌①의 '합성코드 미조인' 한계 없음(매출자료엔 합포/PC/소분 코드가 없음 — 실증 2026-06-15).
- 입고주기(리드타임 proxy): 매입현황 관리코드별 중앙 입고간격(발주→입고 실리드타임은 발주자료 적재 후 후속).

산출(관리코드 단위): 일소진·소진예측일·예상소진일자·입고주기·재발주필요·재고금액·구간(4종).
한계: 박스재고 절대값은 흐름누적 오차(이월·조정) → product_master 앵커값 사용. 매출≠물리출고 100% 아님(추세 신호).
"""
from __future__ import annotations

import unicodedata

import numpy as np
import pandas as pd

# 4구간 (사용자 설계 2026-06-15)
B_IMMINENT = "🔴 품절임박"   # 소진예측일 ≤ 리드타임
B_SOON = "🟡 곧재발주"        # 리드타임 < 소진예측일 ≤ 리드타임×1.5
B_OK = "🟢 충분"             # 소진예측일 > 리드타임×1.5 (판매 있음)
B_DEAD = "⚪ 사장재고"        # 매출 없음(소진≈0) & 박스재고 있음 (과잉·역신호)


def _nfc(v):
    return unicodedata.normalize("NFC", str(v)).strip() if pd.notna(v) else ""


def _num(s):
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False) if s.dtype == object else s,
        errors="coerce",
    ).fillna(0.0)


def depletion_rate(sales: pd.DataFrame, months: int = 3, now=None) -> dict:
    """매출자료 → 관리코드별 일평균 낱개 소진율.

    최근 N개월 수량 합 ÷ 윈도우 실일수. sales 컬럼: 거래일자·관리코드·수량(낱개).
    return {관리코드(NFC): {일소진, 월낱개, 최근판매일}}.
    """
    if sales is None or sales.empty:
        return {}
    now = pd.Timestamp(now) if now is not None else pd.Timestamp.now()
    cutoff = now - pd.DateOffset(months=months)
    s = sales.copy()
    s["_d"] = pd.to_datetime(s["거래일자"], errors="coerce")
    s = s[(s["_d"] >= cutoff) & (s["_d"] <= now)].copy()
    if s.empty:
        return {}
    days = max((now.normalize() - cutoff.normalize()).days, 1)
    s["_code"] = s["관리코드"].map(_nfc)
    s["_qty"] = _num(s["수량"])
    g = s.groupby("_code").agg(_q=("_qty", "sum"), _last=("_d", "max"))
    out = {}
    for code, r in g.iterrows():
        if not code:
            continue
        q = float(r["_q"])
        out[code] = {"일소진": q / days, "월낱개": q / months, "최근판매일": r["_last"]}
    return out


def restock_cadence(buyin: pd.DataFrame, months: int = 12, now=None) -> dict:
    """매입현황 → 관리코드별 입고 주기(리드타임 proxy).

    실입고(합계액>0 & 수량>0)만, 입고일(일자 중복 제거) 사이 중앙 간격.
    return {관리코드(NFC): {입고주기, 최근입고일, 입고횟수}} (입고 1회면 입고주기=None).
    """
    if buyin is None or buyin.empty:
        return {}
    now = pd.Timestamp(now) if now is not None else pd.Timestamp.now()
    cutoff = now - pd.DateOffset(months=months)
    b = buyin.copy()
    b["_d"] = pd.to_datetime(b["기준일"], errors="coerce")
    b = b[(b["_d"] >= cutoff) & (b["_d"] <= now)]
    b = b[(_num(b["합계액"]) > 0) & (_num(b["수량"]) > 0)].copy()
    if b.empty:
        return {}
    b["_code"] = b["관리코드"].map(_nfc)
    out = {}
    for code, g in b.groupby("_code"):
        if not code:
            continue
        days = sorted(set(g["_d"].dt.normalize()))
        last = days[-1]
        if len(days) >= 2:
            gaps = np.diff([d.value for d in days]) / (86400 * 1e9)
            cad = float(np.median(gaps))
        else:
            cad = None
        out[code] = {"입고주기": cad, "최근입고일": last, "입고횟수": len(days)}
    return out


def forecast(pm: pd.DataFrame, depletion: dict, cadence: dict, now=None,
             default_lead_days: float = 14.0, scope: str = "stock_or_velocity",
             exclude_codes=None, exclude_midcat=None) -> pd.DataFrame:
    """현재고 + 소진율 + 입고주기 → 관리코드별 품절·재발주 예측.

    pm = product_master(원본 컬럼). depletion = depletion_rate(). cadence = restock_cadence().
    scope: 'stock_or_velocity'(박스재고>0 또는 소진>0) | 'stock'(박스재고>0만).
    리드타임 = 입고주기(있으면) 아니면 default_lead_days. 재발주필요 = 소진예측일 ≤ 리드타임.
    비판매 제외: 택배비(상품명) + exclude_codes(상품코드·업로드감시 제외목록) + exclude_midcat(중분류).
    음수 재고(이월·조정 drift) = 0으로 클램프(소진예측일 0 = 즉시 품절).
    """
    now = pd.Timestamp(now) if now is not None else pd.Timestamp.now()
    exclude_codes = {_nfc(c) for c in (exclude_codes or set())}
    exclude_midcat = {_nfc(c) for c in (exclude_midcat or set())}
    d = pm.copy()
    d.columns = [c.strip() for c in d.columns]
    for c in ("박스내품", "최종재고", "박스", "낱개", "박스매입단가", "매입단가"):
        if c in d.columns:
            d[c] = _num(d[c])
        else:
            d[c] = 0.0

    rows = []
    for r in d.to_dict("records"):
        code = _nfc(r.get("관리코드", ""))
        if not code:
            continue
        name = _nfc(r.get("상품명", ""))
        if "택배비" in name:                                    # 회계 의사코드(00-1x)
            continue
        if _nfc(r.get("상품코드", "")) in exclude_codes:        # 업로드감시 비판매 제외목록
            continue
        if _nfc(r.get("중분류명", "")) in exclude_midcat:       # 반품·파렛트 등
            continue
        stock_pieces = float(r.get("최종재고", 0.0) or 0.0)
        box_stock = float(r.get("박스", 0.0) or 0.0)
        dinfo = depletion.get(code)
        daily = float(dinfo["일소진"]) if dinfo else 0.0
        if scope == "stock" and box_stock <= 0:
            continue
        if scope != "stock" and box_stock <= 0 and daily <= 0:
            continue
        stock_eff = max(stock_pieces, 0.0)                      # 음수재고 클램프

        cad = cadence.get(code, {})
        cad_days = cad.get("입고주기")
        lead = float(cad_days) if cad_days and cad_days > 0 else float(default_lead_days)
        lead_src = "입고주기" if (cad_days and cad_days > 0) else "기본"

        if daily > 0:
            ttl = stock_eff / daily               # 소진예측일(음수재고=0→즉시품절)
            eta = now.normalize() + pd.Timedelta(days=float(min(ttl, 3650)))
        else:
            ttl = float("inf")
            eta = pd.NaT

        if daily <= 0:
            band = B_DEAD if box_stock > 0 else B_OK
            reorder = False
        elif ttl <= lead:
            band = B_IMMINENT
            reorder = True
        elif ttl <= lead * 1.5:
            band = B_SOON
            reorder = False
        else:
            band = B_OK
            reorder = False

        box_qty = float(r.get("박스내품", 0.0) or 0.0)
        unit_buy = float(r.get("매입단가", 0.0) or 0.0)        # 낱개 매입단가
        month_pieces = float(dinfo["월낱개"]) if dinfo else 0.0
        rows.append({
            "구간": band,
            "관리코드": r.get("관리코드", code),
            "상품명": r.get("상품명", ""),
            "규격": r.get("규격", ""),
            "현재고(낱개)": stock_pieces,
            "박스재고": box_stock,
            "박스내품": box_qty,
            "일소진(낱개)": round(daily, 2),
            "일소진(박스)": round(daily / box_qty, 3) if box_qty > 0 else round(daily, 3),
            "일소진액(매입)": round(daily * unit_buy),       # 하루 매입가 기준 빠지는 금액
            "월낱개": round(month_pieces, 1),
            "월소진액(매입)": round(month_pieces * unit_buy),  # 월 매입가 기준 볼륨
            "소진예측일": (None if not np.isfinite(ttl) else round(ttl, 1)),
            "예상소진일자": (eta.date() if pd.notna(eta) else None),
            "리드타임": round(lead, 1),
            "리드출처": lead_src,
            "최근입고일": (cad.get("최근입고일").date() if cad.get("최근입고일") is not None else None),
            "입고횟수": cad.get("입고횟수", 0),
            "재고금액": round(max(box_stock, 0.0) * float(r.get("박스매입단가", 0.0) or 0.0)),
            "재발주필요": reorder,
        })

    cols = ["구간", "관리코드", "상품명", "규격", "현재고(낱개)", "박스재고", "박스내품",
            "일소진(낱개)", "일소진(박스)", "일소진액(매입)", "월낱개", "월소진액(매입)",
            "소진예측일", "예상소진일자", "리드타임", "리드출처",
            "최근입고일", "입고횟수", "재고금액", "재발주필요"]
    df = pd.DataFrame(rows, columns=cols)
    if df.empty:
        return df
    # 정렬: 재발주필요 먼저 → 소진예측일 오름(임박 먼저, inf 뒤) → 재고금액 desc
    df["_ttl_sort"] = df["소진예측일"].map(lambda v: v if v is not None else 1e12)
    df = df.sort_values(["재발주필요", "_ttl_sort", "재고금액"],
                        ascending=[False, True, False]).drop(columns="_ttl_sort")
    return df.reset_index(drop=True)


def summarize(df: pd.DataFrame) -> dict:
    """구간별 건수·재고금액 요약."""
    if df is None or df.empty:
        return {"건수": {}, "재고금액": {}, "재발주필요": 0}
    cnt = df["구간"].value_counts().to_dict()
    val = df.groupby("구간")["재고금액"].sum().to_dict()
    return {"건수": cnt, "재고금액": {k: int(v) for k, v in val.items()},
            "재발주필요": int(df["재발주필요"].sum())}
