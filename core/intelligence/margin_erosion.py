"""두뇌① 마진 침식 탐지 — 최근 매입가 인상으로 채널 기준마진 아래로 떨어진(아직 미수정) 상품.

intelligence-layer.md §6 ①. 입력:
- 가격이력 price_changes(1a) — 최근 N개월 매입단가 인상 이벤트(관리코드 키).
- 채널 listing(현재 판매가) + baseline_margin(채널 기준마진) → cmm.compute_listing 재사용(마진율·탐지·권장가).
신호 = (최근 매입가 인상) ∩ (현재 채널 마진 < 기준마진). 사용자가 이미 재설정한 건 마진이 기준 이상이라 자동 제외.
범위: base 매익률(오프라인) 아님 — 채널별 baseline(온라인) 기준. 사용자 확정 2026-06-15.
한계: listing '코드'가 관리코드 직매칭인 행만 조인(박스형 다수). 합포/PC낱개/소분 합성코드는 v1 제외.
"""
from __future__ import annotations

import unicodedata

import pandas as pd


def _nfc(v):
    return unicodedata.normalize("NFC", str(v)).strip() if v is not None else ""


def recent_buy_raises(price_changes: pd.DataFrame, months: int = 3, now=None) -> dict:
    """관리코드별 최근 N개월 '매입단가 인상' 요약.

    상품코드 단위로 (window 시작 수정전 → 최신 수정후) 계산, 인상(현재>과거)만.
    관리코드 충돌 시 매입Δ%가 가장 큰 상품코드를 대표로(주 원가동인).
    return {관리코드(NFC): {상품코드, 상품명, 과거매입, 현재매입, 매입Δ, 매입Δ%, 마지막인상일, 상품수}}.
    """
    if price_changes is None or price_changes.empty:
        return {}
    now = pd.Timestamp(now) if now is not None else pd.Timestamp.now()
    cutoff = now - pd.DateOffset(months=months)
    buy = price_changes[(price_changes["수정항목"] == "매입단가")
                        & (pd.to_datetime(price_changes["수정일자"]) >= cutoff)].copy()
    if buy.empty:
        return {}
    buy["수정일자"] = pd.to_datetime(buy["수정일자"])
    buy = buy.sort_values("수정일자")

    per_code = []  # 상품코드 단위 인상
    for sc, g in buy.groupby("상품코드"):
        past = float(g.iloc[0]["수정전"])
        cur = float(g.iloc[-1]["수정후"])
        if pd.isna(past) or pd.isna(cur) or past <= 0 or cur <= past:
            continue
        rises = g[g["수정후"] > g["수정전"]]
        per_code.append({
            "상품코드": _nfc(sc),
            "관리코드": _nfc(g.iloc[-1]["관리코드"]),
            "상품명": _nfc(g.iloc[-1]["상품명"]),
            "과거매입": past, "현재매입": cur, "매입Δ": cur - past, "매입Δ%": (cur - past) / past,
            "마지막인상일": (rises["수정일자"].max() if len(rises) else g["수정일자"].max()),
        })

    # 관리코드로 collapse — Δ% 최대 대표 + 상품수
    out = {}
    by_mgmt = {}
    for rec in per_code:
        by_mgmt.setdefault(rec["관리코드"], []).append(rec)
    for mgmt, recs in by_mgmt.items():
        if not mgmt:
            continue
        rep = max(recs, key=lambda r: r["매입Δ%"])
        out[mgmt] = {**rep, "상품수": len(recs)}
    return out


# 침식 판정 임계(채널마진모니터와 동일 어휘). 탐지=마진율−기준마진율 < 이 값이면 미달.
UNDER_THRESHOLD = -0.01


def erosion_rows(rows: list, channel: str, raises: dict,
                 under_threshold: float = UNDER_THRESHOLD) -> list:
    """cmm.compute_listing 결과(rows) + raises → 침식 경보 행.

    조건: ① 현재 마진이 기준마진보다 낮음(탐지 < under_threshold) AND ② 최근 매입가 인상(관리코드 ∈ raises).
    사용자가 이미 가격 재설정한 상품은 탐지 ≥ 0 → 자동 제외.
    """
    out = []
    for r in rows:
        det = r.get("탐지")
        if det is None or det >= under_threshold:
            continue                                   # 기준 이상(정상/이미 수정) → 제외
        mgmt = _nfc(r.get("관리코드"))
        ri = raises.get(mgmt)
        if not ri:
            continue                                   # 최근 매입가 인상 없음 → 제외(묵은 미달 제외)
        out.append({
            "채널": channel,
            "관리코드": r.get("관리코드"),
            "상품명": r.get("상품명") or ri.get("상품명"),
            "과거매입": ri["과거매입"], "현재매입": ri["현재매입"], "매입Δ%": ri["매입Δ%"],
            "마지막인상일": ri["마지막인상일"], "상품수": ri.get("상품수", 1),
            "현재판매가": r.get("판매가"),
            "마진율": r.get("마진율"), "기준마진율": r.get("기준마진율"), "미달폭": det,
            "권장가": r.get("권장가"),
            "재고": r.get("재고"),
        })
    return out


def to_frame(all_rows: list, sort: str = "미달폭") -> pd.DataFrame:
    """침식 행 list → 정렬 DataFrame. sort='미달폭'(기준대비 부족분 큰 순) | '매입Δ%'(상승폭 큰 순)."""
    cols = ["채널", "관리코드", "상품명", "과거매입", "현재매입", "매입Δ%", "마지막인상일",
            "현재판매가", "마진율", "기준마진율", "미달폭", "권장가", "재고", "상품수"]
    df = pd.DataFrame(all_rows, columns=cols)
    if df.empty:
        return df
    if sort == "매입Δ%":
        df = df.sort_values("매입Δ%", ascending=False)
    else:  # 미달폭: 탐지가 음수 → 가장 작은(=가장 미달) 먼저
        df = df.sort_values("미달폭", ascending=True)
    return df.reset_index(drop=True)
