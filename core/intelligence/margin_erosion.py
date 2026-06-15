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


# ─────────────────────────────────────────────────────────
# 두뇌① 강화 (ADR 0020): velocity 가중 · 2C 예방경보 · 실판매 마진 이상
# 채널 매핑: EA판매처(orders)·상호명(매출) → cmm baseline 채널.
# ─────────────────────────────────────────────────────────

# EA 판매처(orders 'raw') → cmm baseline 채널. 미매핑(제이티·자사몰)=오프라인성(None).
EA_TO_CMM = {
    "스마트스토어": "스마트스토어",
    "G마켓": "ESM", "옥션 #2": "ESM",
    "식봄(마켓보로)": "식봄",
    "캐시노트": "캐시노트",
    "알리익스프레스자동": "알리",
    "쿠팡(자동)": "쿠팡",
    "배민상회": "배민상회", "배민대용량장보기": "배민상회",
    "올웨이즈": "올웨이즈",
}


def _sangho_to_cmm() -> dict:
    """매출 상호명(NFC) → cmm 채널. ship_alloc.SANGHO_TO_EA ∘ EA_TO_CMM."""
    from core.intelligence import ship_alloc as sa
    out = {}
    for sangho, eas in sa.SANGHO_TO_EA.items():
        for ea in eas:
            ch = EA_TO_CMM.get(_nfc(ea))
            if ch:
                out[_nfc(sangho)] = ch
                break
    return out


def channel_velocity(orders, box_lookup, months: int = 3, now=None) -> dict:
    """최근 N개월 orders → (cmm채널, 관리코드) 월 낱개 판매량.

    낱개 = 상품수량(EA 판매단위) × 박스내품. 출고 상태(배송/송장)만.
    box_lookup = 관리코드→박스내품 dict/callable(결측·0→1.0).
    return {"by_code_ch": {(ch,code): 월낱개}, "by_code": {code: 월낱개합}}.
    """
    empty = {"by_code_ch": {}, "by_code": {}}
    if orders is None or orders.empty:
        return empty
    now = pd.Timestamp(now) if now is not None else pd.Timestamp.now()
    cutoff = now - pd.DateOffset(months=months)
    df = orders.copy()
    df = df[pd.to_datetime(df["기준일"]) >= cutoff]
    if "상태" in df.columns:
        df = df[df["상태"].map(_nfc).isin(("배송", "송장"))]
    if df.empty:
        return empty
    df = df.copy()
    df["_ch"] = df["판매처"].map(_nfc).map(EA_TO_CMM)
    df = df[df["_ch"].notna()].copy()
    df["_code"] = df["erp관리코드"].map(_nfc)
    if callable(box_lookup):
        boxn = df["_code"].map(box_lookup)
    else:
        boxn = df["_code"].map(lambda c: box_lookup.get(c, 1.0))
    boxn = pd.to_numeric(boxn, errors="coerce").fillna(1.0)
    boxn = boxn.where(boxn > 0, 1.0)
    qty = pd.to_numeric(df["상품수량"], errors="coerce").fillna(0.0)
    df["_pieces"] = qty * boxn
    g = df.groupby(["_ch", "_code"])["_pieces"].sum() / months
    by_code_ch = {(k[0], k[1]): float(v) for k, v in g.items()}
    gc = df.groupby("_code")["_pieces"].sum() / months
    by_code = {k: float(v) for k, v in gc.items()}
    return {"by_code_ch": by_code_ch, "by_code": by_code}


def latest_buyin_price(buyin, months: int = 6, now=None) -> dict:
    """관리코드별 최근 N개월 '가장 최근 입고일'의 수량가중 평균 실입고 단가(낱개).

    반품(합계<0)·0단가/0수량 제외. return {관리코드(NFC): {실입고가, 입고일, 입고수량}}.
    """
    if buyin is None or buyin.empty:
        return {}
    now = pd.Timestamp(now) if now is not None else pd.Timestamp.now()
    cutoff = now - pd.DateOffset(months=months)
    b = buyin.copy()
    b = b[pd.to_datetime(b["기준일"]) >= cutoff]
    q = pd.to_numeric(b["수량"], errors="coerce")
    p = pd.to_numeric(b["단가"], errors="coerce")
    b = b[(p > 0) & (q > 0)].copy()
    if b.empty:
        return {}
    b["_code"] = b["관리코드"].map(_nfc)
    # 낱개단가 정규화: '단가==박스단가 & 박스내품>1' = 박스단가가 단가칸에 오입력(3%) → 박스단가/박스내품
    up = pd.to_numeric(b["단가"], errors="coerce")
    bp = pd.to_numeric(b["박스단가"], errors="coerce")
    bn = pd.to_numeric(b["박스내품"], errors="coerce")
    box_mistyped = (bn > 1) & (bp > 0) & ((up - bp).abs() < 1)
    b["_unit"] = up.where(~box_mistyped, bp / bn)
    out = {}
    for code, g in b.groupby("_code"):
        if not code:
            continue
        last_day = g["기준일"].max()
        gl = g[g["기준일"] == last_day]
        qq = pd.to_numeric(gl["수량"], errors="coerce").fillna(0.0)
        pp = pd.to_numeric(gl["_unit"], errors="coerce").fillna(0.0)
        wprice = float((qq * pp).sum() / qq.sum()) if qq.sum() > 0 else float(pp.mean())
        out[code] = {"실입고가": wprice, "입고일": last_day, "입고수량": float(qq.sum())}
    return out


def pending_buyin_raises(buyin, master_price: dict, raises: dict,
                         months: int = 6, min_pct: float = 0.03, max_pct: float = 0.6,
                         now=None) -> dict:
    """2C 예방: 최근 실입고가 > master 매입가(min_pct 이상) ∩ master 미수정(∉ raises).

    master_price = {관리코드(NFC): master 매입단가(낱개)}. raises = recent_buy_raises 결과.
    min_pct ≤ 입고Δ% ≤ max_pct 만 경보(상한 초과 = 관리코드 충돌/단위잔류/극과거 master 의심 → suspect).
    return {"alerts": {코드: {...}}, "suspect": {코드: {...}}} — alerts=경보, suspect=검토 필요.
    """
    latest = latest_buyin_price(buyin, months=months, now=now)
    alerts, suspect = {}, {}
    for code, info in latest.items():
        if code in raises:                       # master 이미 인상(탭A에서 잡힘) → 제외
            continue
        m = master_price.get(code)
        if m is None or m <= 0:
            continue
        rp = info["실입고가"]
        if rp <= m * (1 + min_pct):
            continue
        rec = {"master매입": m, "실입고가": rp, "입고Δ%": (rp - m) / m,
               "입고일": info["입고일"], "입고수량": info["입고수량"]}
        (suspect if rec["입고Δ%"] > max_pct else alerts)[code] = rec
    return {"alerts": alerts, "suspect": suspect}


def sales_margin_anomalies(sales, baseline_dict, months: int = 3,
                           buffer: float = 0.02, now=None) -> list:
    """실판매 마진 이상: 역마진 OR 채널 baseline−buffer 미달. (cmm채널, 관리코드) 집계.

    baseline_dict = {관리코드(NFC): {채널: 기준마진}} (cmm.parse_baseline_dict).
    오프라인(상호명 미매핑) = 역마진만. 마진율 = 판매이익/판매금액(실현·정산).
    return rows(dict list).
    """
    if sales is None or sales.empty:
        return []
    now = pd.Timestamp(now) if now is not None else pd.Timestamp.now()
    cutoff = now - pd.DateOffset(months=months)
    s = sales.copy()
    s = s[pd.to_datetime(s["거래일자"]) >= cutoff]
    if s.empty:
        return []
    smap = _sangho_to_cmm()
    s = s.copy()
    s["_ch"] = s["상호명"].map(_nfc).map(smap).fillna("오프라인")
    s["_code"] = s["관리코드"].map(_nfc)
    g = s.groupby(["_ch", "_code"]).agg(
        판매금액=("판매금액", "sum"), 판매이익=("판매이익", "sum"),
        수량=("수량", "sum"), 상품명=("상품명", "first")).reset_index()
    g = g[g["판매금액"] != 0].copy()
    g["마진율"] = g["판매이익"] / g["판매금액"]
    rows = []
    for r in g.to_dict("records"):
        ch, code, mr = r["_ch"], r["_code"], r["마진율"]
        base = None if ch == "오프라인" else (baseline_dict.get(code, {}) or {}).get(ch)
        neg = mr < 0
        under = (base is not None) and (mr < base - buffer)
        if not (neg or under):
            continue
        rows.append({"채널": ch, "관리코드": code, "상품명": r["상품명"],
                     "판매금액": r["판매금액"], "판매이익": r["판매이익"], "수량": r["수량"],
                     "마진율": mr, "기준마진율": base,
                     "미달폭": (mr - base) if base is not None else None, "역마진": neg})
    return rows
