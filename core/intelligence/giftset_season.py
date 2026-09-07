"""선물세트 시즌 관제판 — 시즌 한정 트랙(추석/설)의 진도·재고·채널 실적 통합판.

정본 = workflows/giftset-season-ops.md. 구 VBA 워크북 `스마트스토어_마진율_계산기.xlsm`
`선물세트` 시트(SUMIFS 34행 하드코딩)를 이관한 것.

구 시트 대비 바로잡은 것:
- 행 = 하드코딩 34행 → **자동 산출**(선물세트 유니버스 ∩ (박스재고>0 ∪ 시즌매출>0)).
- 채널 = 상호명 하드코딩(멸치·11번가 등 사장 채널 포함) → 상호명 규칙 매핑(자사몰 편입).
- W~Z = 이름은 '판매량/박스'인데 값은 금액이라 단위가 깨져 있었음 → **수량(세트) 기반**으로 재구축.
- 이익/마진·YoY·D-day 진도 축 신설(구 시트엔 매출액뿐).

★단위 규약: 매출자료(master/sales_*.parquet)의 `수량`은 **세트 단위**(선물세트는 1세트=1수량,
  박스가 아니다). 재고는 product_master `박스`(박스재고)라 세트 환산은 ×박스내품.
★스코프: 목표(ADR 0030)는 **오픈마켓 한정**. 리테일·자사몰·나들·오프라인은 참고 집계다.
"""
from __future__ import annotations

import json
import unicodedata

import numpy as np
import pandas as pd

# ── 채널 규칙 (상호명 → 채널군·채널) ────────────────────────────────────────
GRP_OPEN = "오픈마켓"
GRP_RETAIL = "리테일"
GRP_MALL = "자사몰"
GRP_NADL = "나들"
GRP_OFF = "오프라인"
GRP_SKIP = "제외"

ONLINE_GROUPS = (GRP_OPEN, GRP_RETAIL, GRP_MALL)

# (채널명, 상호명에 포함되면 매칭) — 오픈마켓 하위 채널. 순서 = 판정 우선순위.
OPEN_CHANNELS = [
    ("ESM", ("ESM", "지마켓")),
    ("스마트스토어", ("스마트",)),
    ("쿠팡", ("쿠팡",)),
    ("캐시노트", ("한국신용데이터",)),
    ("알리", ("알리",)),
    ("식봄", ("마켓보로",)),
    ("올웨이즈", ("올웨이즈",)),
    ("배민", ("우아한형제들",)),
    ("토스", ("토스뱅크",)),
]
# 위탁재고 계정 — 매출처가 아니므로 어느 집계에도 넣지 않는다(pitfalls 2026-08-05).
EXCLUDE_STORES = ("쿠팡(로켓창고)",)

# 판정 라벨 (§6 4분면)
V_CUT = "🔴 인하 집행"
V_BUY = "🟡 추가매입 검토"
V_HOLD = "🟢 유지"
V_UP = "🔵 인상 검토"
V_NOSALE = "⚫ 무매출"

DEFAULT_CONFIG = {
    "시즌명": "추석 2026",
    "dday": "2026-09-25",
    "시즌창_일수": 75,
    "마감_D": 5,
    "작년_dday": "2025-10-06",
    "매출목표": 1_500_000_000,
    "계획마진": 0.088,
    "마진하한": 0.073,
    "절대하한": 0.03,
    "이월비용률": 0.033,
    "밴드": {
        "25": [0.01, 0.10], "20": [0.11, 0.19], "18": [0.21, 0.31],
        "15": [0.43, 0.61], "13": [0.64, 0.71], "10": [0.84, 0.92],
        "8": [0.89, 0.98], "5": [0.99, 1.00],
    },
}


def load_config(text: bytes | str | None) -> dict:
    """reference/giftset_season.json 로드. 없거나 깨지면 DEFAULT_CONFIG."""
    if not text:
        return dict(DEFAULT_CONFIG)
    try:
        if isinstance(text, bytes):
            text = text.decode("utf-8-sig")
        cfg = json.loads(text)
    except Exception:
        return dict(DEFAULT_CONFIG)
    out = dict(DEFAULT_CONFIG)
    out.update(cfg or {})
    return out


def _nfc(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return unicodedata.normalize("NFC", str(v)).strip()


def _num(v, default=0.0) -> float:
    if v is None:
        return default
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return default if pd.isna(v) else float(v)
    s = str(v).replace(",", "").strip()
    if not s:
        return default
    try:
        return float(s)
    except ValueError:
        return default


def classify_store(name) -> tuple[str, str]:
    """상호명 → (채널군, 채널). 오픈마켓만 하위 채널을 가른다."""
    s = _nfc(name)
    if not s:
        return GRP_OFF, "미상"
    if s in EXCLUDE_STORES:
        return GRP_SKIP, "제외"
    if "나들커뮤니케이션" in s:
        return GRP_NADL, "나들"
    if "리테일앤인사이트" in s:
        return GRP_RETAIL, "리테일"
    if "자사몰" in s:
        return GRP_MALL, "자사몰"
    if s.startswith("오픈마켓"):
        for ch, keys in OPEN_CHANNELS:
            if any(k in s for k in keys):
                return GRP_OPEN, ch
        return GRP_OPEN, "기타오픈마켓"
    return GRP_OFF, s


def channel_columns() -> list[str]:
    return [c for c, _ in OPEN_CHANNELS] + ["리테일", "자사몰", "나들", "오프라인"]


# ── 시즌 창 / D-day ──────────────────────────────────────────────────────────
def season_window(cfg: dict, now=None) -> dict:
    """시즌 파라미터를 타임스탬프로 풀어낸다."""
    now = pd.Timestamp(now or pd.Timestamp.now().normalize()).normalize()
    dday = pd.Timestamp(cfg["dday"])
    ly_dday = pd.Timestamp(cfg["작년_dday"])
    span = int(cfg.get("시즌창_일수", 75))
    start = dday - pd.Timedelta(days=span)
    ly_start = ly_dday - pd.Timedelta(days=span)
    D = int((dday - now).days)
    close_d = int(cfg.get("마감_D", 5))
    return {
        "now": now, "dday": dday, "start": start, "end": now,
        "ly_dday": ly_dday, "ly_start": ly_start, "ly_end": ly_dday,
        "D": D,
        "잔여일": max(int(D - close_d), 0),          # D-5(배송 마감)까지 남은 판매일
        "경과영업일": len(pd.bdate_range(start, now)),
    }


def ly_aligned_dates(win: dict) -> dict:
    """작년 대조 시점 2종.

    - 달력: 작년 같은 D-day 오프셋.
    - 영업일: 올해 경과 영업일 수와 같은 순번의 작년 영업일 **달력 날짜**.
      ★평일 필터 후 cumsum 금지(작년 주말 매출이 통째로 소실 — pitfalls 2026-09-03).
        반드시 '그 날짜까지 주말 포함 누적'과 대조한다.
    """
    cal = win["ly_dday"] - pd.Timedelta(days=max(win["D"], 0))
    bd = pd.bdate_range(win["ly_start"], win["ly_dday"])
    n = min(max(win["경과영업일"], 1), len(bd))
    return {"달력": cal, "영업일": bd[n - 1]}


# ── 유니버스 ─────────────────────────────────────────────────────────────────
def build_universe(pm: pd.DataFrame, attrs: pd.DataFrame, lineup: pd.DataFrame,
                   season_codes: set | None = None) -> pd.DataFrame:
    """표시 대상 행 = 선물세트 유니버스 ∩ (박스재고>0 ∪ 시즌매출>0).

    선물세트 판정 = product_attributes.최종분류=='선물세트' **∪** 시즌 라인업 정본.
    한쪽에만 있는 코드는 `결손` 플래그를 달아 기준데이터 구멍 탐지를 겸한다
    (2026-08-27: 3종이 attributes에 행 자체가 없어 등재갭 분석에서 누락된 전례).
    """
    attr_codes, line_codes = set(), set()
    if attrs is not None and len(attrs):
        a = attrs.copy()
        a["관리코드"] = a["관리코드"].map(_nfc)
        attr_codes = set(a.loc[a["최종분류"].map(_nfc) == "선물세트", "관리코드"]) - {""}
    if lineup is not None and len(lineup):
        line_codes = set(lineup["관리코드"].map(_nfc)) - {""}
    universe = attr_codes | line_codes
    if not universe:
        return pd.DataFrame()

    m = pm.copy()
    m["관리코드"] = m["관리코드"].map(_nfc)
    m = m[m["관리코드"].isin(universe)]
    # 관리코드 중복(상품코드 여러 행) → 첫 행
    m = m.drop_duplicates(subset=["관리코드"], keep="first")

    rows = []
    sold = season_codes or set()
    for _, r in m.iterrows():
        code = r["관리코드"]
        inner = _num(r.get("박스내품"), 1) or 1
        box = _num(r.get("박스"))
        box_cost = _num(r.get("박스매입단가"))
        rows.append({
            "관리코드": code,
            "상품명": _nfc(r.get("상품명")),
            "박스내품": inner,
            "박스재고": box,
            "세트재고": box * inner,
            "박스매입단가": box_cost,
            "세트원가": box_cost / inner if inner else 0.0,
            "재고금액": box * box_cost,
            "시즌라인업": code in line_codes,
            "결손": ("attributes 없음" if code not in attr_codes else
                     ("라인업 없음" if code not in line_codes else "")),
        })
    # product_master에 아예 없는데 매출은 있는 코드도 살린다
    for code in sorted((universe & sold) - set(m["관리코드"])):
        rows.append({
            "관리코드": code, "상품명": "", "박스내품": 1.0, "박스재고": 0.0,
            "세트재고": 0.0, "박스매입단가": 0.0, "세트원가": 0.0, "재고금액": 0.0,
            "시즌라인업": code in line_codes, "결손": "product_master 없음",
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    keep = (df["박스재고"] > 0) | (df["관리코드"].isin(sold))
    return df[keep].reset_index(drop=True)


# ── 매출 집계 ────────────────────────────────────────────────────────────────
def scope_sales(sales: pd.DataFrame, start, end, codes: set | None = None) -> pd.DataFrame:
    """시즌창으로 자르고 채널군·채널을 붙인다."""
    if sales is None or sales.empty:
        return pd.DataFrame(columns=["거래일자", "관리코드", "채널군", "채널",
                                     "수량", "판매금액", "판매이익"])
    d = sales.copy()
    d["거래일자"] = pd.to_datetime(d["거래일자"], errors="coerce")
    d = d[d["거래일자"].notna()]
    d = d[(d["거래일자"] >= pd.Timestamp(start)) & (d["거래일자"] <= pd.Timestamp(end))]
    d["관리코드"] = d["관리코드"].map(_nfc)
    if codes is not None:
        d = d[d["관리코드"].isin(codes)]
    cls = d["상호명"].map(classify_store)
    d["채널군"] = [c[0] for c in cls]
    d["채널"] = [c[1] for c in cls]
    d = d[d["채널군"] != GRP_SKIP]
    for c in ("수량", "판매금액", "판매이익"):
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
    return d


def channel_pivot(scoped: pd.DataFrame, value: str = "판매금액") -> pd.DataFrame:
    """관리코드 × 채널 피벗. 오픈마켓은 하위 채널, 나머지는 채널군 이름으로."""
    if scoped.empty:
        return pd.DataFrame()
    d = scoped.copy()
    d["열"] = np.where(d["채널군"] == GRP_OPEN, d["채널"], d["채널군"])
    p = d.pivot_table(index="관리코드", columns="열", values=value,
                      aggfunc="sum", fill_value=0.0)
    for c in channel_columns():
        if c not in p.columns:
            p[c] = 0.0
    return p[channel_columns()]


def depletion_sets(scoped: pd.DataFrame, end, days: int) -> pd.Series:
    """최근 N일 일평균 판매 **세트수**. 구 시트 W열(금액÷7)의 단위 오류를 바로잡은 것."""
    if scoped.empty:
        return pd.Series(dtype=float)
    end = pd.Timestamp(end)
    lo = end - pd.Timedelta(days=days - 1)
    w = scoped[scoped["거래일자"] >= lo]
    return w.groupby("관리코드")["수량"].sum() / float(days)


# ── 판정 ─────────────────────────────────────────────────────────────────────
def band_for(cfg: dict, D: int) -> tuple[float, float] | None:
    """D-day 진도 밴드. 정의된 D 중 현재 D 이상인 것 중 가장 가까운 값을 쓴다."""
    band = {int(k): v for k, v in (cfg.get("밴드") or {}).items()}
    if not band:
        return None
    cand = [d for d in band if d >= D]
    key = min(cand) if cand else max(band)
    lo, hi = band[key]
    return float(lo), float(hi)


def judge_row(sets_left: float, daily: float, days_left: int,
              yoy: float | None, has_sale: bool, season_behind: bool = True) -> str:
    """§6 4분면 — 소진예측(재고) × 진도(YoY 대용).

    재고 과잉 = 잔여일 안에 못 뺀다(소진예측일 > 잔여일).
    진도 미달 = 작년 동기 대비 100% 미만. ★작년 실적이 없는 신규 코드는 YoY가 없으므로
    '미달'로 단정하지 않고 **시즌 전체 밴드 판정을 승계**한다(season_behind).
    재고 0 이하 = 팔 게 없음 → 재고부족 쪽.
    """
    if not has_sale:
        return V_NOSALE
    left = max(float(sets_left), 0.0)
    over = False if left <= 0 else (True if daily <= 0 else (left / daily) > days_left)
    behind = season_behind if (yoy is None or pd.isna(yoy)) else (yoy < 1.0)
    if over and behind:
        return V_CUT
    if (not over) and behind:
        return V_BUY
    if over and (not behind):
        return V_HOLD
    return V_UP


def build_board(pm, attrs, lineup, sales, prev_sales, cfg, now=None) -> dict:
    """관제판 한 판. 반환 = {'board': DataFrame, 'win':…, 'kpi':…, 'group':…}"""
    win = season_window(cfg, now)
    ly = ly_aligned_dates(win)

    cur_all = scope_sales(sales, win["start"], win["end"])
    uni = build_universe(pm, attrs, lineup, set(cur_all["관리코드"]) if len(cur_all) else set())
    if uni.empty:
        return {"board": pd.DataFrame(), "win": win, "kpi": {},
                "group": pd.DataFrame(), "이익피벗": pd.DataFrame(),
                "scoped": pd.DataFrame()}
    codes = set(uni["관리코드"])

    cur = cur_all[cur_all["관리코드"].isin(codes)]
    prev = scope_sales(prev_sales, win["ly_start"], win["ly_end"], codes)
    prev_cal = prev[prev["거래일자"] <= ly["달력"]]
    prev_bd = prev[prev["거래일자"] <= ly["영업일"]]

    amt = channel_pivot(cur, "판매금액")
    prof = channel_pivot(cur, "판매이익")

    def _sum(d, col="판매금액"):
        return d.groupby("관리코드")[col].sum() if len(d) else pd.Series(dtype=float)

    b = uni.copy()
    b["시즌매출"] = b["관리코드"].map(_sum(cur)).fillna(0.0)
    b["시즌이익"] = b["관리코드"].map(_sum(cur, "판매이익")).fillna(0.0)
    b["판매세트"] = b["관리코드"].map(_sum(cur, "수량")).fillna(0.0)
    b["마진율"] = np.where(b["시즌매출"] > 0, b["시즌이익"] / b["시즌매출"], np.nan)

    # 채널군 소계 — ★'무매출'은 온라인 기준으로 봐야 한다(오프라인/나들로만 나가는 코드가 흔하다)
    for g, label in [(GRP_OPEN, "오픈마켓계"), (GRP_NADL, "나들계"), (GRP_OFF, "오프라인계")]:
        b[label] = b["관리코드"].map(_sum(cur[cur["채널군"] == g])).fillna(0.0)
    b["온라인계"] = b["관리코드"].map(
        _sum(cur[cur["채널군"].isin(ONLINE_GROUPS)])).fillna(0.0)

    d7 = depletion_sets(cur, win["end"], 7)
    d30 = depletion_sets(cur, win["end"], 30)
    b["일평균세트7"] = b["관리코드"].map(d7).fillna(0.0)
    b["일평균세트30"] = b["관리코드"].map(d30).fillna(0.0)
    left = b["세트재고"].clip(lower=0)
    b["소진예측일"] = np.where(b["일평균세트7"] > 0,
                            left / b["일평균세트7"].replace(0, np.nan),
                            np.where(left > 0, np.inf, 0.0))
    b["재고음수"] = b["박스재고"] < 0     # CJ 라인 매입전표 지연 — 출고 차단 사유 아님
    b["잔여시즌일"] = win["잔여일"]
    b["마감후잔여세트"] = np.maximum(left - b["일평균세트7"] * win["잔여일"], 0.0)
    b["이월재고금액"] = b["마감후잔여세트"] * b["세트원가"]

    b["작년동기_달력"] = b["관리코드"].map(_sum(prev_cal)).fillna(0.0)
    b["작년동기_영업일"] = b["관리코드"].map(_sum(prev_bd)).fillna(0.0)
    b["작년시즌"] = b["관리코드"].map(_sum(prev)).fillna(0.0)
    b["YoY"] = np.where(b["작년동기_달력"] > 0,
                        b["시즌매출"] / b["작년동기_달력"].replace(0, np.nan), np.nan)
    b["작년진도"] = np.where(b["작년시즌"] > 0,
                          b["시즌매출"] / b["작년시즌"].replace(0, np.nan), np.nan)

    for c in channel_columns():
        b[c] = b["관리코드"].map(amt[c] if c in amt.columns else pd.Series(dtype=float)).fillna(0.0)

    # ── KPI (오픈마켓 = 목표 스코프)
    om = cur[cur["채널군"] == GRP_OPEN]
    om_amt, om_prof = om["판매금액"].sum(), om["판매이익"].sum()
    target = float(cfg["매출목표"])
    band = band_for(cfg, win["D"])
    prog = om_amt / target if target else 0.0
    kpi = {
        "D": win["D"], "잔여일": win["잔여일"],
        "오픈마켓매출": om_amt, "오픈마켓이익": om_prof,
        "마진": om_prof / om_amt if om_amt else 0.0,
        "목표": target, "진도": prog, "밴드": band,
        "밴드판정": ("🟢" if band and band[0] <= prog <= band[1]
                  else ("🔴" if band and prog < band[0] else "🔵")),
        "재원": om_amt * (float(cfg["계획마진"]) - float(cfg["마진하한"])),
        "재원잔액": om_prof - om_amt * float(cfg["마진하한"]),
        "적립": om_prof - om_amt * float(cfg["계획마진"]),
        "작년달력": ly["달력"], "작년영업일": ly["영업일"],
    }
    pa_om = prev[prev["채널군"] == GRP_OPEN]
    for lab, cut in [("YoY달력", ly["달력"]), ("YoY영업일", ly["영업일"])]:
        base = pa_om[pa_om["거래일자"] <= cut]["판매금액"].sum()
        kpi[lab] = om_amt / base if base else np.nan

    behind = kpi["밴드판정"] == "🔴"
    b["판정"] = [
        judge_row(r["세트재고"], r["일평균세트7"], win["잔여일"],
                  r["YoY"], r["시즌매출"] > 0, season_behind=behind)
        for _, r in b.iterrows()
    ]

    grp = cur.groupby("채널군").agg(매출=("판매금액", "sum"), 이익=("판매이익", "sum"))
    grp["마진"] = np.where(grp["매출"] > 0, grp["이익"] / grp["매출"], np.nan)
    return {"board": b, "win": win, "kpi": kpi, "group": grp.reset_index(),
            "이익피벗": prof, "scoped": cur}
