"""두뇌④ 기준마진율 최적화 (ADR 0026) — 순수함수. 페이지가 데이터 주입.

입력 = compute_online_margin 결과(prod): 관리코드·채널(상호명 라벨)·거래일자·수량·판매금액·_순(순이익=판매이익−택배).
마진 정의 = **순이익/정산액(=판매금액)** — 채널마진모니터 기준마진율 시트와 동일(분모 정산액·택배 차감).

핵심(workflows/margin-optimizer.md):
- 베이스 = 순이익 누적 PROVEN_CUM 차지 proven 채널의 순이익 가중평균 마진.
- 채널 행동 = 볼륨 ×(마진 vs 베이스): proven=유지(베이스로 절반스텝) / 저볼륨·고마진=베이스로 절반스텝↓ / 저볼륨·저마진=hold-low.
- 노브: 절반스텝 / 게이트 🟢🟡🔴 / 관망=비중<IGNORE_SHARE 저우선 실험큐.
- 탄력성 신호 = 마진 흔든 이력(월별 마진 변동) 있으면 신뢰, 없으면 실험.
"""
from __future__ import annotations

import unicodedata

import numpy as np
import pandas as pd

PROVEN_CUM = 0.85       # 순이익 누적 비중 — proven 채널 경계
IGNORE_SHARE = 0.01     # 순이익 비중 < 1% = 저우선 실험큐(관망 후보)
HALF_STEP = 0.5         # 절반스텝(노브1)
BIG_MOVE = 0.03         # |Δ| > 3%p = 🔴 사람 판단 필수(노브3)
MIN_DELTA = 0.003       # |Δ| < 0.3%p = 미세 → 유지(작업목록서 제외)
MIN_MONTHS = 6          # 신호 판단 최소 개월
FLAT_STD = 0.015        # 월 마진 표준편차 < 1.5%p = "거의 안 흔듦" = 신호 없음


def _nfc(s) -> str:
    if s is None:
        return ""
    t = str(s)
    return "" if t == "nan" else unicodedata.normalize("NFC", t).strip()


def cell_stats(prod: pd.DataFrame) -> pd.DataFrame:
    """(관리코드, 채널) 셀 통계. prod = compute_online_margin 결과(_순 보유).

    return: 관리코드·채널·상품명·매출·순이익·수량·순마진·개월·마진std·corr(월마진↔월수량).
    """
    df = prod.copy()
    df = df[df["관리코드"].astype(str).map(_nfc) != "00-12"]
    df["_code"] = df["관리코드"].astype(str).map(_nfc)
    df["_ym"] = pd.to_datetime(df["거래일자"]).dt.strftime("%Y-%m")
    # 월별 (코드,채널) 집계 → 월 마진/수량
    m = (df.groupby(["_code", "채널", "_ym"], observed=True)
           .agg(매출=("판매금액", "sum"), 순=("_순", "sum"), 수량=("수량", "sum"),
                상품명=("상품명", "first")).reset_index())
    m = m[m["매출"] > 0]
    m["월마진"] = m["순"] / m["매출"]
    # 셀 집계
    rows = []
    for (code, ch), g in m.groupby(["_code", "채널"], observed=True):
        매출 = g["매출"].sum(); 순 = g["순"].sum(); 수량 = g["수량"].sum()
        if 매출 <= 0:
            continue
        std = float(g["월마진"].std(ddof=0)) if len(g) > 1 else 0.0
        corr = (float(np.corrcoef(g["월마진"], g["수량"])[0, 1])
                if len(g) >= MIN_MONTHS and g["월마진"].std() > 0 and g["수량"].std() > 0
                else float("nan"))
        개월 = int(g["_ym"].nunique())
        rows.append(dict(관리코드=code, 채널=ch, 상품명=g["상품명"].iloc[0],
                         매출=매출, 순이익=순, 월순이익=순 / max(개월, 1),
                         수량=int(수량), 순마진=순 / 매출,
                         개월=개월, 마진std=std, corr=corr))
    return pd.DataFrame(rows)


def base_margin(cells: pd.DataFrame) -> tuple[float, list[str]]:
    """proven 채널 = 월순이익(run-rate) 누적 PROVEN_CUM 까지(경계 넘는 채널 포함).
    개월 편차 보정 위해 총 순이익이 아니라 월순이익으로 랭크/가중. return (base, proven채널들)."""
    w = "월순이익"
    c = cells.sort_values(w, ascending=False).copy()
    c = c[c[w] > 0]
    if c.empty:
        tot = cells["순이익"].sum()
        if tot == 0:
            return float(cells["순마진"].mean() if len(cells) else 0.0), []
        return float((cells["순마진"] * cells["순이익"]).sum() / tot), []
    share = c[w] / c[w].sum()
    prev_cum = share.cumsum() - share          # 자기 직전까지 누적
    proven = c[prev_cum < PROVEN_CUM]           # 경계 넘는 채널까지 포함
    if proven.empty:
        proven = c.head(1)
    base = float((proven["순마진"] * proven[w]).sum() / proven[w].sum())
    return base, list(proven["채널"])


def recommend_code(cells_code: pd.DataFrame) -> pd.DataFrame:
    """한 관리코드의 채널별 권장 기준마진율 + 플래그 + 사유.

    cells_code = cell_stats 중 한 관리코드. return 채널별 추천 행.
    """
    cells_code = cells_code.copy()
    base, proven = base_margin(cells_code)
    tot_profit = cells_code["월순이익"].clip(lower=0).sum()
    out = []
    for _, r in cells_code.iterrows():
        ch, m, vol = r["채널"], r["순마진"], r["수량"]
        share = (max(r["월순이익"], 0) / tot_profit) if tot_profit > 0 else 0.0
        is_proven = ch in proven
        signal_weak = (r["개월"] < MIN_MONTHS) or (r["마진std"] < FLAT_STD)
        action, target, reason, flag = "유지", m, "", "🟢"

        if share < IGNORE_SHARE and not is_proven:
            # 관망(저우선 실험큐) — 영구배제 아님, 테스트 큐
            action, target = "실험큐", m
            reason = f"순이익 비중 {share*100:.1f}% (저우선) — 가격 테스트 큐, 무반응 2회 시 park"
            flag = "⚪"
        elif is_proven:
            # 베이스 정의·유지. 베이스보다 낮으면 절반스텝 올림(가격둔감 여지)
            if m < base - 0.005:
                target = m + (base - m) * HALF_STEP
                action, reason, flag = "↑ 절반스텝", "proven인데 베이스 미만 → 절반스텝 상향(가격둔감 여지)", "🟢"
            else:
                action, target, reason, flag = "유지", m, "proven(베이스 정의·적정)", "🟢"
        elif m > base + 0.005:
            # 저볼륨·고마진 → 베이스로 절반스텝 인하
            target = m - (m - base) * HALF_STEP
            action, reason = "↓ 절반스텝", f"저볼륨·고마진(>베이스 {base*100:.1f}%) → 절반스텝 인하 테스트"
            flag = "🟡"
            if signal_weak:
                flag, reason = "🔴", reason + " · 마진 흔든 이력 약함=실험 성격"
        elif share >= IGNORE_SHARE * 10:
            # 비proven이지만 큰 채널(비중≥10%) 베이스 미만 → 상향 여지(가격둔감)
            target = m + (base - m) * HALF_STEP
            action, reason, flag = "↑ 절반스텝", f"주력급(비중 {share*100:.0f}%)인데 베이스 미만 → 절반스텝 상향", "🟡"
        else:
            # 저볼륨·저마진(<베이스) → hold-low
            action, target = "hold-low", m
            reason = f"이미 싼데 안 팔림(<베이스 {base*100:.1f}%) → 낮게 유지·안 올림(팔리면 재평가)"
            flag = "🔴"

        delta = target - m
        if action in ("↑ 절반스텝", "↓ 절반스텝") and abs(delta) < MIN_DELTA:
            action, target, delta, flag = "유지", m, 0.0, "🟢"
            reason = "베이스와 거의 일치(미세) → 유지"
        if abs(delta) > BIG_MOVE:
            flag = "🔴"
        if (action.startswith("↑") or action.startswith("↓")) and signal_weak and flag != "🔴":
            flag = "🟡"

        out.append(dict(
            관리코드=r["관리코드"], 상품명=r["상품명"], 채널=ch,
            현재마진=round(m * 100, 1), 베이스=round(base * 100, 1),
            권장마진=round(target * 100, 1), Δ=round(delta * 100, 1),
            월순이익=int(round(r["순이익"] / max(r["개월"], 1))),
            월볼륨=int(round(vol / max(r["개월"], 1))),
            수량=int(vol), 액션=action, 플래그=flag, 사유=reason))
    df = pd.DataFrame(out)
    # 임팩트(월순이익) 순
    return df.sort_values("월순이익", ascending=False).reset_index(drop=True)


def worklist(cells: pd.DataFrame, codes=None) -> pd.DataFrame:
    """여러 관리코드 → 전체 작업목록. 유지/실험큐 제외하고 '손볼 것'만(플래그 있는 변경)."""
    res = []
    grp = cells.groupby("관리코드", observed=True)
    for code, g in grp:
        if codes is not None and code not in codes:
            continue
        res.append(recommend_code(g))
    if not res:
        return pd.DataFrame()
    df = pd.concat(res, ignore_index=True)
    return df.sort_values("월순이익", ascending=False).reset_index(drop=True)


# ── 측정 루프 (Gate 3 닫기) ───────────────────────────────────────────
MEASURE_MIN_DAYS = 30   # 결정 후 측정 대상 최소 경과일(사용자 확정: 빠른 피드백)
RESP_BAND = 0.10        # ±10% 월순이익 변화 밴드(밖 = 개선/악화)


def _verdict(action, before, after, ready: bool) -> tuple[str, str]:
    """월순이익 before→after 로 결과·제안. action=기록 당시 액션."""
    if (not ready) or after is None or (isinstance(after, float) and pd.isna(after)):
        return "측정대기", ""
    before = float(before or 0.0)
    after = float(after)
    if before > 0:
        hi, lo = before * (1 + RESP_BAND), before * (1 - RESP_BAND)
        v = "개선" if after > hi else ("악화" if after < lo else "무변화")
    else:
        d = after - before
        thr = max(abs(before) * RESP_BAND, 1000.0)
        v = "개선" if d > thr else ("악화" if d < -thr else "무변화")
    if v == "개선":
        s = "유지" + (" · 추가 인하 후보" if str(action).startswith("↓") else "")
    elif v == "악화":
        s = "되돌림"
    else:
        s = "유지(관찰)"
    return v, s


def measure(prod: pd.DataFrame, pending: pd.DataFrame, today=None,
            min_days: int = MEASURE_MIN_DAYS) -> pd.DataFrame:
    """pending 결정들의 *결정일(ts) 이후* 실적으로 측정후 월 run-rate + 결과 판정.

    prod = compute_online_margin 결과(거래일자·수량·판매금액·_순·관리코드·채널·상품명).
    pending = 원장 status=pending 행(decision_id·관리코드·채널·ts·액션·마진_before·마진_적용·
              측정전_월볼륨·측정전_월순이익·베이스·플래그).
    return 행별: 측정전/측정후(월순이익·월볼륨)·측정후마진·측정일수·post_개월·경과일·ready·결과·제안.
    ★ready·run-rate는 **적재 데이터 커버리지** 기준(벽시계 아님): 월별 적재 갭 편향 방지.
      측정일수 = 적재 최신거래일(data_max) − ts. ready = 측정일수 ≥ min_days.
      월순/월볼 = ts 이후 거래 합 ÷ (측정일수/30.4) — **일수 정규화**(부분월·월말 결정 편향 제거).
      동일 마진정의(순/매출). post 매출 0이어도 ready면 월순=0(=판매 붕괴 신호).
    """
    import datetime as _dt
    today = today or _dt.date.today()
    cols_out = ["decision_id", "관리코드", "상품명", "채널", "액션", "ts", "경과일",
                "측정일수", "post_개월", "측정전_월순이익", "측정후_월순이익", "측정전_월볼륨",
                "측정후_월볼륨", "측정후마진", "마진_before", "마진_적용", "베이스",
                "플래그", "ready", "결과", "제안"]
    if pending is None or len(pending) == 0 or prod is None or prod.empty:
        return pd.DataFrame(columns=cols_out)
    df = prod.copy()
    df["_code"] = df["관리코드"].astype(str).map(_nfc)
    df["_ch"] = df["채널"].astype(str).map(_nfc)
    df["_dt"] = pd.to_datetime(df["거래일자"])
    df["_ym"] = df["_dt"].dt.strftime("%Y-%m")
    _dmax = df["_dt"].max()
    data_max_d = _dmax.date() if pd.notna(_dmax) else today  # 적재 프론티어(전 셀 공통)
    name_by = (df.dropna(subset=["상품명"]).groupby("_code")["상품명"].first().to_dict()
               if "상품명" in df.columns else {})
    out = []
    for _, r in pending.iterrows():
        code = _nfc(r.get("관리코드"))
        ch = _nfc(r.get("채널"))
        ts = str(r.get("ts") or "")
        try:
            ts_d = _dt.date.fromisoformat(ts[:10])
        except Exception:
            ts_d = today
        elapsed = (today - ts_d).days          # 벽시계(표시용)
        cov_days = (data_max_d - ts_d).days     # 적재된 post 기간(측정 기준)
        ready = cov_days >= min_days
        sub = df[(df["_code"] == code) & (df["_ch"] == ch)
                 & (df["_dt"] > pd.Timestamp(ts_d))]
        post_ym = int(sub["_ym"].nunique())
        if ready:
            매출 = float(sub["판매금액"].sum())
            순 = float(sub["_순"].sum())
            수량 = float(sub["수량"].sum())
            months = max(cov_days / 30.4, 1e-9)   # 일수 정규화(달력월 개수 아님)
            월순 = 순 / months
            월볼 = 수량 / months
            마진 = (순 / 매출) if 매출 > 0 else float("nan")
            after = 월순
        else:
            월순 = 월볼 = 마진 = float("nan")
            after = None
        before = float(r.get("측정전_월순이익") or 0)
        v, s = _verdict(r.get("액션"), before, after, ready)
        out.append(dict(
            decision_id=r.get("decision_id"), 관리코드=r.get("관리코드"),
            상품명=name_by.get(code, ""), 채널=r.get("채널"), 액션=r.get("액션"),
            ts=ts[:10], 경과일=elapsed, 측정일수=max(cov_days, 0), post_개월=post_ym,
            측정전_월순이익=int(round(before)),
            측정후_월순이익=(int(round(월순)) if ready else None),
            측정전_월볼륨=int(r.get("측정전_월볼륨") or 0),
            측정후_월볼륨=(int(round(월볼)) if ready else None),
            측정후마진=(round(마진 * 100, 1) if (ready and pd.notna(마진)) else None),
            마진_before=r.get("마진_before"), 마진_적용=r.get("마진_적용"),
            베이스=r.get("베이스"), 플래그=r.get("플래그"),
            ready=ready, 결과=v, 제안=s))
    return pd.DataFrame(out, columns=cols_out)
