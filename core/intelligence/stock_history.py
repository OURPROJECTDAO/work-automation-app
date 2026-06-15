"""상품관리 시트 → 재고·매입가 일자 스냅샷(이력). 적립 + 입고/품절 전이 탐지.

이력 엔진 Phase 1 1b (ADR 0018, intelligence-layer.md §3.1·§5 1b).
- 저장: private repo work-automation-data : snapshots/stock_YYYY-MM.parquet (월 파티션).
- 매일 상품관리 업로드가 product_master.csv를 덮어써 역사 소실 → 덮어쓰기 직전 새 업로드 df를 날짜본으로 적립.
- dedup 키 = (스냅샷일자, 상품코드). 같은 날 재업로드 멱등(keep=last).
- 스냅샷일자 = 파일명 Exp{YYMMDD} 추출일(없으면 호출자가 당일 폴백).
- 상품코드 = 정본 키(앞자리0 보존 6자리). PII 없음(재고·가격만).
- 전이: 박스재고 0(이하)↔양수 = 품절/입고 이벤트(두뇌② 입력).

상품관리 컬럼(위치 인덱스, 실데이터 대조 2026-06-15):
  [3]상품코드 [4]관리코드 [8]매입단가(낱개) [9]박스매입단가 [10]매출단가 [12]매익률 [14]박스재고
"""
from __future__ import annotations

import base64
import io
import json
import re
import unicodedata
import urllib.error
import urllib.request
import datetime as _dt

import pandas as pd

PART_DIR = "snapshots"
_PART_RE = re.compile(r"^stock_(\d{4}-\d{2})\.parquet$")
_EXP_RE = re.compile(r"Exp(\d{2})(\d{2})(\d{2})")  # Exp{YY}{MM}{DD}
KEY = ["스냅샷일자", "상품코드"]
COLS = ["스냅샷일자", "상품코드", "관리코드", "박스재고",
        "박스매입단가", "매입단가", "매익률", "매출단가"]

# 상품관리 위치 인덱스
_C_CODE, _C_MGMT, _C_PURCH, _C_BOXPURCH, _C_SELL, _C_MARGIN, _C_BOXSTOCK = 3, 4, 8, 9, 10, 12, 14


def _nfc(v):
    return unicodedata.normalize("NFC", str(v)).strip() if v is not None else ""


def _code6(v):
    c = _nfc(v).split(".")[0]
    return c.zfill(6) if c.isdigit() else _nfc(v)


def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_snapshot_date(filename: str):
    """파일명 'Exp{YYMMDD}_..._상품관리_.xlsx' → date. 없으면 None(호출자 당일 폴백)."""
    if not filename:
        return None
    m = _EXP_RE.search(filename)
    if not m:
        return None
    yy, mm, dd = (int(x) for x in m.groups())
    try:
        return _dt.date(2000 + yy, mm, dd)
    except ValueError:
        return None


def snapshot_from_master(df: pd.DataFrame, snap_date) -> pd.DataFrame:
    """상품관리 업로드 df(위치 인덱스) → 스냅샷 7컬럼 정규화. snap_date = date/datetime."""
    ts = pd.Timestamp(snap_date).normalize()
    n = df.shape[1]
    if n <= _C_BOXSTOCK:
        raise ValueError(f"상품관리 컬럼 부족: {n}열 (≥{_C_BOXSTOCK + 1} 필요)")
    out = []
    for r in df.itertuples(index=False):
        code = _code6(r[_C_CODE])
        if not code:
            continue
        out.append({
            "스냅샷일자": ts,
            "상품코드": code,
            "관리코드": _nfc(r[_C_MGMT]),
            "박스재고": _num(r[_C_BOXSTOCK]),
            "박스매입단가": _num(r[_C_BOXPURCH]),
            "매입단가": _num(r[_C_PURCH]),
            "매익률": _num(r[_C_MARGIN]),
            "매출단가": _num(r[_C_SELL]),
        })
    snap = pd.DataFrame(out, columns=COLS)
    snap["스냅샷일자"] = pd.to_datetime(snap["스냅샷일자"])
    for c in ("상품코드", "관리코드"):
        snap[c] = snap[c].astype("string")
    # 같은 파일 내 상품코드 중복(이론상 없음) 방어
    return snap.drop_duplicates(KEY, keep="last").reset_index(drop=True)


# ---- GitHub parquet R/W (work-automation-data) — price_history/store 패턴 ----
def _gh(url, pat, method="GET", data=None, accept="application/vnd.github+json"):
    headers = {"Authorization": "Bearer " + pat, "User-Agent": "wa-app", "Accept": accept}
    if data is not None:
        headers["Content-Type"] = "application/json"
        data = json.dumps(data).encode("utf-8")
    return urllib.request.urlopen(urllib.request.Request(url, data=data, method=method, headers=headers))


def _part_path(ym: str) -> str:
    return "%s/stock_%s.parquet" % (PART_DIR, ym)


def read_snapshots(pat, repo, ym: str) -> pd.DataFrame:
    """월 파티션 stock_YYYY-MM.parquet → DataFrame. 없으면 빈 DataFrame."""
    url = "https://api.github.com/repos/%s/contents/%s" % (repo, _part_path(ym))
    try:
        with _gh(url + "?ref=main", pat, accept="application/vnd.github.raw") as r:
            return pd.read_parquet(io.BytesIO(r.read()))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return pd.DataFrame(columns=COLS)
        raise


def list_snapshot_months(pat, repo) -> list:
    """snapshots/ 디렉토리의 stock_YYYY-MM 목록(정렬)."""
    url = "https://api.github.com/repos/%s/contents/%s" % (repo, PART_DIR)
    try:
        items = json.load(_gh(url + "?ref=main", pat))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return []
        raise
    months = [m.group(1) for it in items for m in [_PART_RE.match(it.get("name", ""))] if m]
    return sorted(months)


def read_all_snapshots(pat, repo) -> pd.DataFrame:
    """전 월 파티션 합본(전이 탐지용 — 월경계 연속성 보존)."""
    parts = [read_snapshots(pat, repo, ym) for ym in list_snapshot_months(pat, repo)]
    parts = [p for p in parts if p is not None and len(p)]
    if not parts:
        return pd.DataFrame(columns=COLS)
    return (pd.concat(parts, ignore_index=True)
              .sort_values(["상품코드", "스냅샷일자"]).reset_index(drop=True))


def ingest_snapshot(snap_df: pd.DataFrame, pat, repo) -> dict:
    """스냅샷을 해당 월 파티션에 dedup-append 후 커밋. 멱등((스냅샷일자·상품코드) keep=last)."""
    if snap_df.empty:
        return {"rows": 0, "months": [], "added": 0}
    snap_df = snap_df.copy()
    snap_df["스냅샷일자"] = pd.to_datetime(snap_df["스냅샷일자"])
    touched, total_added = [], 0
    for ym, grp in snap_df.groupby(snap_df["스냅샷일자"].dt.strftime("%Y-%m")):
        cur = read_snapshots(pat, repo, ym)
        before = len(cur)
        merged = pd.concat([cur, grp], ignore_index=True)
        merged["스냅샷일자"] = pd.to_datetime(merged["스냅샷일자"])
        merged = (merged.drop_duplicates(KEY, keep="last")
                        .sort_values(["상품코드", "스냅샷일자"]).reset_index(drop=True))
        buf = io.BytesIO()
        merged.to_parquet(buf, index=False)
        content = base64.b64encode(buf.getvalue()).decode("ascii")
        url = "https://api.github.com/repos/%s/contents/%s" % (repo, _part_path(ym))
        sha = None
        try:
            sha = json.load(_gh(url + "?ref=main", pat)).get("sha")
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
        body = {"message": "data: stock 스냅샷 적립 %s (+%d행)" % (ym, len(merged) - before),
                "content": content}
        if sha:
            body["sha"] = sha
        _gh(url, pat, method="PUT", data=body)
        touched.append(ym)
        total_added += len(merged) - before
    return {"rows": len(snap_df), "months": sorted(touched), "added": total_added}


# ---- 전이 탐지: 입고(0→양수) / 품절(양수→0) ----
def detect_transitions(snaps: pd.DataFrame, in_stock_threshold: float = 0.0) -> pd.DataFrame:
    """상품코드별 연속 스냅샷 박스재고 비교 → 전이 이벤트.

    입고 = 직전 ≤thr & 현재 >thr · 품절 = 직전 >thr & 현재 ≤thr.
    반환: 상품코드·관리코드·전일·금일·전일재고·금일재고·전이(입고/품절).
    """
    if snaps.empty:
        return pd.DataFrame(columns=["상품코드", "관리코드", "전일", "금일",
                                     "전일재고", "금일재고", "전이"])
    s = snaps.sort_values(["상품코드", "스냅샷일자"]).copy()
    s["박스재고"] = pd.to_numeric(s["박스재고"], errors="coerce").fillna(0.0)
    s["_prev재고"] = s.groupby("상품코드")["박스재고"].shift(1)
    s["_prev일"] = s.groupby("상품코드")["스냅샷일자"].shift(1)
    thr = in_stock_threshold
    prev_in = s["_prev재고"] > thr
    cur_in = s["박스재고"] > thr
    is_restock = (~prev_in) & cur_in & s["_prev재고"].notna()
    is_stockout = prev_in & (~cur_in)
    ev = s[is_restock | is_stockout].copy()
    ev["전이"] = ev.apply(lambda r: "입고" if (r["_prev재고"] <= thr and r["박스재고"] > thr)
                          else "품절", axis=1)
    ev = ev.rename(columns={"_prev일": "전일", "스냅샷일자": "금일",
                            "_prev재고": "전일재고", "박스재고": "금일재고"})
    return (ev[["상품코드", "관리코드", "전일", "금일", "전일재고", "금일재고", "전이"]]
            .sort_values(["금일", "상품코드"]).reset_index(drop=True))
