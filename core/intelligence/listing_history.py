"""채널 listing(상품관리 스냅샷) → 채널별 가격 일자 스냅샷(이력). 적립 + 가격변경 탐지.

이력 엔진 Phase 1 1d (ADR 0018, intelligence-layer.md §5 1d).
- 저장: private repo work-automation-data : snapshots/listing_YYYY-MM.parquet (월 파티션).
- channel-margin-monitor '상품관리 갱신'이 reference/listing_<key>.csv 를 덮어써 가격 역사 소실
  → 갱신(커밋) 직후 그 채널 listing 가격을 날짜본으로 적립(stock_history 1b 동형).
- dedup 키 = (스냅샷일자, 채널, 상품번호). 같은 날 재갱신 멱등(keep=last).
- 스냅샷일자 = 갱신 당일(KST). listing 다운로드는 추출일시 보장 없음 → 갱신 시점.
- forward 축적(과거 소급 불가 — 현재 listing만 존재). PII 없음(가격만).
- 두뇌③ 채널 가격 A/B 가격변경 전후 = 이 이력에서 산출.

listing 레코드 키(cmm.parse_download / recs_to_csv):
  상품번호 · 코드(관리코드) · 상품명 · 판매가 · 정가 · 배송비 · 즉시할인 · 포인트 · 바코드 …
"""
from __future__ import annotations

import base64
import io
import json
import re
import unicodedata
import urllib.error
import urllib.request

import pandas as pd

PART_DIR = "snapshots"
_PART_RE = re.compile(r"^listing_(\d{4}-\d{2})\.parquet$")
KEY = ["스냅샷일자", "채널", "상품번호"]
COLS = ["스냅샷일자", "채널", "상품번호", "관리코드",
        "판매가", "정가", "배송비", "즉시할인", "포인트"]


def _nfc(v):
    return unicodedata.normalize("NFC", str(v)).strip() if v is not None else ""


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


def snapshot_from_recs(recs: list, channel: str, snap_date) -> pd.DataFrame:
    """cmm listing 레코드 리스트 → 가격 스냅샷 9컬럼 정규화. snap_date = date/datetime."""
    ts = pd.Timestamp(snap_date).normalize()
    ch = _nfc(channel)
    out = []
    for r in recs or []:
        pid = _nfc(r.get("상품번호"))
        if not pid:
            continue
        out.append({
            "스냅샷일자": ts,
            "채널": ch,
            "상품번호": pid,
            "관리코드": _nfc(r.get("코드")),
            "판매가": _num(r.get("판매가")),
            "정가": _num(r.get("정가")),
            "배송비": _num(r.get("배송비")),
            "즉시할인": _num(r.get("즉시할인")),
            "포인트": _num(r.get("포인트")),
        })
    snap = pd.DataFrame(out, columns=COLS)
    snap["스냅샷일자"] = pd.to_datetime(snap["스냅샷일자"])
    for c in ("채널", "상품번호", "관리코드"):
        snap[c] = snap[c].astype("string")
    # 같은 파일 내 (채널·상품번호) 중복(dedup_key 미처리 채널 방어)
    return snap.drop_duplicates(KEY, keep="last").reset_index(drop=True)


# ---- GitHub parquet R/W (work-automation-data) — stock_history 패턴 ----
def _gh(url, pat, method="GET", data=None, accept="application/vnd.github+json"):
    headers = {"Authorization": "Bearer " + pat, "User-Agent": "wa-app", "Accept": accept}
    if data is not None:
        headers["Content-Type"] = "application/json"
        data = json.dumps(data).encode("utf-8")
    return urllib.request.urlopen(urllib.request.Request(url, data=data, method=method, headers=headers))


def _part_path(ym: str) -> str:
    return "%s/listing_%s.parquet" % (PART_DIR, ym)


def read_listing_snapshots(pat, repo, ym: str) -> pd.DataFrame:
    """월 파티션 listing_YYYY-MM.parquet → DataFrame. 없으면 빈 DataFrame."""
    url = "https://api.github.com/repos/%s/contents/%s" % (repo, _part_path(ym))
    try:
        with _gh(url + "?ref=main", pat, accept="application/vnd.github.raw") as r:
            return pd.read_parquet(io.BytesIO(r.read()))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return pd.DataFrame(columns=COLS)
        raise


def list_listing_months(pat, repo) -> list:
    """snapshots/ 의 listing_YYYY-MM 목록(정렬)."""
    url = "https://api.github.com/repos/%s/contents/%s" % (repo, PART_DIR)
    try:
        items = json.load(_gh(url + "?ref=main", pat))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return []
        raise
    months = [m.group(1) for it in items for m in [_PART_RE.match(it.get("name", ""))] if m]
    return sorted(months)


def read_all_listing_snapshots(pat, repo) -> pd.DataFrame:
    """전 월 파티션 합본(가격변경 탐지 — 월경계 연속성 보존)."""
    parts = [read_listing_snapshots(pat, repo, ym) for ym in list_listing_months(pat, repo)]
    parts = [p for p in parts if p is not None and len(p)]
    if not parts:
        return pd.DataFrame(columns=COLS)
    return (pd.concat(parts, ignore_index=True)
              .sort_values(["채널", "상품번호", "스냅샷일자"]).reset_index(drop=True))


def ingest_listing_snapshot(snap_df: pd.DataFrame, pat, repo) -> dict:
    """스냅샷을 해당 월 파티션에 dedup-append 후 커밋. 멱등((스냅샷일자·채널·상품번호) keep=last)."""
    if snap_df is None or snap_df.empty:
        return {"rows": 0, "months": [], "added": 0}
    snap_df = snap_df.copy()
    snap_df["스냅샷일자"] = pd.to_datetime(snap_df["스냅샷일자"])
    touched, total_added = [], 0
    for ym, grp in snap_df.groupby(snap_df["스냅샷일자"].dt.strftime("%Y-%m")):
        cur = read_listing_snapshots(pat, repo, ym)
        before = len(cur)
        merged = pd.concat([cur, grp], ignore_index=True)
        merged["스냅샷일자"] = pd.to_datetime(merged["스냅샷일자"])
        merged = (merged.drop_duplicates(KEY, keep="last")
                        .sort_values(["채널", "상품번호", "스냅샷일자"]).reset_index(drop=True))
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
        body = {"message": "data: listing 가격 스냅샷 적립 %s (+%d행)" % (ym, len(merged) - before),
                "content": content}
        if sha:
            body["sha"] = sha
        _gh(url, pat, method="PUT", data=body)
        touched.append(ym)
        total_added += len(merged) - before
    return {"rows": len(snap_df), "months": sorted(touched), "added": total_added}


PRICE_FIELDS = {"판매가": "판매가", "정가": "정가"}


def detect_listing_price_changes(snaps: pd.DataFrame, threshold: float = 0.0,
                                 fields=("판매가", "정가")) -> pd.DataFrame:
    """채널×상품번호별 연속 스냅샷 가격 비교 → |변동률| >= threshold 변동 이벤트.

    stock_history.detect_price_changes 자매(키=채널+상품번호, listing 판매가/정가).
    두뇌③ 채널 가격 A/B 가격변경 전후 측정용. threshold 0 = 모든 변동.
    반환: 채널·상품번호·관리코드·구분(판매가/정가)·전일·금일·전일가·금일가·변동률(%)·방향.
    ★ 직전값 결측·직전값 0 이하 제외. forward 적립이라 적립 시작 이후만.
    """
    cols = ["채널", "상품번호", "관리코드", "구분", "전일", "금일",
            "전일가", "금일가", "변동률", "방향"]
    if snaps is None or snaps.empty:
        return pd.DataFrame(columns=cols)
    s = snaps.sort_values(["채널", "상품번호", "스냅샷일자"]).copy()
    s["스냅샷일자"] = pd.to_datetime(s["스냅샷일자"])
    grp = [s["채널"], s["상품번호"]]
    parts = []
    for field in fields:
        if field not in s.columns:
            continue
        v = pd.to_numeric(s[field], errors="coerce")
        prev_v = v.groupby(grp).shift(1)
        prev_d = s["스냅샷일자"].groupby(grp).shift(1)
        rate = (v - prev_v) / prev_v
        # threshold=0 = '모든 실변동'(0% 무변동 제외). threshold>0 = 그 비율 이상.
        mask = (prev_v.notna() & v.notna() & (prev_v > 0)
                & (rate != 0) & (rate.abs() >= threshold))
        if not mask.any():
            continue
        part = pd.DataFrame({
            "채널": s["채널"], "상품번호": s["상품번호"], "관리코드": s["관리코드"],
            "구분": PRICE_FIELDS.get(field, field),
            "전일": prev_d, "금일": s["스냅샷일자"],
            "전일가": prev_v, "금일가": v, "변동률": rate,
        })[mask]
        parts.append(part)
    if not parts:
        return pd.DataFrame(columns=cols)
    res = pd.concat(parts, ignore_index=True)
    res["방향"] = res["변동률"].map(lambda x: "인상" if x > 0 else "인하")
    res["변동률"] = (res["변동률"] * 100).round(2)
    return (res[cols].sort_values(["금일", "채널", "상품번호"], ascending=[False, True, True])
            .reset_index(drop=True))
