"""data backbone 커버리지 — 시계열로 누적 적립되는 데이터의 적재 현황(범위·갭) 산출.

이력 엔진(ADR 0018). work-automation-data private repo 대상.
- monthly: 디렉토리 목록으로 월 범위·갭 산출 + 첫/마지막 파티션 1~2개만 read해
           실제 거래일자 min/max(일자까지) 산출.
- single: 단일파일 1개만 read해 date_col min/max로 범위 산출(가벼움).

분류 기준 = "시계열로 누적되느냐". 상품관리(product_master)·채널 listing 등
'현재 상태 덮어쓰기' 자료는 카탈로그에서 제외.

upload: direct(페이지서 업로드) · auto(부산물 자동적립·현황만) · planned(예정·미구현)
kind:   monthly(월 파티션) · single(단일파일·이벤트 누적)
범위 표시: first/last = 'YYYY-MM'(갭·타임라인용) · first_day/last_day = 'YYYY-MM-DD'(표시용)
"""
from __future__ import annotations
import io, json, re, urllib.request, urllib.error

import pandas as pd

CATALOG = [
    {"id": "sales", "label": "매출", "dir": "master", "rx": r"sales_(\d{4}-\d{2})\.parquet",
     "date_col": "거래일자",
     "kind": "monthly", "upload": "direct", "note": "천년경영 매출 · 정산 진실원천"},
    {"id": "orders", "label": "주문", "dir": "orders", "rx": r"easyadmin_(\d{4}-\d{2})\.parquet",
     "date_col": "발주일",
     "kind": "monthly", "upload": "direct", "note": "EasyAdmin 확장주문검색 · velocity·송장그룹"},
    {"id": "price", "label": "가격이력", "dir": "history", "file": "price_changes.parquet",
     "date_col": "수정일자", "kind": "single", "upload": "direct", "note": "상품수정삭제로그 · 1년 롤링 누적"},
    {"id": "stock", "label": "재고 스냅샷", "dir": "snapshots", "rx": r"stock_(\d{4}-\d{2})\.parquet",
     "date_col": "스냅샷일자",
     "kind": "monthly", "upload": "auto", "note": "상품관리 업로드 부산물 · 자동 적립"},
    {"id": "purchases", "label": "매입현황", "dir": "purchases", "rx": r"buyin_(\d{4}-\d{2})\.parquet",
     "date_col": "기준일",
     "kind": "monthly", "upload": "planned", "note": "유형별매입현황 · 실입고(예정)"},
    {"id": "demand", "label": "발주자료", "dir": "demand", "rx": r"order_(\d{4}-\d{2})\.parquet",
     "kind": "monthly", "upload": "planned", "note": "logistics 발주 · 리드타임 산출(예정)"},
]


def _api(url, pat, accept="application/vnd.github+json"):
    return urllib.request.urlopen(urllib.request.Request(
        url, headers={"Authorization": "Bearer " + pat, "User-Agent": "wa-app", "Accept": accept}))


def _list_dir(pat, repo, d):
    url = "https://api.github.com/repos/%s/contents/%s?ref=main" % (repo, d)
    try:
        return json.load(_api(url, pat))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return []
        raise


def _file_bytes(pat, repo, path):
    url = "https://api.github.com/repos/%s/contents/%s?ref=main" % (repo, path)
    return _api(url, pat, accept="application/vnd.github.raw").read()


def _day_range(pat, repo, path, date_col):
    """파티션 1개를 읽어 date_col의 (min, max) 일자 'YYYY-MM-DD'. 실패/공란이면 (None, None)."""
    try:
        b = _file_bytes(pat, repo, path)
        s = pd.to_datetime(pd.read_parquet(io.BytesIO(b), columns=[date_col])[date_col],
                           errors="coerce").dropna()
        if len(s):
            return s.min().strftime("%Y-%m-%d"), s.max().strftime("%Y-%m-%d")
    except Exception:
        pass
    return None, None


def _msnum(s):
    y, m = map(int, s.split("-"))
    return y * 12 + m


def _mslbl(v):
    return "%d-%02d" % ((v - 1) // 12, (v - 1) % 12 + 1)


def next_month(ym):
    """'2026-05' → '2026-06' (타임라인 막대 끝)."""
    return _mslbl(_msnum(ym) + 1)


def coverage(pat, repo):
    """카탈로그별 적재 현황 리스트.

    각: id·label·kind·upload·note·status·files·size_kb·first·last·gaps + first_day·last_day.
    first/last = 'YYYY-MM'(갭·타임라인용), first_day/last_day = 'YYYY-MM-DD'(표시용·없으면 None).
    """
    out = []
    for c in CATALOG:
        items = _list_dir(pat, repo, c["dir"])
        rec = dict(c)
        rec["first"] = rec["last"] = None
        rec["first_day"] = rec["last_day"] = None
        rec["gaps"] = []
        if c["kind"] == "monthly":
            name_by_month = {}
            pairs = []
            for it in items:
                m = re.match(c["rx"], it.get("name", ""))
                if m:
                    pairs.append((m.group(1), it.get("size", 0)))
                    name_by_month[m.group(1)] = it.get("name", "")
            months = sorted(p[0] for p in pairs)
            rec["files"] = len(months)
            rec["size_kb"] = round(sum(p[1] for p in pairs) / 1024)
            if months:
                lo, hi = _msnum(months[0]), _msnum(months[-1])
                have = {_msnum(x) for x in months}
                rec["first"], rec["last"] = months[0], months[-1]
                rec["gaps"] = [_mslbl(v) for v in range(lo, hi + 1) if v not in have]
                rec["status"] = "ok"
                # 일자까지: 첫 달·마지막 달 파티션만 read해 실제 거래일 min/max
                dc = c.get("date_col")
                if dc:
                    if months[0] == months[-1]:
                        mn, mx = _day_range(pat, repo, "%s/%s" % (c["dir"], name_by_month[months[0]]), dc)
                    else:
                        mn, _ = _day_range(pat, repo, "%s/%s" % (c["dir"], name_by_month[months[0]]), dc)
                        _, mx = _day_range(pat, repo, "%s/%s" % (c["dir"], name_by_month[months[-1]]), dc)
                    rec["first_day"], rec["last_day"] = mn, mx
            else:
                rec["status"] = "empty"
        else:  # single — 파일 1개 read해 date_col 범위
            f = next((it for it in items if it.get("name") == c.get("file")), None)
            rec["files"] = 1 if f else 0
            rec["size_kb"] = round(f.get("size", 0) / 1024) if f else 0
            rec["status"] = "ok" if f else "empty"
            if f and c.get("date_col"):
                try:
                    b = _file_bytes(pat, repo, "%s/%s" % (c["dir"], c["file"]))
                    s = pd.to_datetime(pd.read_parquet(io.BytesIO(b))[c["date_col"]], errors="coerce").dropna()
                    if len(s):
                        rec["first"] = s.min().strftime("%Y-%m")
                        rec["last"] = s.max().strftime("%Y-%m")
                        rec["first_day"] = s.min().strftime("%Y-%m-%d")
                        rec["last_day"] = s.max().strftime("%Y-%m-%d")
                except Exception:
                    pass
        out.append(rec)
    return out
