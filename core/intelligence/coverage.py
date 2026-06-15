"""data backbone 커버리지 — 시계열로 누적 적립되는 데이터의 적재 현황(범위·갭) 산출.

이력 엔진(ADR 0018). work-automation-data private repo의 디렉토리 목록만 사용
(파일 내용 read 0회 → 가볍고 캐싱 친화). 통합 데이터 현황 페이지가 호출.

분류 기준 = "시계열로 누적되느냐". 상품관리(product_master)·채널 listing 같은
'현재 상태 덮어쓰기' 자료는 카탈로그에서 제외(통합관리 대상 아님).

upload: direct(페이지서 업로드) · auto(부산물 자동적립·현황만) · planned(예정·미구현)
kind:   monthly(월 파티션) · single(단일파일·이벤트 누적)
"""
from __future__ import annotations
import json, re, urllib.request, urllib.error

CATALOG = [
    {"id": "sales", "label": "매출", "dir": "master", "rx": r"sales_(\d{4}-\d{2})\.parquet",
     "kind": "monthly", "upload": "direct", "note": "천년경영 매출 · 정산 진실원천"},
    {"id": "orders", "label": "주문", "dir": "orders", "rx": r"easyadmin_(\d{4}-\d{2})\.parquet",
     "kind": "monthly", "upload": "direct", "note": "EasyAdmin 확장주문검색 · velocity·송장그룹"},
    {"id": "price", "label": "가격이력", "dir": "history", "file": "price_changes.parquet",
     "kind": "single", "upload": "direct", "note": "상품수정삭제로그 · 1년 롤링 누적"},
    {"id": "stock", "label": "재고 스냅샷", "dir": "snapshots", "rx": r"stock_(\d{4}-\d{2})\.parquet",
     "kind": "monthly", "upload": "auto", "note": "상품관리 업로드 부산물 · 자동 적립"},
    {"id": "purchases", "label": "매입현황", "dir": "purchases", "rx": r"buyin_(\d{4}-\d{2})\.parquet",
     "kind": "monthly", "upload": "planned", "note": "유형별매입현황 · 실입고(예정)"},
    {"id": "demand", "label": "발주자료", "dir": "demand", "rx": r"order_(\d{4}-\d{2})\.parquet",
     "kind": "monthly", "upload": "planned", "note": "logistics 발주 · 리드타임 산출(예정)"},
]


def _api(url, pat):
    return urllib.request.urlopen(urllib.request.Request(
        url, headers={"Authorization": "Bearer " + pat, "User-Agent": "wa-app",
                      "Accept": "application/vnd.github+json"}))


def _list_dir(pat, repo, d):
    url = "https://api.github.com/repos/%s/contents/%s?ref=main" % (repo, d)
    try:
        return json.load(_api(url, pat))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return []
        raise


def _msnum(s):
    y, m = map(int, s.split("-"))
    return y * 12 + m


def _mslbl(v):
    return "%d-%02d" % ((v - 1) // 12, (v - 1) % 12 + 1)


def next_month(ym):
    """'2026-05' → '2026-06' (타임라인 막대 끝)."""
    return _mslbl(_msnum(ym) + 1)


def coverage(pat, repo):
    """카탈로그별 적재 현황 리스트. 각: id·label·kind·upload·note·status·files·size_kb·first·last·gaps."""
    out = []
    for c in CATALOG:
        items = _list_dir(pat, repo, c["dir"])
        rec = dict(c)
        if c["kind"] == "monthly":
            pairs = [(m.group(1), it.get("size", 0)) for it in items
                     for m in [re.match(c["rx"], it.get("name", ""))] if m]
            months = sorted(p[0] for p in pairs)
            rec["files"] = len(months)
            rec["size_kb"] = round(sum(p[1] for p in pairs) / 1024)
            if months:
                lo, hi = _msnum(months[0]), _msnum(months[-1])
                have = {_msnum(x) for x in months}
                rec["first"], rec["last"] = months[0], months[-1]
                rec["gaps"] = [_mslbl(v) for v in range(lo, hi + 1) if v not in have]
                rec["status"] = "ok"
            else:
                rec["first"] = rec["last"] = None
                rec["gaps"] = []
                rec["status"] = "empty"
        else:  # single
            f = next((it for it in items if it.get("name") == c.get("file")), None)
            rec["files"] = 1 if f else 0
            rec["size_kb"] = round(f.get("size", 0) / 1024) if f else 0
            rec["first"] = rec["last"] = None
            rec["gaps"] = []
            rec["status"] = "ok" if f else "empty"
        out.append(rec)
    return out
