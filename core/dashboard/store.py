"""대시보드 master 저장 어댑터 — private repo work-automation-data 의 월별 parquet 파티션 R/W.

프레임워크 무관(core): token·repo를 인자로 받음. Streamlit 페이지가 st.secrets로 주입.
GitHub contents API: 읽기=raw 헤더(바이너리 직행), 쓰기=base64 PUT(기존 sha GET 후).
한글 경로 없음(master/sales_YYYY-MM.parquet) → ASCII. 대용량은 urllib(curl arg 한계 회피).

설계 근거: decisions/0006 (#3 월파티션, #4 날짜구간 교체, #7→B2 저장소).
"""
from __future__ import annotations

import base64
import io
import json
import re
import urllib.error
import urllib.parse
import urllib.request

import pandas as pd

from .sales_data import parse_sales, split_by_month, date_range_replace

_PART_DIR = "master"
_PART_RE = re.compile(r"^sales_(\d{4}-\d{2})\.parquet$")


def _req(token: str, url: str, *, data=None, method="GET", accept="application/vnd.github+json"):
    headers = {"Authorization": "Bearer " + token, "Accept": accept,
               "User-Agent": "work-automation-app"}
    if data is not None:
        headers["Content-Type"] = "application/json"
    return urllib.request.Request(url, data=data, method=method, headers=headers)


def _contents_url(repo: str, path: str) -> str:
    return f"https://api.github.com/repos/{repo}/contents/{urllib.parse.quote(path)}"


def _get_sha(token: str, repo: str, path: str):
    try:
        with urllib.request.urlopen(_req(token, _contents_url(repo, path))) as r:
            return json.load(r).get("sha")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise


def list_partition_months(token: str, repo: str) -> list[str]:
    """master/ 디렉토리의 sales_YYYY-MM.parquet 들에서 YYYY-MM 목록(정렬)."""
    try:
        with urllib.request.urlopen(_req(token, _contents_url(repo, _PART_DIR))) as r:
            items = json.load(r)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return []
        raise
    months = []
    for it in items:
        m = _PART_RE.match(it.get("name", ""))
        if m:
            months.append(m.group(1))
    return sorted(months)


def read_partition(token: str, repo: str, ym: str) -> pd.DataFrame | None:
    """월 파티션 parquet → DataFrame. 없으면 None."""
    path = f"{_PART_DIR}/sales_{ym}.parquet"
    try:
        # raw 헤더 → 바이너리 직행(base64 디코딩 불필요)
        with urllib.request.urlopen(
                _req(token, _contents_url(repo, path), accept="application/vnd.github.raw")) as r:
            return pd.read_parquet(io.BytesIO(r.read()))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise


def write_partition(token: str, repo: str, ym: str, df: pd.DataFrame, msg: str | None = None):
    """DataFrame → 월 파티션 parquet 업로드(기존 sha GET 후 덮어씀)."""
    path = f"{_PART_DIR}/sales_{ym}.parquet"
    buf = io.BytesIO()
    df.to_parquet(buf, index=False)
    body = {"message": msg or f"dashboard: {ym} 파티션 갱신 ({len(df)}행)",
            "content": base64.b64encode(buf.getvalue()).decode(), "branch": "main"}
    sha = _get_sha(token, repo, path)
    if sha:
        body["sha"] = sha
    req = _req(token, _contents_url(repo, path), data=json.dumps(body).encode(), method="PUT")
    with urllib.request.urlopen(req) as r:
        json.load(r)


def delete_partition(token: str, repo: str, ym: str) -> bool:
    """월 파티션 삭제. 없으면 False. (잘못 적재된 달 제거용.)"""
    path = f"{_PART_DIR}/sales_{ym}.parquet"
    sha = _get_sha(token, repo, path)
    if not sha:
        return False
    body = {"message": f"dashboard: {ym} 파티션 삭제", "sha": sha, "branch": "main"}
    req = _req(token, _contents_url(repo, path), data=json.dumps(body).encode(), method="DELETE")
    with urllib.request.urlopen(req):
        return True


def load_master(token: str, repo: str) -> pd.DataFrame:
    """전 파티션을 합쳐 master DataFrame. @st.cache_data로 1회 로드 권장."""
    months = list_partition_months(token, repo)
    parts = [read_partition(token, repo, ym) for ym in months]
    parts = [p for p in parts if p is not None and len(p)]
    if not parts:
        return pd.DataFrame()
    return pd.concat(parts, ignore_index=True).sort_values("거래일자").reset_index(drop=True)


def ingest(token: str, repo: str, file_or_buf) -> dict:
    """영업이익현황 업로드 → 영향받은 달만 날짜구간 교체로 갱신.

    파일이 여러 달에 걸치면 각 달 파티션에 대해, 그 달 슬라이스의 [min,max] 구간을
    기존 파티션에서 지우고 새 데이터 삽입(date_range_replace). 건드린 달만 재기록.
    반환: 처리 요약(rows, months, date_range).
    """
    new = parse_sales(file_or_buf)
    if new.empty:
        return {"rows": 0, "months": [], "date_range": None}
    touched = []
    for ym, new_month in split_by_month(new).items():
        existing = read_partition(token, repo, ym)
        merged = date_range_replace(existing, new_month)
        write_partition(token, repo, ym, merged,
                        msg=f"dashboard ingest: {ym} 날짜구간 교체 (+{len(new_month)}행)")
        touched.append(ym)
    return {
        "rows": len(new),
        "months": sorted(touched),
        "date_range": (str(new["거래일자"].min().date()), str(new["거래일자"].max().date())),
    }


# ── 거래처 그룹 매핑 (상호명→그룹) — private repo groups/store_groups.csv ──
_GROUPS_PATH = "groups/store_groups.csv"


def read_groups(token: str, repo: str) -> pd.DataFrame:
    """상호명→그룹 매핑 DataFrame(컬럼 상호명,그룹). 없으면 빈 DataFrame."""
    try:
        with urllib.request.urlopen(
                _req(token, _contents_url(repo, _GROUPS_PATH),
                     accept="application/vnd.github.raw")) as r:
            return pd.read_csv(io.BytesIO(r.read()), dtype=str, encoding="utf-8-sig")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return pd.DataFrame(columns=["상호명", "그룹"])
        raise


def write_groups(token: str, repo: str, df: pd.DataFrame, msg: str | None = None):
    """상호명,그룹 DataFrame → groups/store_groups.csv 커밋(기존 sha GET 후 덮어씀)."""
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    body = {"message": msg or f"dashboard: 거래처 그룹 갱신 ({len(df)}행)",
            "content": base64.b64encode(csv_bytes).decode(), "branch": "main"}
    sha = _get_sha(token, repo, _GROUPS_PATH)
    if sha:
        body["sha"] = sha
    req = _req(token, _contents_url(repo, _GROUPS_PATH),
               data=json.dumps(body).encode(), method="PUT")
    with urllib.request.urlopen(req) as r:
        json.load(r)
