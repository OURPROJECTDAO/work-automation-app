"""두뇌④ 결정 원장 (ADR 0026 · Gate 3, forward-only).

기준마진율 변경 *결정* + 이후 *반응*을 work-automation-data:history/decisions.parquet 에 누적.
백필 불가 — 지금부터 forward. "시도→측정→유지/되돌림" 루프의 기억(다음 사이클이 측정 채움).
비-PII(관리코드·채널·마진·집계수치만).
"""
from __future__ import annotations

import base64
import io
import json
import urllib.error
import urllib.parse
import urllib.request
import uuid
from datetime import date

import pandas as pd

_PATH = "history/decisions.parquet"
COLS = [
    "decision_id", "ts", "관리코드", "채널", "액션",
    "마진_before", "마진_권장", "마진_적용", "베이스", "플래그", "사유",
    "측정전_월볼륨", "측정전_월순이익",
    "status", "측정일", "측정후_월볼륨", "측정후_월순이익", "결과",
]


def _api(repo: str, path: str) -> str:
    seg = "/".join(urllib.parse.quote(s) for s in path.split("/"))
    return "https://api.github.com/repos/%s/contents/%s" % (repo, seg)


def _get(repo: str, path: str, pat: str):
    req = urllib.request.Request(_api(repo, path), headers={
        "Authorization": "Bearer " + pat, "Accept": "application/vnd.github+json",
        "User-Agent": "wb"})
    try:
        with urllib.request.urlopen(req) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise


def read_all(pat: str, repo: str) -> pd.DataFrame:
    j = _get(repo, _PATH, pat)
    if not j or "content" not in j:
        return pd.DataFrame(columns=COLS)
    raw = base64.b64decode(j["content"])
    try:
        return pd.read_parquet(io.BytesIO(raw))
    except Exception:
        return pd.DataFrame(columns=COLS)


def _put(pat: str, repo: str, df: pd.DataFrame, msg: str):
    buf = io.BytesIO()
    df.to_parquet(buf, index=False)
    body = {"message": msg, "content": base64.b64encode(buf.getvalue()).decode()}
    j = _get(repo, _PATH, pat)
    if j and "sha" in j:
        body["sha"] = j["sha"]
    req = urllib.request.Request(_api(repo, _PATH), method="PUT",
        data=json.dumps(body).encode(), headers={
            "Authorization": "Bearer " + pat, "Accept": "application/vnd.github+json",
            "User-Agent": "wb"})
    with urllib.request.urlopen(req) as r:
        return json.load(r)


def append(records: list[dict], pat: str, repo: str) -> int:
    """결정 기록 추가. records 각 dict 키 = 관리코드·채널·액션·마진_before·마진_권장·
    (옵션 마진_적용)·베이스·플래그·사유·측정전_월볼륨·측정전_월순이익. status=pending 으로 들어감."""
    if not records:
        return 0
    cur = read_all(pat, repo)
    today = date.today().isoformat()
    rows = []
    for r in records:
        row = {c: r.get(c) for c in COLS}
        row["decision_id"] = uuid.uuid4().hex[:12]
        row["ts"] = today
        if row.get("마진_적용") is None:
            row["마진_적용"] = r.get("마진_권장")
        row["status"] = "pending"
        rows.append(row)
    new = pd.concat([cur, pd.DataFrame(rows, columns=COLS)], ignore_index=True)
    _put(pat, repo, new, "decisions: +%d 두뇌④ 결정 기록" % len(rows))
    return len(rows)


def init_empty(pat: str, repo: str):
    """원장이 없으면 빈 스키마로 생성(seed). 있으면 no-op."""
    j = _get(repo, _PATH, pat)
    if j:
        return False
    _put(pat, repo, pd.DataFrame(columns=COLS), "decisions: 빈 원장 생성(두뇌④ Gate3)")
    return True


def update(records: list[dict], pat: str, repo: str) -> int:
    """decision_id 로 기존 행 in-place 갱신(측정/조치). records 각 dict = decision_id + 덮어쓸 COLS 키.
    예 측정: {"decision_id":..,"측정일":..,"측정후_월볼륨":..,"측정후_월순이익":..,"결과":..,"status":"measured"}.
    예 조치: {"decision_id":..,"status":"closed"|"reverted"}. 없는 id/키는 무시.
    """
    if not records:
        return 0
    cur = read_all(pat, repo)
    if cur.empty:
        return 0
    cur = cur.set_index("decision_id")
    n = 0
    for r in records:
        did = r.get("decision_id")
        if did is None or did not in cur.index:
            continue
        for k, v in r.items():
            if k == "decision_id" or k not in cur.columns:
                continue
            cur.loc[did, k] = v
        n += 1
    cur = cur.reset_index()
    if n:
        _put(pat, repo, cur, "decisions: 측정/조치 갱신 %d건" % n)
    return n
