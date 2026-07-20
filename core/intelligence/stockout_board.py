"""품절 알림판 — 발주 품절목록 → 재입고(박스재고 양수 전환) 시 입고로그+자동삭제 (ADR 0024).

데일리 대시보드 섹션. 비-PII(관리코드·상품명·날짜·박스재고)라 private data repo 영속:
- history/stockout_board.json : 현재 알림판 {관리코드: {상품명, since(품절시작일), 발주수량, seed현재고}}
- history/restock_log.csv     : 입고 로그 append(관리코드·상품명·품절시작일·입고일·품절일수·입고시박스재고)

흐름: 발주(Phase2) 품절목록 → seed(없는 코드 추가·시작일=그날·있으면 시작일 유지) →
      상품관리 갱신 후 데일리 대시보드 열 때 reconcile(박스재고>0 회복 → 입고로그+제거) · 수동 삭제(로그 없음).
★ 재입고 기준 = 박스재고 > 0 (현 품절건 다 음수 → 양수 전환, 사용자 확정 2026-06-16). 0=재고없음(유지).
★ 박스재고 = product_master '박스' 컬럼(박스내품 아님).
"""
from __future__ import annotations

import base64
import io
import json
import time
import unicodedata
import urllib.error
import urllib.request

import pandas as pd

BOARD_PATH = "history/stockout_board.json"
LOG_PATH = "history/restock_log.csv"
LOG_COLS = ["관리코드", "상품명", "품절시작일", "입고일", "품절일수", "입고시박스재고"]


def _nfc(v):
    return unicodedata.normalize("NFC", str(v)).strip() if v is not None and pd.notna(v) else ""


def _key(code, code_map: dict = None) -> str:
    """조회키 — 낱개/소분 코드는 **원코드(박스)**로 치환.

    ★ 매입현황(cadence)·상품관리(박스재고)는 전부 박스 관리코드 기준이라, 알림판에 낱개/소분
      코드로 등록된 건은 치환 없이는 최근입고/평균주기/박스재고가 전부 미매칭 → 재입고 자동삭제도
      영영 안 걸림. code_map = {낱개코드: 원코드} (logistics_order.unit_origin_map).
    """
    k = _nfc(code)
    if code_map:
        return _nfc(code_map.get(k, k))
    return k


def _num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError, AttributeError):
        return 0.0


def _gh(url, pat, method="GET", data=None, accept="application/vnd.github+json"):
    headers = {"Authorization": "Bearer " + pat, "User-Agent": "wa-app", "Accept": accept}
    if data is not None:
        headers["Content-Type"] = "application/json"
        data = json.dumps(data).encode("utf-8")
    return urllib.request.urlopen(urllib.request.Request(url, data=data, method=method, headers=headers))


def _url(repo, path):
    return "https://api.github.com/repos/%s/contents/%s" % (repo, path)


def _get_sha(pat, repo, path):
    try:
        return json.load(_gh(_url(repo, path) + "?ref=main", pat)).get("sha")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise


def _put_retry(pat, repo, path, content_bytes: bytes, msg: str, max_attempts: int = 4):
    """content_bytes 를 path 에 PUT. 매 시도 직전 fresh sha 재취득.

    GitHub 는 같은 파일을 짧은 간격으로 연속 PUT 하면 409(Conflict)/422 를 던지고,
    쓰기가 다발이면 403(secondary rate limit)을 던진다. 데일리 대시보드는 rerun 마다
    reconcile→쓰기가 돌 수 있고 raw read 지연으로 같은 건이 반복 처리될 수 있어(2026-07-20 사례),
    이들 코드는 sha 재취득 + backoff 로 재시도한다(429 도 포함).
    """
    content = base64.b64encode(content_bytes).decode("ascii")
    last = None
    for attempt in range(max_attempts):
        sha = _get_sha(pat, repo, path)
        body = {"message": msg, "content": content}
        if sha:
            body["sha"] = sha
        try:
            _gh(_url(repo, path), pat, method="PUT", data=body)
            return
        except urllib.error.HTTPError as e:
            last = e
            if e.code in (409, 422, 403, 429) and attempt < max_attempts - 1:
                time.sleep(1.5 * (attempt + 1))
                continue
            raise
    if last is not None:
        raise last


def read_board(pat, repo) -> dict:
    try:
        with _gh(_url(repo, BOARD_PATH) + "?ref=main", pat, accept="application/vnd.github.raw") as r:
            return json.loads(r.read().decode("utf-8")) or {}
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return {}
        raise


def write_board(pat, repo, board: dict, msg: str):
    content = json.dumps(board, ensure_ascii=False, indent=1).encode("utf-8")
    _put_retry(pat, repo, BOARD_PATH, content, msg)


def read_log(pat, repo) -> pd.DataFrame:
    try:
        with _gh(_url(repo, LOG_PATH) + "?ref=main", pat, accept="application/vnd.github.raw") as r:
            return pd.read_csv(io.BytesIO(r.read()), dtype=str)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return pd.DataFrame(columns=LOG_COLS)
        raise


def append_log(pat, repo, rows: list, msg: str) -> list:
    """입고 로그 append. **멱등** — 이미 (관리코드, 입고일)이 로그에 있으면 skip.

    ★ raw read 지연으로 같은 재입고 건이 rerun 사이 반복 판정되면 로그가 중복 기록되고,
      그 반복 PUT 이 GitHub 409/403 를 유발했다(2026-07-20 사례). 실제 커밋 직전 최신 로그를
      다시 읽어 (관리코드, 입고일) 중복분을 걸러낸다. return = 실제로 새로 추가된 rows.
    """
    if not rows:
        return []
    cur = read_log(pat, repo)
    existing = set()
    if not cur.empty and {"관리코드", "입고일"}.issubset(cur.columns):
        existing = {(_nfc(a), _nfc(b)) for a, b in zip(cur["관리코드"], cur["입고일"])}
    fresh = [r for r in rows
             if (_nfc(r.get("관리코드")), _nfc(r.get("입고일"))) not in existing]
    if not fresh:
        return []
    new = pd.concat([cur, pd.DataFrame(fresh, columns=LOG_COLS)], ignore_index=True)
    _put_retry(pat, repo, LOG_PATH, new.to_csv(index=False).encode("utf-8-sig"), msg)
    return fresh


def seed_from_stockout(board: dict, so_df, today: str):
    """품절목록(관리코드·상품명·발주수량·현재고) → 알림판에 없는 코드 추가(시작일=today·이미 있으면 유지).

    return (board, 추가된 관리코드 list).
    """
    added = []
    if so_df is None or len(so_df) == 0:
        return board, added
    for r in so_df.to_dict("records"):
        code = _nfc(r.get("관리코드", ""))
        if not code or code in board:
            continue
        board[code] = {"상품명": _nfc(r.get("상품명", "")), "since": today,
                       "발주수량": _num(r.get("발주수량", 0)), "seed현재고": _num(r.get("현재고", 0))}
        added.append(code)
    return board, added


def reconcile(board: dict, box_stock: dict, today: str, threshold: float = 0.0,
              code_map: dict = None):
    """알림판 각 항목 현재 박스재고 > threshold(0) 회복 시 입고로그 행 생성·알림판에서 제거.

    box_stock = {관리코드(NFC): 박스재고}. return (남은 board, 입고로그 rows).
    """
    restocked = []
    keep = {}
    td = pd.Timestamp(today)
    for code, info in board.items():
        bs = box_stock.get(_key(code, code_map))
        if bs is not None and bs > threshold:
            since = info.get("since", "")
            try:
                days = (td - pd.Timestamp(since)).days
            except Exception:
                days = ""
            restocked.append({"관리코드": code, "상품명": info.get("상품명", ""),
                              "품절시작일": since, "입고일": today,
                              "품절일수": days, "입고시박스재고": bs})
        else:
            keep[code] = info
    return keep, restocked


def manual_remove(board: dict, code: str) -> dict:
    """수동 제거(로그 없음)."""
    return {k: v for k, v in board.items() if _nfc(k) != _nfc(code)}


def board_to_frame(board: dict, box_stock: dict, today: str, cadence: dict = None,
                   code_map: dict = None) -> pd.DataFrame:
    cadence = cadence or {}
    td = pd.Timestamp(today)
    rows = []
    for code, info in board.items():
        since = info.get("since", "")
        try:
            days = (td - pd.Timestamp(since)).days
        except Exception:
            days = None
        _k = _key(code, code_map)
        ci = cadence.get(_k, {})
        last = ci.get("최근입고일")
        avg = ci.get("평균주기")
        cnt = ci.get("입고횟수")
        rows.append({"관리코드": code, "상품명": info.get("상품명", ""),
                     "품절시작일": since, "N일째": days,
                     "현재박스재고": box_stock.get(_k),
                     "발주수량": info.get("발주수량"),
                     "최근입고일": (last.strftime("%Y-%m-%d") if (last is not None and pd.notna(last)) else ""),
                     "평균매입주기": (round(avg) if avg is not None else None),
                     "입고횟수(1년)": (int(cnt) if cnt else None)})
    df = pd.DataFrame(rows, columns=["관리코드", "상품명", "품절시작일", "N일째", "현재박스재고",
                                     "발주수량", "최근입고일", "평균매입주기", "입고횟수(1년)"])
    if not df.empty:
        df = df.sort_values("품절시작일").reset_index(drop=True)
    return df
