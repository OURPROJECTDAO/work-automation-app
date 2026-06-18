"""시장 지능(nadl) — 경쟁 행사가 로드 + 상품 매칭 모델 + 매핑 영속.

시장지능 트랙(ADR 0025, intelligence-layer.md Phase 3). 로컬 수집기(구조 B)가
work-automation-data:market/nadl/prices_{date}.parquet 에 행사 개당가를 적립.
이 모듈은 그 parquet 로드 + 우리 product_master 매칭 후보 생성 + 확정 매핑(ps_goid↔관리코드) 저장.

매칭 모델 v2 (대화 검증 21건 top3 recall 16/16·1위 15/16):
  1. 박스규격 하드게이트 = 용량(L→ml·kg→g 정규화) 일치 ∧ 팩수 일치.
     (nadl=마트 납품·소분 없음 → 박스 규격 다르면 다른 상품, 사용자 규칙)
  2. 랭킹 = 글자 bigram 자카드(띄어쓰기·어순 무시) + 브랜드 가중(동일+0.15 / 다른 명시 브랜드−0.15).
  3. top3 제시 · 자동 "없음" 판정 안 함(점수는 신뢰 힌트). 최종 확정은 사람.

매핑 = market/nadl/nadl_map.csv (ps_goid 키). 중복등록(스팸 마일드 등)은 ps_goid당 복수행 허용.
  status: matched(관리코드 보유) / none(검토했으나 매칭 없음).
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

PRICE_DIR = "market/nadl"
_PRICE_RE = re.compile(r"^prices_(\d{4}-\d{2}-\d{2})\.parquet$")
MAP_PATH = "market/nadl/nadl_map.csv"
MAP_COLS = ["ps_goid", "nadl_name", "nadl_spec", "관리코드", "status", "updated"]

# 데이터에서 관찰된 주요 FMCG 브랜드(가장 긴 매칭 우선). 실사용하며 보강.
BRANDS = [
    "롯데칠성", "롯데", "동원", "한성", "사조", "백설", "cj", "오뚜기", "청정원", "해태",
    "농심", "삼양", "남양", "매일", "빙그레", "풀무원", "대상", "애경", "유한", "피죤",
    "다우니", "환타", "코카콜라", "스프라이트", "펩시", "탑씨", "웅진", "오란씨", "써니텐",
    "코코팜", "티로그", "이프로", "스팸", "샘표", "하선정", "광동", "킨사이다", "미원",
    "맥심", "담터", "동서", "립톤", "게토레이", "파워에이드", "포카리", "데미소다", "솔의눈",
]


def _nfc(v) -> str:
    return unicodedata.normalize("NFC", str(v)).strip() if v is not None else ""


def _namejoin(name) -> str:
    return re.sub(r"[^0-9a-z가-힣]+", "", _nfc(name).lower())


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


# ── 박스규격 파싱 ─────────────────────────────────────────────────────────
def parse_size(s):
    """문자열 → (용량값, 단위) 정규화(L→ml·kg→g). 못 읽으면 None."""
    t = _nfc(s).lower().replace(",", "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*(ml|kg|g|l|t)\b", t) or \
        re.search(r"(\d+(?:\.\d+)?)\s*(ml|kg|g|l|t)", t)
    if not m:
        return None
    v = float(m.group(1)); u = m.group(2)
    if u == "l":
        v *= 1000; u = "ml"
    elif u == "kg":
        v *= 1000; u = "g"
    return (round(v, 1), u)


def nadl_pack(spec):
    """nadl 규격 '용량단위 * N개' → 팩수 N."""
    s = _nfc(spec)
    m = re.search(r"\*\s*(\d+)\s*개", s) or re.search(r"\*\s*(\d+)", s)
    return int(m.group(1)) if m else None


def nadl_box(name, spec):
    """nadl 행 → (용량, 단위, 팩수). 규격 우선, 용량은 이름 폴백."""
    sz = parse_size(spec) or parse_size(name)
    pk = nadl_pack(spec)
    if sz is None or pk is None:
        return None
    return (sz[0], sz[1], pk)


def pm_box(상품명, 규격, 박스내품):
    """우리 행 → (용량, 단위, 팩수). 용량=규격→상품명 폴백, 팩수=박스내품>0 우선→규격 *N/+N."""
    sz = parse_size(규격) or parse_size(상품명)
    pk = None
    try:
        bx = int(float(박스내품))
        if bx > 0:
            pk = bx
    except (TypeError, ValueError):
        pass
    if pk is None:
        m = re.search(r"[\*\+]\s*(\d+)", _nfc(규격))
        pk = int(m.group(1)) if m else None
    if sz is None or pk is None:
        return None
    return (sz[0], sz[1], pk)


def find_brand(name):
    """이름에서 가장 긴 브랜드 토큰(없으면 None)."""
    j = _namejoin(name)
    hits = [b for b in BRANDS if b and b in j]
    return max(hits, key=len) if hits else None


def _bigrams(s):
    j = _namejoin(s)
    return set(j[i:i + 2] for i in range(len(j) - 1)) if len(j) >= 2 else ({j} if j else set())


def _bigram_jac(a, b):
    A, B = _bigrams(a), _bigrams(b)
    return len(A & B) / len(A | B) if A and B else 0.0


def _lead_token(name):
    t = _nfc(name).split()
    return t[0] if t else ""


def score(nadl_name, pm_name, pm_brand, nadl_brand):
    """이름 점수 = bigram 자카드 + 브랜드 가중."""
    sc = _bigram_jac(nadl_name, pm_name)
    nb = nadl_brand or _namejoin(_lead_token(nadl_name))
    if nb:
        ourj = _namejoin(pm_name)
        if nb in ourj:
            sc += 0.15
        elif pm_brand and pm_brand != nadl_brand:
            sc -= 0.15
    return sc


# ── product_master 인덱싱 + 후보 생성 ─────────────────────────────────────
def build_pm_index(pm: pd.DataFrame) -> dict:
    """product_master → 박스키 (용량,단위,팩수) → [{관리코드·상품명·규격·박스내품·매입단가·매출단가·brand}]."""
    idx: dict = {}
    for r in pm.itertuples(index=False):
        d = r._asdict()
        box = pm_box(d.get("상품명"), d.get("규격"), d.get("박스내품"))
        if box is None:
            continue
        rec = {
            "관리코드": _nfc(d.get("관리코드")),
            "상품명": _nfc(d.get("상품명")),
            "규격": _nfc(d.get("규격")),
            "박스내품": d.get("박스내품"),
            "매입단가": _num(d.get("매입단가")),
            "매출단가": _num(d.get("매출단가")),
            "brand": find_brand(d.get("상품명")),
        }
        idx.setdefault(box, []).append(rec)
    return idx


def suggest(nadl_name, nadl_spec, pm_index: dict, topn: int = 3) -> list:
    """nadl 한 건 → 박스 일치 후보 중 점수 top3. 각: 점수·관리코드·상품명·규격·매입단가·매출단가."""
    box = nadl_box(nadl_name, nadl_spec)
    if box is None:
        return []
    cands = pm_index.get(box, [])
    if not cands:
        return []
    nb = find_brand(nadl_name)
    scored = []
    for c in cands:
        sc = score(nadl_name, c["상품명"], c["brand"], nb)
        scored.append((round(sc, 3), c))
    scored.sort(key=lambda x: x[0], reverse=True)
    out = []
    for sc, c in scored[:topn]:
        out.append({"점수": sc, "관리코드": c["관리코드"], "상품명": c["상품명"],
                    "규격": c["규격"], "매입단가": c["매입단가"], "매출단가": c["매출단가"]})
    return out


# ── GitHub I/O (work-automation-data) ─────────────────────────────────────
def _gh(url, pat, method="GET", data=None, accept="application/vnd.github+json"):
    headers = {"Authorization": "Bearer " + pat, "User-Agent": "wa-app", "Accept": accept}
    if data is not None:
        headers["Content-Type"] = "application/json"
        data = json.dumps(data).encode("utf-8")
    return urllib.request.urlopen(urllib.request.Request(url, data=data, method=method, headers=headers))


def list_price_dates(pat, repo) -> list:
    """market/nadl/ 의 prices_YYYY-MM-DD 날짜 목록(정렬)."""
    url = "https://api.github.com/repos/%s/contents/%s?ref=main" % (repo, PRICE_DIR)
    try:
        items = json.load(_gh(url, pat))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return []
        raise
    dates = [m.group(1) for it in items for m in [_PRICE_RE.match(it.get("name", ""))] if m]
    return sorted(dates)


def load_prices(pat, repo, date=None) -> pd.DataFrame:
    """prices_{date}.parquet 로드(date 미지정=최신). 없으면 빈 DataFrame."""
    if date is None:
        ds = list_price_dates(pat, repo)
        if not ds:
            return pd.DataFrame()
        date = ds[-1]
    path = "%s/prices_%s.parquet" % (PRICE_DIR, date)
    url = "https://api.github.com/repos/%s/contents/%s?ref=main" % (repo, path)
    try:
        with _gh(url, pat, accept="application/vnd.github.raw") as r:
            df = pd.read_parquet(io.BytesIO(r.read()))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return pd.DataFrame()
        raise
    df["_date"] = date
    return df


def read_map(pat, repo) -> pd.DataFrame:
    """nadl_map.csv → DataFrame. 없으면 빈(스키마 보존)."""
    url = "https://api.github.com/repos/%s/contents/%s?ref=main" % (repo, MAP_PATH)
    try:
        with _gh(url, pat, accept="application/vnd.github.raw") as r:
            df = pd.read_csv(io.BytesIO(r.read()), dtype=str, encoding="utf-8-sig")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return pd.DataFrame(columns=MAP_COLS)
        raise
    for c in MAP_COLS:
        if c not in df.columns:
            df[c] = ""
    return df[MAP_COLS].fillna("")


def write_map(df: pd.DataFrame, pat, repo) -> dict:
    """nadl_map.csv 전체 덮어쓰기(sha GET 후 PUT). utf-8-sig."""
    df = df.reindex(columns=MAP_COLS).fillna("")
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    content = base64.b64encode(buf.getvalue().encode("utf-8-sig")).decode("ascii")
    url = "https://api.github.com/repos/%s/contents/%s" % (repo, MAP_PATH)
    sha = None
    try:
        sha = json.load(_gh(url + "?ref=main", pat)).get("sha")
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
    body = {"message": "data: nadl 매핑 갱신 (%d행)" % len(df), "content": content}
    if sha:
        body["sha"] = sha
    _gh(url, pat, method="PUT", data=body)
    return {"rows": len(df)}


# ── 매칭본(매칭 결과 결합) ────────────────────────────────────────────────
def build_matched(prices: pd.DataFrame, map_df: pd.DataFrame, pm: pd.DataFrame) -> pd.DataFrame:
    """전 nadl 행 + 매칭된 건은 우리 컬럼 결합, 미매칭은 공란.

    중복등록(ps_goid 복수 관리코드)은 복수행. status=none(검토완료)은 미매칭과 구분 표시.
    """
    pm = pm.copy()
    pm["_mg"] = pm["관리코드"].map(_nfc)
    pm_lookup = {r["_mg"]: r for _, r in pm.iterrows()}

    matched = map_df[map_df["status"] == "matched"].copy()
    reviewed_none = set(map_df.loc[map_df["status"] == "none", "ps_goid"].map(_nfc))
    by_goid: dict = {}
    for _, m in matched.iterrows():
        by_goid.setdefault(_nfc(m["ps_goid"]), []).append(_nfc(m["관리코드"]))

    rows = []
    for r in prices.itertuples(index=False):
        d = r._asdict()
        goid = _nfc(d.get("ps_goid"))
        base = {
            "nadl_상품명": d.get("name"), "nadl_규격": d.get("spec"),
            "박스가": d.get("box_price"), "개당가": d.get("unit_price"),
            "ps_goid": goid,
        }
        codes = by_goid.get(goid, [])
        if codes:
            for code in codes:
                pr = pm_lookup.get(code)
                row = dict(base)
                row["상태"] = "매칭"
                row["관리코드"] = code
                if pr is not None:
                    row["우리_상품명"] = pr["상품명"]
                    row["우리_규격"] = pr["규격"]
                    row["박스내품"] = pr["박스내품"]
                    row["우리_매입단가"] = _num(pr["매입단가"])
                    row["우리_매출단가"] = _num(pr["매출단가"])
                rows.append(row)
        else:
            row = dict(base)
            row["상태"] = "검토완료(없음)" if goid in reviewed_none else "미매칭"
            row["관리코드"] = ""
            rows.append(row)
    cols = ["상태", "nadl_상품명", "nadl_규격", "박스가", "개당가", "관리코드",
            "우리_상품명", "우리_규격", "박스내품", "우리_매입단가", "우리_매출단가", "ps_goid"]
    out = pd.DataFrame(rows)
    for c in cols:
        if c not in out.columns:
            out[c] = None
    return out[cols]
