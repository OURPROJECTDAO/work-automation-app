"""EasyAdmin 확장주문검색(.xls·HTML 위장) → 주문 이력(velocity·송장 그룹) 적재.

이력 엔진 (ADR 0018, intelligence-layer.md §3.5). 채널별 판매 velocity + 송장번호 그룹(택배 실배분, P2).
- 저장: private repo work-automation-data : orders/easyadmin_YYYY-MM.parquet (월 파티션).
- ★ HTML 위장 .xls → pd.read_html(header=0). 첫 행이 텍스트 헤더.
- ★ PII 전제거(수령자/주문자 이름·전화·주소·주문번호·관리번호·CS). 송장번호 → 해시(그룹키만).
- 기준일 = 발주일 우선(없으면 주문일) — export 필터=발주일이라 파일↔월 매핑 정확(주문일 우선이면 경계 scatter).
- 정산금액은 raw(매출자료가 정산 진실, §2A) — 보관만, 마진엔 미사용.
- 멱등 = 기준일 날짜구간 교체(고유 거래ID 없음 — sales와 동일).

컬럼(위치 인덱스, 실물 대조 2026-06-15 Mar-May 25,822행 / KB §3.5):
  보존 0 erp관리코드·3 상품명·4 옵션명·6 판매처·8 상태·9 상품수량·11 판매처상품코드·13 택배사·
        23 상품코드·28 카테고리·29 주문수량·30 판매가·31 정산금액raw·32 수수료·33 주문일·34 발주일·
        39 판매처그룹·40 옵션추가항목1 · 10 송장번호→송장그룹(해시)
  PII(제외) 1·2·5·7·12·14·15·16·17·18·19·20·21·22·24·25·26·27·35·36·37·38·41
"""
from __future__ import annotations

import base64
import hashlib
import io
import json
import re
import unicodedata
import urllib.error
import urllib.request

import pandas as pd

PART_DIR = "orders"
_PART_RE = re.compile(r"^easyadmin_(\d{4}-\d{2})\.parquet$")

_KEEP = {0: "erp관리코드", 3: "상품명", 4: "옵션명", 6: "판매처", 8: "상태", 9: "상품수량",
         11: "판매처상품코드", 13: "택배사", 23: "상품코드", 28: "카테고리", 29: "주문수량",
         30: "판매가", 31: "정산금액raw", 32: "수수료", 33: "주문일", 34: "발주일",
         39: "판매처그룹", 40: "옵션추가항목1"}
_INVOICE_COL = 10
# 개별 .xls(HTML) PII 위치(키 생성용·저장 안 함). 통파일(.xlsx)은 Source.Name +1 시프트 → 헤더명 매핑 별도.
_PII = {"이름": 12, "휴대폰": 14, "전화": 17, "주소": 18}
_KEEP_CH = 6      # 판매처(채널 prefix)
_KEEP_BALJU = 34  # 발주일(합포박스키 일자)
_STR_COLS = ("erp관리코드", "상품명", "옵션명", "판매처", "상태", "판매처상품코드", "택배사",
             "상품코드", "카테고리", "판매처그룹", "옵션추가항목1")
_NUM_COLS = ("상품수량", "주문수량", "판매가", "정산금액raw", "수수료")
COLS = (["기준일"] + list(_KEEP.values()) + ["송장그룹", "고객키", "합포박스키"])


def _nfc(v):
    return unicodedata.normalize("NFC", str(v)).strip() if pd.notna(v) else ""


def _num(v):
    if pd.isna(v):
        return None
    s = str(v).replace(",", "").strip()
    if s in ("", "-"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _hash_invoice(v):
    s = _nfc(v)
    return hashlib.sha1(s.encode()).hexdigest()[:12] if s and s.lower() != "nan" else ""


# ===== 채널 내 고객키 / 합포박스키 (ADR 0021) =====
# robust 키 = 보이는 최소공통분모(마스킹/풀 무관 시간 일관). 솔트=st.secrets[data].customer_key_salt.
# - 이름: 끝1자 제외(마스킹=끝1자, 실측 99.99%)·괄호 주석 제거·공백 제거.
# - 전화: 뒤4자리(마스킹돼도 끝4 보임). 안심번호(4-4-4 가상·G마켓/옥션/쿠팡 등) 제외.
# - 주소(고객키): 숫자/* 토큰 전까지 앞 토큰(시군구+도로명).
# - 합포박스키: 풀이름+풀주소+발주일(박스=1회성·다운로드 내 일관).  솔트 없으면 키 빈값.
_PAREN = re.compile(r"[\(\[（【].*?[\)\]）】]")
ANON_CH = {"쿠팡(자동)", "11번가", "멸치쇼핑", "셀러허브 #1#1", "옥션 #2",
           "G마켓", "G마켓 #2", "위메이크프라이스(자동)"}


def _norm_ws(s):
    return re.sub(r"\s+", " ", s).strip()


def _name_front(nm):
    s = re.sub(r"\s+", "", _PAREN.sub("", _nfc(nm)))
    if len(s) >= 2:
        s = s[:-1]
    return s.rstrip("*")


def _phone_last4(ph, anon):
    if anon:
        return ""
    s = _nfc(ph)
    parts = s.split("-")
    if len(parts) == 3 and parts[0].isdigit() and len(parts[0]) == 4 and "*" not in s:
        return ""                                  # 안심번호(4-4-4) per-row
    d = re.findall(r"\d", s)
    return "".join(d[-4:]) if len(d) >= 4 else ""


def _addr_coarse(ad):
    out = []
    for t in _norm_ws(_nfc(ad)).split():
        if any(c.isdigit() for c in t) or "*" in t:
            break
        out.append(t)
        if len(out) >= 3:
            break
    return " ".join(out)


def _date10(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v)[:10]
    return s if len(s) >= 8 and s[:4].isdigit() else ""


def _cust_key(salt, ch, nm, ph, ad):
    nf = _name_front(nm)
    if not (salt and nf):
        return ""
    frag = "|".join([ch, nf, _phone_last4(ph, ch in ANON_CH), _addr_coarse(ad)])
    return hashlib.sha1((salt + "|C|" + frag).encode()).hexdigest()[:16]


def _box_key(salt, ch, nm, ad, balju):
    fn, fa, bd = _norm_ws(_nfc(nm)), _norm_ws(_nfc(ad)), _date10(balju)
    if not (salt and fn and fa and bd):
        return ""
    frag = "|".join([ch, fn, fa, bd])
    return hashlib.sha1((salt + "|B|" + frag).encode()).hexdigest()[:16]


def parse_orders(file, salt: str | None = None) -> pd.DataFrame:
    """확장주문검색 .xls(경로/bytes/파일객체) → 비-PII 주문 DataFrame(기준일·송장그룹·고객키·합포박스키).

    salt(=st.secrets[data].customer_key_salt) 있으면 robust 고객키/합포박스키 생성, 없으면 빈값.
    PII(이름·전화·주소)는 키 생성에만 쓰고 저장 안 함(원본 폐기).
    """
    src = io.BytesIO(file) if isinstance(file, (bytes, bytearray)) else file
    raw = pd.read_html(src, header=0)[0]
    if raw.shape[1] <= max(_KEEP):
        raise ValueError(f"확장주문검색 컬럼 부족: {raw.shape[1]}열 (≥{max(_KEEP) + 1} 필요)")
    out = pd.DataFrame({name: raw.iloc[:, i] for i, name in _KEEP.items()})
    out["송장그룹"] = raw.iloc[:, _INVOICE_COL].map(_hash_invoice)        # 원본 미저장(해시 그룹키)
    # 키 생성(PII positional → 해시 → raw 폐기)
    ch = raw.iloc[:, _KEEP_CH].map(_nfc)
    nm = raw.iloc[:, _PII["이름"]]
    ph = raw.iloc[:, _PII["휴대폰"]].where(raw.iloc[:, _PII["휴대폰"]].notna(), raw.iloc[:, _PII["전화"]])
    ad = raw.iloc[:, _PII["주소"]]
    bj = raw.iloc[:, _KEEP_BALJU]
    out["고객키"] = [_cust_key(salt, c, n, p, a) for c, n, p, a in zip(ch, nm, ph, ad)]
    out["합포박스키"] = [_box_key(salt, c, n, a, b) for c, n, a, b in zip(ch, nm, ad, bj)]
    for c in _STR_COLS:
        out[c] = out[c].map(_nfc).astype("string")
    for c in _NUM_COLS:
        out[c] = out[c].map(_num)
    out["주문일"] = pd.to_datetime(out["주문일"], errors="coerce")
    out["발주일"] = pd.to_datetime(out["발주일"], errors="coerce")
    out["기준일"] = out["발주일"].fillna(out["주문일"])                   # 발주일 우선(없으면 주문일)
    out = out[out["기준일"].notna()].copy()                              # 양쪽 결측 행 제외
    return out[COLS].sort_values(["기준일", "판매처"]).reset_index(drop=True)


def split_by_month(df: pd.DataFrame) -> dict:
    """기준일 월(YYYY-MM)별 슬라이스."""
    return {ym: g.copy() for ym, g in df.groupby(df["기준일"].dt.strftime("%Y-%m"))}


def date_range_replace(existing, new: pd.DataFrame) -> pd.DataFrame:
    """기존 파티션에서 new 기준일 [min,max] 구간을 지우고 new 삽입(고유ID 없음 → 구간 교체)."""
    if existing is None or len(existing) == 0:
        return new.reset_index(drop=True)
    lo, hi = new["기준일"].min(), new["기준일"].max()
    keep = existing[(existing["기준일"] < lo) | (existing["기준일"] > hi)]
    return (pd.concat([keep, new], ignore_index=True)
              .sort_values(["기준일", "판매처"]).reset_index(drop=True))


# ---- GitHub parquet R/W (work-automation-data) ----
def _gh(url, pat, method="GET", data=None, accept="application/vnd.github+json"):
    headers = {"Authorization": "Bearer " + pat, "User-Agent": "wa-app", "Accept": accept}
    if data is not None:
        headers["Content-Type"] = "application/json"
        data = json.dumps(data).encode("utf-8")
    return urllib.request.urlopen(urllib.request.Request(url, data=data, method=method, headers=headers))


def _part_path(ym):
    return "%s/easyadmin_%s.parquet" % (PART_DIR, ym)


def read_partition(pat, repo, ym) -> pd.DataFrame:
    url = "https://api.github.com/repos/%s/contents/%s" % (repo, _part_path(ym))
    try:
        with _gh(url + "?ref=main", pat, accept="application/vnd.github.raw") as r:
            return pd.read_parquet(io.BytesIO(r.read()))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return pd.DataFrame(columns=COLS)
        raise


def list_months(pat, repo) -> list:
    url = "https://api.github.com/repos/%s/contents/%s" % (repo, PART_DIR)
    try:
        items = json.load(_gh(url + "?ref=main", pat))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return []
        raise
    return sorted(m.group(1) for it in items for m in [_PART_RE.match(it.get("name", ""))] if m)


def read_all(pat, repo) -> pd.DataFrame:
    parts = [read_partition(pat, repo, ym) for ym in list_months(pat, repo)]
    parts = [p for p in parts if p is not None and len(p)]
    if not parts:
        return pd.DataFrame(columns=COLS)
    return pd.concat(parts, ignore_index=True).sort_values(["기준일", "판매처"]).reset_index(drop=True)


def _write_partition(pat, repo, ym, df, msg):
    buf = io.BytesIO()
    df.to_parquet(buf, index=False)
    url = "https://api.github.com/repos/%s/contents/%s" % (repo, _part_path(ym))
    sha = None
    try:
        sha = json.load(_gh(url + "?ref=main", pat)).get("sha")
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
    body = {"message": msg, "content": base64.b64encode(buf.getvalue()).decode("ascii")}
    if sha:
        body["sha"] = sha
    _gh(url, pat, method="PUT", data=body)


def ingest(df: pd.DataFrame, pat, repo) -> dict:
    """주문 DataFrame → 월별 파티션에 날짜구간 교체로 적재. 건드린 달만 재기록."""
    if df.empty:
        return {"rows": 0, "months": []}
    summary = []
    for ym, new in split_by_month(df).items():
        existing = read_partition(pat, repo, ym)
        before = len(existing)
        merged = date_range_replace(existing, new)
        _write_partition(pat, repo, ym, merged,
                         "data: easyadmin orders %s 적재 (%d→%d행)" % (ym, before, len(merged)))
        summary.append({"ym": ym, "before": before, "rows_in": len(new), "after": len(merged)})
    return {"rows": len(df), "months": summary}
