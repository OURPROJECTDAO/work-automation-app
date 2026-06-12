"""upload-monitor (업로드감시) — 코어 로직.

박스재고가 있는데 각 채널에 아직 등록(업로드) 안 된 관리코드 탐지
(+ 역으로 재고0인데 채널엔 살아있는 = 품절처리 대상).

핵심: 낱개(PC+상품코드)·박스(관리코드)·소분(변환-원코드)·합포(-CB-)가
전부 같은 product_master 행(상품코드)으로 환원된다(channel_margin_monitor
resolve_code 4-tier 재사용). → "하나라도 업로드면 업로드됨"이 상품코드 키로 자동 성립.

정규화 키 = 상품코드(유니크). 표시 = 관리코드. listing = 마진모니터 스냅샷 공유.
설계 = workflows/upload-monitor.md (ADR 0017).
"""
from __future__ import annotations

import csv
from pathlib import Path

from .channel_margin_monitor import _nfc, _num, load_references

# 8채널 (마진모니터 listing 보유) — (key, 표시라벨, register 자동폼 여부)
CHANNELS: list[tuple[str, str, bool]] = [
    ("smartstore", "스마트스토어", True),
    ("esm", "ESM", True),
    ("sikbom", "식봄", False),
    ("cashnote", "캐시노트", False),
    ("baemin", "배민상회", False),
    ("coupang", "쿠팡", False),
    ("allways", "올웨이즈", False),
    ("ali", "알리", False),
]
CHANNEL_KEYS = [c[0] for c in CHANNELS]
CHANNEL_LABEL = {c[0]: c[1] for c in CHANNELS}
REGISTER_AUTO = {c[0] for c in CHANNELS if c[2]}  # 자동 등록폼 가능 채널

# 상태 통제어휘
ST_OK = "이상없음"
ST_NEED_UP = "업로드필요"
ST_NEED_SOLD = "품절처리필요"
ST_SKIP = "업로드불필요"
ST_SKIP_CH = "업로드제외"      # 해당 채널 업로드x (사용자 지정, 업로드필요보다 우선)

# 비판매(회계·부자재) 제외. 반품/파렛트 중분류는 명확한 비판매라 코드 제외.
# 그 외 부자재(포장박스 '3번'·테이프·랩·환불상계 등)는 실분류(창고존, 예 '통조림-C동')에
# 섞여 있어 카테고리만으론 못 거름 → 사용자 유지 제외목록(아래)으로 보완(발견 시 추가).
EXCLUDE_MIDCAT = {"반품", "파렛트"}


def resolve_identity(code: str, refs: dict) -> list[str]:
    """판매자상품코드 → [상품코드, ...] (base 정체). resolve_code 분류 분기 재사용.

    합포(-CB-)=구성코드별 다중 / 소분·박스=단일 / PC낱개=코드[2:]=상품코드.
    미매칭·빈코드 = [].
    """
    c = _nfc(code)
    if not c:
        return []
    pm_m, pm_p, sobun = refs["pm_by_mgmt"], refs["pm_by_prod"], refs["sobun"]

    # 1) 합포 (코드1-CB-코드2[-CB-코드3])
    if "-CB-" in c:
        out: list[str] = []
        for p in c.split("-CB-"):
            r = pm_m.get(_nfc(p))
            if r:
                sc = _nfc(r.get("상품코드"))
                if sc:
                    out.append(sc)
        return out
    # 2) 소분 (변환관리코드 → 원코드 → 상품코드)
    if c in sobun:
        base = _nfc(sobun[c].get("원코드"))
        r = pm_m.get(base)
        if r:
            sc = _nfc(r.get("상품코드"))
            return [sc] if sc else []
        return []
    # 3) PC 낱개 (PC + 상품코드)
    if c.upper().startswith("PC"):
        sc = _nfc(c[2:])
        return [sc] if sc in pm_p else []
    # 4) 박스 (관리코드 → 상품코드)
    r = pm_m.get(c)
    if r:
        sc = _nfc(r.get("상품코드"))
        return [sc] if sc else []
    return []


def _listing_codes(ref_dir: Path, key: str) -> list[str]:
    """reference/listing_<key>.csv 의 '코드' 컬럼(관리코드) 리스트. 없으면 []."""
    p = Path(ref_dir) / f"listing_{key}.csv"
    if not p.exists():
        return []
    out: list[str] = []
    with open(p, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            out.append(row.get("코드", ""))
    return out


def _load_exclude(ref_dir) -> set[str]:
    """reference/upload_monitor_exclude.csv (상품코드 컬럼) — 사용자 유지 비판매 제외목록.

    포장박스·부자재·회계항목처럼 카테고리로 못 거르는 것을 발견 시 추가. 없으면 빈 set.
    """
    p = Path(ref_dir) / "upload_monitor_exclude.csv"
    if not p.exists():
        return set()
    out: set[str] = set()
    with open(p, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            sc = _nfc(row.get("상품코드"))
            if sc:
                out.add(sc)
    return out


def parse_skip_text(text: str) -> set[tuple[str, str]]:
    """upload_skip.csv 텍스트 → {(상품코드, 채널key)}. 채널별 업로드제외 쌍."""
    import io
    out: set[tuple[str, str]] = set()
    if not text:
        return out
    for row in csv.DictReader(io.StringIO(text)):
        sc, ch = _nfc(row.get("상품코드")), _nfc(row.get("채널"))
        if sc and ch:
            out.add((sc, ch))
    return out


def build_skip_text(pairs) -> str:
    """{(상품코드, 채널key)} → upload_skip.csv 텍스트(헤더·정렬·dedup)."""
    lines = ["상품코드,채널"]
    for sc, ch in sorted(set(pairs)):
        lines.append(f"{sc},{ch}")
    return "\n".join(lines) + "\n"


def _load_skip(ref_dir) -> set[tuple[str, str]]:
    """reference/upload_skip.csv → {(상품코드, 채널key)}. 없으면 빈 set."""
    p = Path(ref_dir) / "upload_skip.csv"
    if not p.exists():
        return set()
    with open(p, encoding="utf-8-sig") as f:
        return parse_skip_text(f.read())


def build_uploaded_sets(ref_dir, refs: dict, keys: list[str] | None = None) -> dict[str, set[str]]:
    """채널별 업로드된 상품코드 집합. {key: set(상품코드)}."""
    keys = keys or CHANNEL_KEYS
    uploaded: dict[str, set[str]] = {}
    for key in keys:
        s: set[str] = set()
        for code in _listing_codes(ref_dir, key):
            s.update(resolve_identity(code, refs))
        uploaded[key] = s
    return uploaded


def build_gap_table(ref_dir, refs: dict | None = None,
                    skip_pairs: set | None = None) -> list[dict]:
    """업로드감시 메인 테이블 (재고금액 desc).

    각 row: 상품코드·관리코드·상품명·박스재고·박스매입가·재고금액 + 채널키별 상태.
    상태 = 이상없음 / 업로드필요 / 품절처리필요 / 업로드불필요 / 업로드제외.
    - 노이즈(재고≤0 & 어디에도 미업로드) 행은 제외.
    - 채널별 업로드제외(skip): (상품코드,채널) 쌍은 그 채널이 '업로드필요'일 때 '업로드제외'로 덮음(우선).
      skip_pairs 인자가 있으면 그걸 쓰고(페이지 라이브 read), 없으면 reference/upload_skip.csv.
    """
    if refs is None:
        refs = load_references(ref_dir)
    uploaded = build_uploaded_sets(ref_dir, refs)
    exclude = _load_exclude(ref_dir)
    skip = _load_skip(ref_dir) if skip_pairs is None else set(skip_pairs)
    skip_by_sc: dict[str, set[str]] = {}
    for sc, ch in skip:
        skip_by_sc.setdefault(sc, set()).add(ch)

    rows: list[dict] = []
    for sc, r in refs["pm_by_prod"].items():
        if sc in exclude or _nfc(r.get("중분류명")) in EXCLUDE_MIDCAT:
            continue  # 비판매(반품·파렛트·사용자 제외목록)
        stock = _num(r.get("박스"))
        buy = _num(r.get("박스매입단가"))
        any_up = any(sc in uploaded[k] for k in CHANNEL_KEYS)
        if stock <= 0 and not any_up:
            continue  # 노이즈 제외
        sc_skip = skip_by_sc.get(sc, set())
        row = {
            "상품코드": sc,
            "관리코드": _nfc(r.get("관리코드")),
            "상품명": _nfc(r.get("상품명")),
            "박스재고": stock,
            "박스매입가": buy,
            "재고금액": stock * buy,
        }
        for key in CHANNEL_KEYS:
            up = sc in uploaded[key]
            if stock > 0:
                stt = ST_OK if up else ST_NEED_UP
            else:
                stt = ST_NEED_SOLD if up else ST_SKIP
            if stt == ST_NEED_UP and key in sc_skip:   # 채널별 업로드제외 우선
                stt = ST_SKIP_CH
            row[key] = stt
        rows.append(row)
    rows.sort(key=lambda x: x["재고금액"], reverse=True)
    return rows


def channel_summary(rows: list[dict]) -> list[dict]:
    """채널별 라이트 KPI: 업로드필요 · 품절처리필요 · 업로드제외 건수."""
    out = []
    for key, label, _ in CHANNELS:
        need = sum(1 for r in rows if r[key] == ST_NEED_UP)
        sold = sum(1 for r in rows if r[key] == ST_NEED_SOLD)
        skip = sum(1 for r in rows if r[key] == ST_SKIP_CH)
        out.append({"key": key, "label": label,
                    "업로드필요": need, "품절처리필요": sold, "업로드제외": skip})
    return out


def gap_list_for_channel(rows: list[dict], key: str) -> list[dict]:
    """특정 채널의 업로드필요 상품(재고금액 desc, 이미 정렬됨)."""
    return [r for r in rows if r.get(key) == ST_NEED_UP]
