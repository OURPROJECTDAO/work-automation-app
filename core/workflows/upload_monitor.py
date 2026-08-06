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

import concurrent.futures as _cf
import csv
import urllib.request as _ur
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


# ── 파생코드 (소분 / PC낱개) ─────────────────────────────────────────────────
# 원박스 관리코드에서 파생돼 "실제로 등록된" 코드만 보여준다(표시 전용·판정 무관).
#   소분   = sobun.csv(변환관리코드→원코드) — resolve_identity가 쓰는 정본과 동일 소스.
#   PC낱개 = unit_list.csv·sub_list.csv 의 PC 행(원코드 키). PC+상품코드 규칙상 코드는 늘
#            만들 수 있지만, 여기선 **등록된 것만** — 미등록은 빈칸(신규 채번 대상).
# 복수면 ', ' 결합. 판정(업로드필요/이상없음)에는 일절 관여하지 않는다.
DERIVED_COLS = ["소분코드", "PC코드"]


def _pc_origin_rows(ref_dir):
    """unit_list·sub_list → (관리코드, 원코드) 제너레이터. 파일 없으면 건너뜀."""
    for name in ("unit_list.csv", "sub_list.csv"):
        p = Path(ref_dir) / name
        if not p.exists():
            continue
        with open(p, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                yield _nfc(row.get("관리코드")), _nfc(row.get("원코드"))


def derived_code_map(ref_dir, refs: dict) -> dict[str, dict[str, str]]:
    """원박스 관리코드 → {"소분코드": str, "PC코드": str}. 없으면 빈 문자열."""
    sub: dict[str, list[str]] = {}
    for code, r in (refs.get("sobun") or {}).items():
        origin, c = _nfc(r.get("원코드")), _nfc(code)
        if origin and c and c not in sub.setdefault(origin, []):
            sub[origin].append(c)
    pc: dict[str, list[str]] = {}
    for code, origin in _pc_origin_rows(ref_dir):
        if origin and code.upper().startswith("PC") and code not in pc.setdefault(origin, []):
            pc[origin].append(code)
    return {k: {"소분코드": ", ".join(sorted(sub.get(k, []))),
                "PC코드": ", ".join(sorted(pc.get(k, [])))}
            for k in set(sub) | set(pc)}


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

    각 row: 상품코드·관리코드·소분코드·PC코드·상품명·박스재고·박스매입가·재고금액 + 채널키별 상태.
    소분코드/PC코드 = 그 관리코드에서 파생돼 등록된 코드(표시 전용·미등록이면 빈칸).
    상태 = 이상없음 / 업로드필요 / 품절처리필요 / 업로드불필요 / 업로드제외.
    - 노이즈(재고≤0 & 어디에도 미업로드) 행은 제외.
    - 채널별 업로드제외(skip): (상품코드,채널) 쌍은 그 채널이 '업로드필요'일 때 '업로드제외'로 덮음(우선).
      skip_pairs 인자가 있으면 그걸 쓰고(페이지 라이브 read), 없으면 reference/upload_skip.csv.
    """
    if refs is None:
        refs = load_references(ref_dir)
    uploaded = build_uploaded_sets(ref_dir, refs)
    derived = derived_code_map(ref_dir, refs)
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
        mg = _nfc(r.get("관리코드"))
        dv = derived.get(mg, {})
        row = {
            "상품코드": sc,
            "관리코드": mg,
            "소분코드": dv.get("소분코드", ""),
            "PC코드": dv.get("PC코드", ""),
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


# ── 이미지 실검사 (대표 A1 / 상세 B1, jpg→png) ───────────────────────────────
# 호스트 패턴: gi.esmplus.com/td680708/{관리코드}/{관리코드}_A1|B1.{ext}. 키=관리코드.
# 등록 워크플로우 공통 패턴(smartstore/esm-register)과 동일. A1/B1 독립 프로브.
IMG_HOST = "https://gi.esmplus.com/td680708"
IMG_COLS = ["대표이미지유무", "대표확장자", "대표이미지URL",
            "상세이미지유무", "상세확장자", "상세이미지URL"]


def _img_url(mg: str, slot: str, ext: str) -> str:
    return f"{IMG_HOST}/{mg}/{mg}_{slot}.{ext}"


def _head_ok(u: str, timeout: float = 6.0) -> bool:
    req = _ur.Request(u, method="HEAD")
    req.add_header("User-Agent", "Mozilla/5.0")
    try:
        with _ur.urlopen(req, timeout=timeout) as r:
            return r.status == 200
    except Exception:
        return False


def probe_slot(mg: str, slot: str) -> tuple[str, str]:
    """관리코드+슬롯(A1/B1) → (확장자, URL). jpg→png 실검사. 없으면 ('','')."""
    mg = _nfc(mg)
    if not mg:
        return "", ""
    for ext in ("jpg", "png"):
        u = _img_url(mg, slot, ext)
        if _head_ok(u):
            return ext, u
    return "", ""


def probe_image(mg: str) -> dict:
    """관리코드 → 대표(A1)·상세(B1) 유무/확장자/URL dict (IMG_COLS)."""
    a_ext, a_url = probe_slot(mg, "A1")
    b_ext, b_url = probe_slot(mg, "B1")
    return {
        "대표이미지유무": "O" if a_url else "X", "대표확장자": a_ext, "대표이미지URL": a_url,
        "상세이미지유무": "O" if b_url else "X", "상세확장자": b_ext, "상세이미지URL": b_url,
    }


def probe_images(codes, workers: int = 24) -> dict[str, dict]:
    """관리코드 리스트 → {관리코드: probe_image dict}. 빈값 제외·dedup·병렬."""
    uniq = sorted({_nfc(c) for c in codes if _nfc(c)})
    out: dict[str, dict] = {}
    if not uniq:
        return out
    with _cf.ThreadPoolExecutor(max_workers=min(workers, len(uniq))) as ex:
        futs = {ex.submit(probe_image, c): c for c in uniq}
        for f in _cf.as_completed(futs):
            out[futs[f]] = f.result()
    return out
