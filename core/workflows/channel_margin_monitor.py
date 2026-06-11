"""channel-margin-monitor — 채널 가격·마진 모니터 (코어 로직).

채널 상품관리 다운로드(라이브 리스팅)를 받아 상품별 마진율 계산 →
기준마진 대비 이탈(탐지) + 기준마진 달성 권장가 역산(100원 올림 → 기준마진 이상 보장).

판매자상품코드 4-tier 해석:
  박스(관리코드) / PC낱개(PC+상품코드) / 소분(변환코드-원코드) / 합포(코드1-CB-코드2).
매입가 = (코드해석 base) × N,  N = 판매자바코드(빈값/0→1, 분수 가능).
정산액 = 판매가net×(1-수수료) + 배송비×정산계수.
이익   = 정산액 - 매입가 - 실택배비.   마진율 = 이익/정산액.
권장가 = ⌈((매입가+실택배비)/(1-확정마진율) - 배송비×정산계수)/(1-수수료)⌉ (100원 올림).

reference: product_master.csv · baseline_margin.csv · sobun.csv · margin_floor.csv (app reference/).
공식·근거 = workflows/channel-margin-monitor.md. (검증: 2026-06-10 골든 705/706)
"""
from __future__ import annotations

import csv
import math
import unicodedata
from io import BytesIO, StringIO
from pathlib import Path

from openpyxl import load_workbook

# ── 채널 config (채널 추가 = 여기에 한 세트) ────────────────────────────────
CHANNEL_CONFIG: dict[str, dict] = {
    "스마트스토어": {
        "key": "smartstore",       # 저장 파일명 reference/listing_<key>.csv
        "commission": 0.06,        # 판매수수료 → (1-수수료)=0.94 가 판매가에 곱
        "ship_settle": 0.967,      # 배송비 정산계수
        "real_ship": 2700,         # 실택배비 (단일)
        "baseline_col": "스마트스토어",  # baseline_margin.csv 의 채널 컬럼
        "apply_floor": True,       # 마진제한(하한 텍스트) 적용
        "sheet": None,             # None=첫 시트
        "header_row": 2,
        "data_start": 6,
        # 다운로드 컬럼 위치(1-indexed)
        "cols": {"상품번호": 1, "코드": 2, "상품명": 4, "판매가": 6,
                 "배송비": 41, "즉시할인": 58, "포인트": 69, "바코드": 78},
    },
}


# 마진미달 판정 임계 (탐지 = 마진율 − 기준마진율 < 이 값). -0.01 = 기준보다 1%p↑ 낮음.
MARGIN_UNDER_THRESHOLD = -0.01


def _nfc(s) -> str:
    return unicodedata.normalize("NFC", str(s)).strip() if s not in (None, "") else ""


def _num(v, d: float = 0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return d


# ── reference 로딩 ──────────────────────────────────────────────────────────
def load_references(ref_dir) -> dict:
    """app reference/ 에서 4종 로드 → dict."""
    ref_dir = Path(ref_dir)
    pm_by_mgmt: dict[str, dict] = {}
    pm_by_prod: dict[str, dict] = {}
    with open(ref_dir / "product_master.csv", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            mg, pr = _nfc(row.get("관리코드")), _nfc(row.get("상품코드"))
            if mg:
                pm_by_mgmt.setdefault(mg, row)
            if pr:
                pm_by_prod.setdefault(pr, row)

    def _load(name, key):
        d: dict[str, dict] = {}
        with open(ref_dir / name, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                k = _nfc(row.get(key))
                if k:
                    d.setdefault(k, row)
        return d

    return {
        "pm_by_mgmt": pm_by_mgmt,
        "pm_by_prod": pm_by_prod,
        "sobun": _load("sobun.csv", "변환관리코드"),
        "baseline": _load("baseline_margin.csv", "관리코드"),
        "floor": _load("margin_floor.csv", "관리코드"),
    }


# ── 코드 4-tier 해석 ────────────────────────────────────────────────────────
def resolve_code(code: str, refs: dict) -> tuple[str, float | None, float | None, str, str]:
    """returns (코드유형, base매입가, 재고, 규격, 비고).  base=None → 미매칭."""
    c = _nfc(code)
    if not c:
        return ("빈코드", None, None, "", "판매자상품코드 없음")
    pm_m, pm_p, sobun = refs["pm_by_mgmt"], refs["pm_by_prod"], refs["sobun"]
    # 1) 합포 (코드1-CB-코드2[-CB-코드3])
    if "-CB-" in c:
        prices, stocks, miss = [], [], []
        for p in c.split("-CB-"):
            r = pm_m.get(_nfc(p))
            if r:
                prices.append(_num(r["박스매입단가"]))
                stocks.append(_num(r["박스"]))
            else:
                miss.append(p)
        if miss:
            return ("합포", None, None, "", f"합포 구성코드 미등록: {','.join(miss)}")
        return ("합포", sum(prices) + 700, sum(stocks), "", "")
    # 2) 소분 (변환관리코드)
    if c in sobun:
        s = sobun[c]
        base = _nfc(s["원코드"])
        div = _num(s["내품나누기"], 0)
        r = pm_m.get(base)
        if not r or not div:
            return ("소분", None, None, _nfc(s.get("소분규격")), "소분 원코드 미등록/내품나누기 0")
        return ("소분", _num(r["박스매입단가"]) / div, _num(r["박스"]), _nfc(s.get("소분규격")), "")
    # 3) PC 낱개 (PC+상품코드)
    if c.upper().startswith("PC"):
        r = pm_p.get(_nfc(c[2:]))
        if not r:
            return ("낱개", None, None, "", f"상품코드 미등록: {c[2:]}")
        return ("낱개", _num(r["매입단가"]), _num(r["낱개"]), _nfc(r["규격"]), "")
    # 4) 박스 (관리코드)
    r = pm_m.get(c)
    if not r:
        return ("박스", None, None, "", "관리코드 미등록")
    return ("박스", _num(r["박스매입단가"]), _num(r["박스"]), _nfc(r["규격"]), "")


def _ceil100(x: float) -> int:
    return int(math.ceil(x / 100) * 100)


# ── 다운로드 파싱 ───────────────────────────────────────────────────────────
def parse_download(file, cfg: dict) -> list[dict]:
    """채널 상품관리 다운로드(.xlsx) → 레코드 리스트."""
    src = BytesIO(file) if isinstance(file, (bytes, bytearray)) else file
    wb = load_workbook(src, data_only=True)  # read_only 금지(pitfalls)
    ws = wb[cfg["sheet"]] if cfg.get("sheet") else wb[wb.sheetnames[0]]
    col = cfg["cols"]
    recs = []
    for r in range(cfg["data_start"], ws.max_row + 1):
        pid = ws.cell(r, col["상품번호"]).value
        if pid in (None, ""):
            continue
        recs.append({
            "상품번호": _nfc(pid),
            "코드": _nfc(ws.cell(r, col["코드"]).value),
            "상품명": _nfc(ws.cell(r, col["상품명"]).value),
            "판매가": _num(ws.cell(r, col["판매가"]).value),
            "배송비": _num(ws.cell(r, col["배송비"]).value),
            "즉시할인": _num(ws.cell(r, col["즉시할인"]).value),
            "포인트": _num(ws.cell(r, col["포인트"]).value),
            "바코드": ws.cell(r, col["바코드"]).value,
        })
    return recs


# ── 마진 계산 ───────────────────────────────────────────────────────────────
def compute(recs: list[dict], refs: dict, cfg: dict) -> list[dict]:
    comm, settle, ship = cfg["commission"], cfg["ship_settle"], cfg["real_ship"]
    bcol = cfg["baseline_col"]
    apply_floor = cfg.get("apply_floor", True)
    rate = 1 - comm  # 판매가에 곱하는 정산비율 (예 0.94)
    out = []
    for rec in recs:
        typ, base, stock, spec, note = resolve_code(rec["코드"], refs)
        n_raw = _num(rec["바코드"], 0)
        N = 1.0 if n_raw == 0 else n_raw  # 빈값/0 → 1, 분수 허용
        row = {
            "상품번호": rec["상품번호"], "관리코드": rec["코드"], "상품명": rec["상품명"],
            "규격": spec, "코드유형": typ, "N": N, "재고": stock,
            "매입가": None, "판매가": rec["판매가"], "배송비": rec["배송비"],
            "정산액": None, "마진율": None, "기준마진율": None, "탐지": None,
            "권장가": None, "제한": "", "비고": note,
        }
        # baseline 확정마진율 (판매자상품코드 직조인)
        bm = refs["baseline"].get(rec["코드"], {})
        bv = bm.get(bcol, "")
        base_margin = _num(bv, None) if bv not in (None, "") else None
        row["기준마진율"] = base_margin
        # 마진제한 텍스트
        fl = refs["floor"].get(rec["코드"]) if apply_floor else None
        if fl:
            row["제한"] = _nfc(fl.get("제한내용")) or _nfc(fl.get("비고"))

        if base is None:
            out.append(row)
            continue
        매입가 = base * N
        판매가net = rec["판매가"] - rec["즉시할인"] - rec["포인트"]
        정산액 = 판매가net * rate + rec["배송비"] * settle
        row["매입가"] = round(매입가)
        row["정산액"] = round(정산액)
        if 정산액 > 0:
            마진율 = (정산액 - 매입가 - ship) / 정산액
            row["마진율"] = round(마진율, 4)
            if base_margin is not None:
                row["탐지"] = round(마진율 - base_margin, 4)
        # 권장가 (기준마진 달성 판매가, 100원 올림 → 기준마진 이상 보장)
        if base_margin is not None and base_margin < 1:
            권장 = ((매입가 + ship) / (1 - base_margin) - rec["배송비"] * settle) / rate
            row["권장가"] = _ceil100(권장)
        out.append(row)
    return out


def _stats(rows: list[dict]) -> dict:
    margins = [r["마진율"] for r in rows if r["마진율"] is not None]
    return {
        "총건수": len(rows),
        "미매칭": sum(1 for r in rows if r["매입가"] is None),
        "미설정": sum(1 for r in rows if r["기준마진율"] is None and r["매입가"] is not None),
        "마진미달": sum(1 for r in rows if r["탐지"] is not None and r["탐지"] < MARGIN_UNDER_THRESHOLD),
        "제한상품": sum(1 for r in rows if r["제한"]),
        "평균마진율": round(sum(margins) / len(margins), 4) if margins else None,
    }


def compute_listing(recs: list[dict], channel: str, ref_dir) -> tuple[list[dict], dict]:
    """저장된 listing 레코드 + 채널 → (결과 레코드, 통계)."""
    if channel not in CHANNEL_CONFIG:
        raise ValueError(f"지원하지 않는 채널: {channel}")
    refs = load_references(ref_dir)
    rows = compute(recs, refs, CHANNEL_CONFIG[channel])
    return rows, _stats(rows)


def run(file, channel: str, ref_dir) -> tuple[list[dict], dict]:
    """다운로드(.xlsx) + 채널 → (결과 레코드, 통계)."""
    if channel not in CHANNEL_CONFIG:
        raise ValueError(f"지원하지 않는 채널: {channel}")
    recs = parse_download(file, CHANNEL_CONFIG[channel])
    return compute_listing(recs, channel, ref_dir)


# ── 저장 listing (연동데이터) 직렬화 / 병합 ──────────────────────────────────
LISTING_COLS = ["상품번호", "코드", "상품명", "판매가", "배송비", "즉시할인", "포인트", "바코드"]


def recs_to_csv(recs: list[dict]) -> str:
    """parse_download 레코드 → CSV 텍스트 (저장용)."""
    buf = StringIO()
    w = csv.DictWriter(buf, fieldnames=LISTING_COLS)
    w.writeheader()
    for r in recs:
        bar = r.get("바코드")
        w.writerow({
            "상품번호": r["상품번호"], "코드": r["코드"], "상품명": r["상품명"],
            "판매가": r["판매가"], "배송비": r["배송비"],
            "즉시할인": r["즉시할인"], "포인트": r["포인트"],
            "바코드": "" if bar in (None, "") else bar,
        })
    return buf.getvalue()


def csv_text_to_recs(text: str) -> list[dict]:
    """저장 CSV 텍스트 → parse_download 호환 레코드."""
    recs = []
    for row in csv.DictReader(StringIO(text)):
        recs.append({
            "상품번호": _nfc(row.get("상품번호")), "코드": _nfc(row.get("코드")),
            "상품명": _nfc(row.get("상품명")), "판매가": _num(row.get("판매가")),
            "배송비": _num(row.get("배송비")), "즉시할인": _num(row.get("즉시할인")),
            "포인트": _num(row.get("포인트")), "바코드": row.get("바코드") or "",
        })
    return recs


def merge_listing(existing: list[dict], new: list[dict]) -> tuple[list[dict], int]:
    """기존 + 신규: 기존 상품번호는 유지, 새 상품번호만 추가. → (병합, 추가건수)."""
    seen = {r["상품번호"] for r in existing}
    added = [r for r in new if r["상품번호"] and r["상품번호"] not in seen]
    return existing + added, len(added)


# ── 가격 일괄변경 (할인 우선 규칙) ───────────────────────────────────────────
def adjust_price(판매가: float, 즉시할인: float, 포인트: float,
                 target_net: float) -> tuple[int, int]:
    """net(=판매가-즉시할인-포인트)을 target_net으로 맞추는 (새 판매가, 새 즉시할인).

    ★ 할인 우선: 인상 시 즉시할인을 먼저 줄이고 모자라면 판매가를 올린다.
       인하 시 즉시할인을 먼저 늘리고 모자라면 판매가를 내린다. 포인트는 불변.
    """
    판매가, 즉시할인, 포인트 = float(판매가), float(즉시할인), float(포인트)
    cur_net = 판매가 - 즉시할인 - 포인트
    delta = target_net - cur_net
    new_price, new_disc = 판매가, 즉시할인
    if delta > 0:                                   # 인상: 할인 축소 우선
        cut = min(즉시할인, delta)
        new_disc = 즉시할인 - cut
        new_price = 판매가 + (delta - cut)
    elif delta < 0:                                 # 인하: 할인 확대 우선
        need = -delta
        room = max(cur_net, 0.0)                    # net 0까지만 할인 가능
        add = min(need, room)
        new_disc = 즉시할인 + add
        new_price = 판매가 - (need - add)
    return int(round(new_price)), int(round(new_disc))


def compute_new_prices(rows: list[dict], recs: list[dict],
                       pids: set) -> tuple[dict, list[str]]:
    """체크된 상품번호(pids) → {상품번호: (새 판매가, 새 즉시할인)} + 건너뛴 목록.

    target = 권장가(기준마진 달성가). 권장가 없는(미매칭/기준미설정) 상품은 skip.
    """
    rec_by = {r["상품번호"]: r for r in recs}
    row_by = {r["상품번호"]: r for r in rows}
    new_prices, skipped = {}, []
    for pid in pids:
        row, rec = row_by.get(pid), rec_by.get(pid)
        if not row or not rec or row.get("권장가") is None:
            skipped.append(pid)
            continue
        np_, nd_ = adjust_price(rec["판매가"], rec["즉시할인"], rec["포인트"],
                                row["권장가"])
        new_prices[pid] = (np_, nd_)
    return new_prices, skipped


def build_bulk_price_xlsx(raw_xlsx: bytes, new_prices: dict,
                          cfg: dict) -> tuple[bytes, int, list[str]]:
    """원본 일괄변경 양식(raw_xlsx, 전체 컬럼) → 체크 상품 행만 가격 수정 후 남김.

    헤더(data_start 이전 행) 보존. new_prices 에 없는 데이터 행은 삭제.
    returns (xlsx bytes, 남긴 행수, 원본에 없던 상품번호 목록).
    """
    wb = load_workbook(BytesIO(raw_xlsx))           # 값+서식 보존 (read_only 금지)
    ws = wb[cfg["sheet"]] if cfg.get("sheet") else wb[wb.sheetnames[0]]
    col = cfg["cols"]
    c_pid, c_price, c_disc = col["상품번호"], col["판매가"], col["즉시할인"]
    c_unit = c_disc + 1                             # 즉시할인 단위(BG=BF+1)
    start = cfg["data_start"]
    found, drop = set(), []
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, c_pid).value
        pid = _nfc(v) if v not in (None, "") else ""
        if pid and pid in new_prices:
            price, disc = new_prices[pid]
            ws.cell(r, c_price).value = price
            if disc and disc > 0:
                ws.cell(r, c_disc).value = disc
                ws.cell(r, c_unit).value = "원"
            else:
                ws.cell(r, c_disc).value = None
                ws.cell(r, c_unit).value = None
            found.add(pid)
        else:
            drop.append(r)                          # 미체크 행 + 빈행 삭제
    for r in sorted(drop, reverse=True):
        ws.delete_rows(r, 1)
    out = BytesIO()
    wb.save(out)
    missing = [p for p in new_prices if p not in found]
    return out.getvalue(), len(found), missing


def append_rows_to_raw(raw_xlsx: bytes, src_xlsx: bytes,
                       pids: set, cfg: dict) -> bytes:
    """저장 원본(raw)에 src 양식의 신규 상품번호(pids) 행을 값으로 추가."""
    tgt = load_workbook(BytesIO(raw_xlsx))
    tws = tgt[cfg["sheet"]] if cfg.get("sheet") else tgt[tgt.sheetnames[0]]
    src = load_workbook(BytesIO(src_xlsx), data_only=True)
    sws = src[cfg["sheet"]] if cfg.get("sheet") else src[src.sheetnames[0]]
    c_pid = cfg["cols"]["상품번호"]
    start = cfg["data_start"]
    src_rows = {}
    for r in range(start, sws.max_row + 1):
        v = sws.cell(r, c_pid).value
        pid = _nfc(v) if v not in (None, "") else ""
        if pid:
            src_rows[pid] = r
    ncol = max(tws.max_column, sws.max_column)
    dest = tws.max_row + 1
    for pid in pids:
        sr = src_rows.get(pid)
        if not sr:
            continue
        for c in range(1, ncol + 1):
            tws.cell(dest, c).value = sws.cell(sr, c).value
        dest += 1
    out = BytesIO()
    tgt.save(out)
    return out.getvalue()
