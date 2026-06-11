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
import random
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
        "unitprice_use_col": 7,    # G 단위가격 사용여부: 양식 출력 시 비었으면 'N' 채움
    },
    "식봄": {
        "key": "sikbom",
        "commission": 0.07,        # 식봄 수수료 7% → (1-수수료)=0.93
        "ship_settle": 0.967,      # 배송비 정산계수 (스마트스토어 동일)
        "real_ship": 2700,         # 실택배비 (스마트스토어 기준 단일 — 골든 3000/3700 폐기)
        "ship_fee_const": 3000,    # 식봄 다운로드엔 '배송비명'뿐(숫자 없음) → income측 배송비 상수
        "baseline_col": "식봄",     # baseline_margin.csv 식봄 컬럼
        "apply_floor": True,
        "n_source": "ref",         # 합포량 N = hapo_multiplier(상품번호) — 다운로드 바코드 없음
        "sheet": "식봄붙여넣기",
        "header_row": 4,
        "data_start": 5,
        # 다운로드 컬럼(1-indexed). 정가=권장가 산출 시 정가≥판매단가 보존용. 즉시할인·포인트·배송비·바코드 없음
        "cols": {"상품번호": 1, "코드": 2, "상품명": 6, "판매가": 19, "정가": 16},
        # 가격 일괄변경 = 다운로드와 별개 '상품 일괄수정' 양식에 선택 행을 채워 넣는 append 방식
        "price_form": {
            "mode": "append",                       # 템플릿에 선택 행만 기입(스마트스토어=filter와 다름)
            "template": "sikbom_price_template.xlsx",  # reference/ 고정 양식
            "sheet": "(식봄)양식",
            "data_start": 7,                        # r1~3 안내·r4~6 헤더/설명
            "cols": {"상품번호": 1, "코드": 2, "상품명": 3, "정가": 4, "판매단가": 6},
            "fixed": {5: "n"},                      # E열 수량별 판매단가 설정 = n 고정
            "source": {"상품번호": "상품번호", "코드": "관리코드", "상품명": "상품명",
                       "정가": "정가", "판매단가": "권장가"},
            "price_field": "판매단가",
            "jeong_field": "정가",
        },
    },
    "캐시노트": {
        "key": "cashnote",
        "commission": 0.06,        # 6% (요율컬럼=6·천년경영 0.94·헤더 '6%기준' 근거. 골든 정산식 0.93은
                                   #   시트 내부 불일치로 미채택. 행사 차등수수료 0.88(12%)는 무시=단일 수수료, 사용자 확정 2026-06-11)
        "ship_settle": 0.967,      # 배송비 정산계수 (전 채널 동일)
        "real_ship": 2700,         # 실택배비 (스마트스토어 표준 단일 — 골든 3000/3700 미채택)
        "baseline_col": "캐시노트",  # baseline_margin.csv 캐시노트 컬럼
        "apply_floor": True,
        "n_source": "ref",         # 합포량 N = hapo_multiplier(상품번호) 채널무관 — 다운로드 바코드 없음
        "sheet": "상품",
        "header_row": 3,
        "data_start": 4,
        # 다운로드 컬럼(1-indexed). A=ID(상품번호)·E=입점사 관리 코드·C=상품명·N=판매 단가·O=할인 전 단가(정가).
        # 즉시할인·포인트·바코드 컬럼 없음(식봄형).
        "cols": {"상품번호": 1, "코드": 5, "상품명": 3, "판매가": 14, "정가": 15},
        # 배송비 = 배송정책코드(Y열=25) 조건부: DVP212991→3000, 그 외(DVP447716 등)→0. 골든 J식과 일치.
        "ship_fee_policy": {"col": 25, "map": {"DVP212991": 3000}, "default": 0},
        # 가격변경 양식(A=오퍼코드 OFR·D=옵션코드 SKU)이 다운로드 Q(17)·R(18)에만 있어 listing에 보존.
        "extra_cols": {"오퍼코드": 17, "옵션코드": 18},
        # 가격 일괄변경 = '(캐시노트)양식' append. F=수정·L=Y·N=9999 고정, G=판매단가(권장가)·H=할인전단가(≥판매단가).
        "price_form": {
            "mode": "append",
            "template": "cashnote_price_template.xlsx",   # reference/ 고정 양식(업로드 폼)
            "sheet": "(캐시노트)양식",
            "data_start": 4,                              # r1~3 그룹헤더/안내, r2=컬럼명
            "cols": {"오퍼코드": 1, "옵션코드": 4, "판매단가": 7, "할인전단가": 8, "관리코드": 15},
            "fixed": {6: "수정", 12: "Y", 14: 9999},       # F 변경타입·L 진열여부·N 재고수량
            "source": {"오퍼코드": "오퍼코드", "옵션코드": "옵션코드", "관리코드": "관리코드",
                       "할인전단가": "정가", "판매단가": "권장가"},
            "price_field": "판매단가",
            "jeong_field": "할인전단가",
            # 할인전단가(H)는 무늬용 가짜 정가 — 실제 판매가는 G. 권장가 기준 +20~30% 랜덤·100원 반올림(>판매단가).
            # (마진/대시보드는 항상 판매단가만 사용. 실제 정가 보존이 아니라 매번 생성 → 일부 채널만.)
            "jeong_fake": {"min_pct": 0.20, "max_pct": 0.30, "round": 100},
        },
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


def _pid(v) -> str:
    """상품번호 정규화: 정수값 float(엑셀 숫자셀 46903.0)는 '46903'으로.

    캐시노트 등 다운로드의 상품번호(ID)가 숫자셀로 들어와 _nfc만 쓰면 '46903.0'이
    되어 hapo_multiplier 키('46903')·골든과 매칭 실패 → N=1 오류. 정수 float만 int화.
    """
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return _nfc(v)


def _pick_ws(wb, cfg):
    """cfg['sheet'] 가 있고 존재하면 그 시트, 아니면 첫 시트.

    채널 다운로드의 실제 시트명이 cfg와 다를 수 있어 첫 시트 폴백(예: 식봄 신규
    다운로드 시트명 != '식봄붙여넣기'). 식봄 다운로드는 단일 시트라 폴백 안전.
    """
    name = cfg.get("sheet")
    if name and name in wb.sheetnames:
        return wb[name]
    return wb[wb.sheetnames[0]]


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

    # 합포량(N) — 상품번호별 판매배수. 바코드 없는 채널 공용(마진율 예외). 파일 없으면 빈 dict.
    hapo: dict[str, float] = {}
    hp = ref_dir / "hapo_multiplier.csv"
    if hp.exists():
        with open(hp, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                k = _nfc(row.get("상품번호"))
                if k and k not in hapo:
                    hapo[k] = _num(row.get("합포량"), 1.0)

    return {
        "pm_by_mgmt": pm_by_mgmt,
        "pm_by_prod": pm_by_prod,
        "sobun": _load("sobun.csv", "변환관리코드"),
        "baseline": _load("baseline_margin.csv", "관리코드"),
        "floor": _load("margin_floor.csv", "관리코드"),
        "hapo": hapo,
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
        # 재고 = 그 상품코드 행의 **박스** 재고(낱개[15] 아님). 매입가는 낱개 매입단가 유지.
        return ("낱개", _num(r["매입단가"]), _num(r["박스"]), _nfc(r["규격"]), "")
    # 4) 박스 (관리코드)
    r = pm_m.get(c)
    if not r:
        return ("박스", None, None, "", "관리코드 미등록")
    return ("박스", _num(r["박스매입단가"]), _num(r["박스"]), _nfc(r["규격"]), "")


def _ceil100(x: float) -> int:
    return int(math.ceil(x / 100) * 100)


def _ranges_desc(nums: list) -> list:
    """정렬 정수들 → 연속 구간 [(start,end)] 내림차순(아래부터 삭제용)."""
    if not nums:
        return []
    nums = sorted(set(nums))
    ranges, s, p = [], nums[0], nums[0]
    for n in nums[1:]:
        if n == p + 1:
            p = n
        else:
            ranges.append((s, p)); s = p = n
    ranges.append((s, p))
    return list(reversed(ranges))


# ── 다운로드 파싱 ───────────────────────────────────────────────────────────
def parse_download(file, cfg: dict) -> list[dict]:
    """채널 상품관리 다운로드(.xlsx) → 레코드 리스트.

    채널별로 없는 컬럼(즉시할인·포인트·배송비·바코드)은 cfg['cols']에서 생략 가능
    → 0/상수/None 처리. 배송비 출처 3종(우선순위):
      ① cfg['ship_fee_const'] 상수(식봄) ② cfg['ship_fee_policy'] 배송정책코드 등
      컬럼값 조건부(캐시노트: DVP212991→3000, 그외→0) ③ cfg['cols']['배송비'] 숫자(스마트스토어).
    """
    src = BytesIO(file) if isinstance(file, (bytes, bytearray)) else file
    wb = load_workbook(src, data_only=True)  # read_only 금지(pitfalls)
    ws = _pick_ws(wb, cfg)
    col = cfg["cols"]
    ship_const = cfg.get("ship_fee_const")
    ship_policy = cfg.get("ship_fee_policy")  # {col, map, default} — 컬럼값 조건부 배송비(캐시노트)

    def _opt(r, key, default=0.0):
        c = col.get(key)
        return _num(ws.cell(r, c).value, default) if c else default

    def _ship(r):
        if ship_const is not None:
            return float(ship_const)
        if ship_policy:
            pol = _nfc(ws.cell(r, ship_policy["col"]).value)
            return float(ship_policy["map"].get(pol, ship_policy.get("default", 0)))
        return _opt(r, "배송비")

    recs = []
    for r in range(cfg["data_start"], ws.max_row + 1):
        pid = ws.cell(r, col["상품번호"]).value
        if pid in (None, ""):
            continue
        bc = col.get("바코드")
        rec = {
            "상품번호": _pid(pid),
            "코드": _nfc(ws.cell(r, col["코드"]).value),
            "상품명": _nfc(ws.cell(r, col["상품명"]).value),
            "판매가": _opt(r, "판매가"),
            "배송비": _ship(r),
            "즉시할인": _opt(r, "즉시할인"),
            "포인트": _opt(r, "포인트"),
            "정가": _opt(r, "정가"),
            "바코드": ws.cell(r, bc).value if bc else None,
            "오퍼코드": "", "옵션코드": "",          # 가격변경 양식 A/D용(캐시노트). 그 외 채널 공백
        }
        for name, c in cfg.get("extra_cols", {}).items():   # 다운로드 추가 컬럼 보존(OFR/SKU 등)
            rec[name] = _nfc(ws.cell(r, c).value)
        recs.append(rec)
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
        if cfg.get("n_source") == "ref":
            nv = refs.get("hapo", {}).get(rec["상품번호"], 1.0)  # 합포량(상품번호) 기본 1
            N = 1.0 if not nv else nv
        else:
            n_raw = _num(rec["바코드"], 0)
            N = 1.0 if n_raw == 0 else n_raw  # 빈값/0 → 1, 분수 허용
        row = {
            "상품번호": rec["상품번호"], "관리코드": rec["코드"], "상품명": rec["상품명"],
            "규격": spec, "코드유형": typ, "N": N, "재고": stock,
            "매입가": None, "판매가": rec["판매가"], "정가": rec.get("정가", 0), "배송비": rec["배송비"],
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
LISTING_COLS = ["상품번호", "코드", "상품명", "판매가", "정가", "배송비", "즉시할인", "포인트", "바코드",
                "오퍼코드", "옵션코드"]


def recs_to_csv(recs: list[dict]) -> str:
    """parse_download 레코드 → CSV 텍스트 (저장용)."""
    buf = StringIO()
    w = csv.DictWriter(buf, fieldnames=LISTING_COLS)
    w.writeheader()
    for r in recs:
        bar = r.get("바코드")
        w.writerow({
            "상품번호": r["상품번호"], "코드": r["코드"], "상품명": r["상품명"],
            "판매가": r["판매가"], "정가": r.get("정가", ""), "배송비": r["배송비"],
            "즉시할인": r["즉시할인"], "포인트": r["포인트"],
            "바코드": "" if bar in (None, "") else bar,
            "오퍼코드": r.get("오퍼코드", ""), "옵션코드": r.get("옵션코드", ""),
        })
    return buf.getvalue()


def csv_text_to_recs(text: str) -> list[dict]:
    """저장 CSV 텍스트 → parse_download 호환 레코드."""
    recs = []
    for row in csv.DictReader(StringIO(text)):
        recs.append({
            "상품번호": _nfc(row.get("상품번호")), "코드": _nfc(row.get("코드")),
            "상품명": _nfc(row.get("상품명")), "판매가": _num(row.get("판매가")),
            "정가": _num(row.get("정가")),
            "배송비": _num(row.get("배송비")), "즉시할인": _num(row.get("즉시할인")),
            "포인트": _num(row.get("포인트")), "바코드": row.get("바코드") or "",
            "오퍼코드": _nfc(row.get("오퍼코드")), "옵션코드": _nfc(row.get("옵션코드")),
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


def build_append_items(pf: dict, rows: list[dict], recs: list[dict],
                       pids) -> tuple[list[dict], list[dict], list[str]]:
    """append형 가격변경 양식의 (items, preview, skipped) 생성 — 채널 무관.

    pf['source'] {양식필드: 소스키}: row(우선)/rec 에서 값 추출.
    pf['price_field']: 권장가가 들어갈 양식필드(정수).
    pf['jeong_field']: (선택) 정가/할인전단가 필드 → max(소스값, 판매단가) 보장(정가≥판매가).
    권장가 없는(미매칭/기준 미설정) 상품은 skip.
    """
    row_by = {r["상품번호"]: r for r in rows}
    rec_by = {r["상품번호"]: r for r in recs}
    src = pf.get("source", {})
    price_f = pf["price_field"]
    jeong_f = pf.get("jeong_field")
    items, preview, skipped = [], [], []
    for pid in pids:
        ro = row_by.get(pid)
        if not ro or ro.get("권장가") is None:
            skipped.append(pid)
            continue
        merged = {**rec_by.get(pid, {}), **ro}       # row 우선
        price = int(ro["권장가"])
        it = {field: merged.get(key, "") for field, key in src.items()}
        it[price_f] = price
        if jeong_f:
            fake = pf.get("jeong_fake")
            if fake:                                  # 무늬용 가짜 정가: 판매가 +20~30% 랜덤·단위 반올림(>판매가)
                pct = random.uniform(fake.get("min_pct", 0.20), fake.get("max_pct", 0.30))
                unit = int(fake.get("round", 100))
                val = int(round(price * (1 + pct) / unit) * unit)
                it[jeong_f] = val if val > price else price + unit
            else:
                it[jeong_f] = int(max(_num(it.get(jeong_f)), price))   # 실제 정가 보존(식봄)
        items.append(it)
        cur = int(_num(ro.get("판매가")))
        preview.append({
            "상품명": ro.get("상품명"), "현재판매가": cur, "새판매단가": price,
            "정가": it.get(jeong_f) if jeong_f else "",
            "방향": "인상" if price > cur else ("인하" if price < cur else "유지"),
        })
    return items, preview, skipped


def build_price_form_append(template_xlsx: bytes, items: list[dict], pf: dict) -> bytes:
    """채널 '가격변경 양식' 템플릿에 선택 상품 행만 채워 append (식봄·캐시노트형).

    items: [{양식필드: 값}] — build_append_items 가 판매단가(=권장가)·정가/할인전단가까지
    계산해 넣는다. pf['cols'] {양식필드: 컬럼} 로 기입, pf['fixed'] {컬럼: 값} 고정값.
    템플릿의 기존/예시 데이터행은 모두 제거하고 data_start부터 기입.
    빈행 방지 위해 keep_last 초과 row_dimensions 정리(전역 pitfalls).
    """
    wb = load_workbook(BytesIO(template_xlsx))
    ws = wb[pf["sheet"]] if pf.get("sheet") else wb[wb.sheetnames[0]]
    cols = pf["cols"]
    start = pf["data_start"]
    fixed = pf.get("fixed", {})
    if ws.max_row >= start:                          # 예시/기존 데이터행 제거
        ws.delete_rows(start, ws.max_row - start + 1)
    for i, it in enumerate(items):
        r = start + i
        for field, c in cols.items():
            if field in it:
                ws.cell(r, c).value = it[field]
        for c, val in fixed.items():                 # 고정값(예: 변경타입 '수정'·진열 'Y'·재고 9999)
            ws.cell(r, int(c)).value = val
    keep_last = start - 1 + len(items)
    for rr in [x for x in ws.row_dimensions if x > keep_last]:
        del ws.row_dimensions[rr]
    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def build_bulk_price_xlsx(raw_xlsx: bytes, new_prices: dict,
                          cfg: dict) -> tuple[bytes, int, list[str]]:
    """원본 일괄변경 양식(raw_xlsx, 전체 컬럼) → 체크 상품 행만 가격 수정 후 남김.

    헤더(data_start 이전 행) 보존. new_prices 에 없는 데이터 행은 삭제.
    returns (xlsx bytes, 남긴 행수, 원본에 없던 상품번호 목록).
    """
    wb = load_workbook(BytesIO(raw_xlsx))           # 값+서식 보존 (read_only 금지)
    ws = _pick_ws(wb, cfg)
    col = cfg["cols"]
    c_pid, c_price, c_disc = col["상품번호"], col["판매가"], col["즉시할인"]
    c_unit = c_disc + 1                             # 즉시할인 단위(BG=BF+1)
    c_up = cfg.get("unitprice_use_col")             # 단위가격 사용여부(G) — 있을 때만
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
            if c_up and ws.cell(r, c_up).value in (None, ""):
                ws.cell(r, c_up).value = "N"        # 비었으면 N, 값 있으면 보존
            found.add(pid)
        else:
            drop.append(r)                          # 미체크 행 + 빈행 → 삭제(업로드 시 빈행 방지)
    for a, b in _ranges_desc(drop):                 # 연속 구간 묶어 아래부터 삭제
        ws.delete_rows(a, b - a + 1)
    # delete_rows가 남기는 빈 row_dimensions(빈 <row> 요소 유발) 정리 — 마지막 데이터행 이후 제거
    keep_last = (start - 1) + len(found)
    for rr in [x for x in ws.row_dimensions if x > keep_last]:
        del ws.row_dimensions[rr]
    out = BytesIO()
    wb.save(out)
    missing = [p for p in new_prices if p not in found]
    return out.getvalue(), len(found), missing


def append_rows_to_raw(raw_xlsx: bytes, src_xlsx: bytes,
                       pids: set, cfg: dict) -> bytes:
    """저장 원본(raw)에 src 양식의 신규 상품번호(pids) 행을 값으로 추가."""
    tgt = load_workbook(BytesIO(raw_xlsx))
    tws = _pick_ws(tgt, cfg)
    src = load_workbook(BytesIO(src_xlsx), data_only=True)
    sws = _pick_ws(src, cfg)
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
