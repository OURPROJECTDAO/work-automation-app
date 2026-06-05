"""
송장번호 일괄입력 (invoice-fill) 워크플로우 — 채널별 송장 채우기.

원리 (전 채널 공통, 템플릿 양식만 다름):
  입력 = 채널 '송장번호 일괄입력 템플릿' (처리전, OLE2 .xls; 송장번호 빈칸)
         + 공통 송장 마스터(송장출력 .xlsx)
  처리 = VLOOKUP(처리전.<match_col> ⟷ 마스터.<master_key>) → 택배사·송장번호 (첫 매칭)
  N/A  = 매칭 실패 행 →
           ① 동일 배송지에 송장 채워진 행(박스) 존재 → 합포장 후보로 사용자 확인
                (같은 주소 박스 여럿이면 사용자가 어느 박스인지 선택; 확정 시 그 송장 복사)
           ② 없음(독립개체) 또는 사용자가 합포장 부정 → N/A 유지
  전달 = 잔존 N/A 행 전체 삭제 후 원본 .xls 양식 그대로 출력 + N/A 건수 보고

채널 추가 시: CHANNEL_CONFIG 에 (match_col, master_key) 한 줄 등록. (역추적 불필요)
출력은 반드시 원본 .xls 양식 (모든 판매처가 원본 양식만 허용).
"""
import io
import unicodedata
from collections import OrderedDict

import openpyxl


def nfc(s) -> str:
    if s is None:
        return ""
    return unicodedata.normalize("NFC", str(s)).strip()


def to_invoice_number(v):
    """송장번호를 숫자(int)로. 전부 숫자면 int, 아니면 원본 문자열 유지.
    openpyxl 등이 float로 읽어 '....0' 꼬리표가 붙은 경우도 정수화."""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return int(s) if s.isdigit() else (s if s else "")


# ── 채널 설정 ─────────────────────────────────────────────
# match_col  : 처리전(채널 템플릿)에서 VLOOKUP 키로 쓰는 컬럼명
# master_key : 송장 마스터에서 매칭되는 컬럼명 (채널마다 사용자가 알려줌)
CHANNEL_CONFIG = {
    "식봄": {"match_col": "상품주문번호", "master_key": "주문번호", "courier": "한진택배"},
    # "올웨이즈": {"match_col": "...", "master_key": "주문번호"},  # 샘플 받으면 추가
    # "배민상회": {"match_col": "...", "master_key": "주문번호"},
    # "캐시노트": {"match_col": "...", "master_key": "주문번호"},
}

# 송장 마스터(송장출력) 표준 컬럼
MASTER_COLS = ["상태", "관리번호", "발주일", "판매처", "주문번호",
               "수령자", "주소", "상품명", "택배사", "송장번호"]


# ── IO ───────────────────────────────────────────────────
def parse_template_xls(file_bytes: bytes) -> dict:
    """채널 송장 템플릿(.xls OLE2) 파싱. r0=헤더, r1=안내문, r2+=데이터.
    셀 타입을 보존해 출력 시 원본 양식 재현.
    """
    import xlrd
    book = xlrd.open_workbook(file_contents=file_bytes)
    sh = book.sheet_by_index(0)
    header = [sh.cell_value(0, c) for c in range(sh.ncols)]
    guide = [sh.cell_value(1, c) for c in range(sh.ncols)] if sh.nrows > 1 else [""] * sh.ncols
    rows, types = [], []
    for r in range(2, sh.nrows):
        rows.append({header[c]: sh.cell_value(r, c) for c in range(sh.ncols)})
        types.append([sh.cell_type(r, c) for c in range(sh.ncols)])
    return {"sheet_name": sh.name, "header": header, "guide": guide,
            "rows": rows, "types": types}


def parse_master(file_bytes: bytes) -> list:
    """공통 송장 마스터(.xlsx, 시트 '송장출력' 또는 첫 시트) → list[dict]."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb["송장출력"] if "송장출력" in wb.sheetnames else wb[wb.sheetnames[0]]
    rit = ws.iter_rows(values_only=True)
    hdr = list(next(rit))
    out = [{hdr[i]: row[i] for i in range(len(hdr))} for row in rit if any(v is not None for v in row)]
    wb.close()
    return out


def write_template_xls(parsed: dict, rows: list, keep_idx: list, song_col_name="송장번호", courier=None) -> bytes:
    """살아남은 행(keep_idx)만, 원본 .xls 양식(시트명·헤더·안내문·타입)으로 재작성."""
    import xlrd
    import xlwt
    header = parsed["header"]
    song_col = header.index(song_col_name)
    wbk = xlwt.Workbook(encoding="utf-8")
    sh = wbk.add_sheet(parsed["sheet_name"])
    for c, v in enumerate(header):
        sh.write(0, c, v)
    for c, v in enumerate(parsed["guide"]):
        sh.write(1, c, v)
    out_r = 2
    for i in keep_idx:
        row = rows[i]
        types = parsed["types"][i]
        for c in range(len(header)):
            if c == song_col:
                sh.write(out_r, c, to_invoice_number(row["_송장"]))   # 송장번호 = 숫자 형식
            elif header[c] == "택배사":
                # 택배사: 채널 지정 courier 일괄 기입(없으면 lookup 택배사)
                sh.write(out_r, c, courier or row.get("_택배사") or "")
            else:
                val = row.get(header[c])
                if types[c] == xlrd.XL_CELL_NUMBER:
                    sh.write(out_r, c, val)                 # 숫자 보존
                else:
                    sh.write(out_r, c, "" if val is None else val)
        out_r += 1
    buf = io.BytesIO()
    wbk.save(buf)
    return buf.getvalue()


# ── 로직 ─────────────────────────────────────────────────
def build_master_lookup(master_rows, master_key="주문번호"):
    """마스터 key → (택배사, 송장번호). VLOOKUP 의미상 '첫 매칭'만 유지(분할배송 첫 박스)."""
    lk = OrderedDict()
    for m in master_rows:
        k = nfc(m.get(master_key))
        if not k:
            continue
        if k not in lk:
            lk[k] = (str(m.get("택배사") or "").strip(),
                     str(m.get("송장번호") or "").strip())
    return lk


def vlookup_fill(before_rows, master_lookup, channel):
    """STEP1: VLOOKUP. 각 처리전 행에 _송장/_택배사/_status(matched|na) 부여."""
    mc = CHANNEL_CONFIG[channel]["match_col"]
    out = []
    for row in before_rows:
        r = dict(row)
        hit = master_lookup.get(nfc(r.get(mc)))
        if hit:
            r["_택배사"], r["_송장"], r["_status"] = hit[0], hit[1], "matched"
        else:
            r["_택배사"], r["_송장"], r["_status"] = "", None, "na"
        out.append(r)
    return out


def _recv(r):
    return str(r.get("수취인명(받는사람)") or r.get("수취인") or "").strip()


def find_consolidation_candidates(rows):
    """STEP2: N/A 행별 합포장 후보. 동일 배송지(NFC+trim)에 송장 채워진 박스가 있으면 후보.
    반환: (candidates, independents)
      candidates: [{na_index, na_row, boxes:[{송장,택배사,수취인,상품명}...]}]
      independents: [na_index ...]  (동일주소 박스 없음 → N/A 유지)
    """
    addr_boxes = {}
    for r in rows:
        if r["_status"] == "matched" and r["_송장"]:
            addr = nfc(r.get("배송지"))
            addr_boxes.setdefault(addr, OrderedDict())
            if r["_송장"] not in addr_boxes[addr]:
                addr_boxes[addr][r["_송장"]] = {
                    "송장": r["_송장"], "택배사": r["_택배사"],
                    "수취인": _recv(r), "상품명": str(r.get("상품명") or "").strip(),
                }
    candidates, independents = [], []
    for i, r in enumerate(rows):
        if r["_status"] != "na":
            continue
        boxes = list(addr_boxes.get(nfc(r.get("배송지")), {}).values())
        (candidates.append({"na_index": i, "na_row": r, "boxes": boxes})
         if boxes else independents.append(i))
    return candidates, independents


def apply_decisions(rows, decisions):
    """STEP3: 합포장 결정 반영. decisions: {na_index: 송장 or None}.
    송장 → 그 박스 송장/택배사 복사(status=consolidated), None → N/A 유지.
    """
    song_to_courier = {}
    for r in rows:
        if r["_송장"]:
            song_to_courier.setdefault(r["_송장"], r["_택배사"])
    for i, song in decisions.items():
        if song:
            rows[i]["_송장"] = song
            rows[i]["_택배사"] = song_to_courier.get(song, rows[i]["_택배사"])
            rows[i]["_status"] = "consolidated"
    return rows


def finalize(rows):
    """STEP4: 잔존 N/A 행 제외. 반환 (keep_idx, na_count, na_rows)."""
    keep_idx, na = [], []
    for i, r in enumerate(rows):
        if r["_status"] == "na" or not r["_송장"]:
            na.append(r)
        else:
            keep_idx.append(i)
    return keep_idx, len(na), na
