"""📅 데일리 대시보드 — 매일 반복 업무 산출물로 당일 인사이트 (구 마진침식 탭D 승격).

천년경영업로드 output + 송장출력 + 상품관리(master)로 오늘 판 것의 대략 실현마진을 즉시 점검.
- 매출 = 천년경영 실제기입단가×수량(net·수수료적용) · 원가 = master 매입가(낱개)×낱개수량
- 택배 = 채널 flat × 실제 물리박스(합포 2시나리오: 250/355 H열 다품목 + 175~200ml 30개입 수령자 ceil(팩/3))
- 이상 = 역마진 OR 마진율 < 채널 baseline − buffer.
파일 자동 인계: 파일처리에서 오픈마켓(송장출력)·천년경영(output) 실행 시 이 세션 인박스에 자동 적재 →
재업로드 불요. 수동 갱신 시 그 시점 파일로 override. 상품관리는 reference 라이브(항상 최신).
★ 조기 트립와이어. 정산 진실 = 매출자료 월정산(대시보드·마진침식 실판매). PII는 박스 그룹키로만·미저장.
"""
import csv
import io
import json
import sys
import unicodedata
import urllib.error
import urllib.request
from datetime import datetime, timezone, timedelta
from pathlib import Path

_KST = timezone(timedelta(hours=9))
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import pandas as pd
import streamlit as st

from core.intelligence import daily_margin as dm
from core.intelligence import daily_inbox as inbox
from core.intelligence import stockout_board as sb

_REF = Path(__file__).parent.parent.parent / "reference"
_APP_API = "https://api.github.com/repos/OURPROJECTDAO/work-automation-app/contents"


def _pat() -> str:
    return st.secrets.get("GITHUB_PAT", "")


def _data_secret():
    repo = "OURPROJECTDAO/work-automation-data"
    try:
        d = st.secrets["data"]
        return d["pat"], d.get("repo", repo)
    except Exception:
        return st.secrets.get("GITHUB_PAT", ""), repo


def _gh_raw(path: str):
    req = urllib.request.Request(
        f"{_APP_API}/{path}",
        headers={"Authorization": f"Bearer {_pat()}", "Accept": "application/vnd.github.raw"},
    )
    try:
        with urllib.request.urlopen(req) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, b""


def _nfc(v):
    return unicodedata.normalize("NFC", str(v)).strip() if pd.notna(v) else ""


@st.cache_data(ttl=600, show_spinner=False)
def _master_lookup():
    """관리코드(NFC) → (박스내품, master매입단가, 상품명)."""
    code, text = _gh_raw("reference/product_master.csv")
    if code != 200:
        return {}, {}, {}
    df = pd.read_csv(io.BytesIO(text), dtype=str)
    df["k"] = df["관리코드"].map(_nfc)
    box = {k: float(v) for k, v in zip(df["k"], pd.to_numeric(df["박스내품"], errors="coerce").fillna(1.0))}
    price = {k: float(v) for k, v in zip(df["k"], pd.to_numeric(df["매입단가"], errors="coerce")) if pd.notna(v)}
    name = {k: n for k, n in zip(df["k"], df["상품명"].fillna(""))}
    return box, price, name


@st.cache_data(ttl=600, show_spinner=False)
def _box_stock_lookup():
    """관리코드(NFC) → 박스재고(product_master '박스' 컬럼). 품절 알림판 재입고 판정용."""
    code, text = _gh_raw("reference/product_master.csv")
    if code != 200:
        return {}
    df = pd.read_csv(io.BytesIO(text), dtype=str)
    return {_nfc(k): float(v) for k, v in
            zip(df["관리코드"], pd.to_numeric(df["박스"], errors="coerce").fillna(0.0))}


@st.cache_data(ttl=600, show_spinner=False)
def _pc_lookup():
    """상품코드(NFC) → 관리코드 — 송장 PC낱개 해소용."""
    code, text = _gh_raw("reference/product_master.csv")
    if code != 200:
        return {}
    df = pd.read_csv(io.BytesIO(text), dtype=str)
    return {_nfc(sc): _nfc(mg) for sc, mg in zip(df["상품코드"], df["관리코드"]) if _nfc(sc)}


@st.cache_data(ttl=600, show_spinner=False)
def _hapo_codes():
    """175~200ml 30개입 합포가능 관리코드 set — 송장 택배 합포 배분(시나리오2)."""
    code, text = _gh_raw("reference/hapo_175_190.csv")
    if code != 200:
        return set()
    lines = text.decode("utf-8-sig").splitlines()
    return {_nfc(row["관리코드"]) for row in csv.DictReader(lines) if _nfc(row.get("관리코드", ""))}


@st.cache_data(ttl=600, show_spinner=False)
def _baseline_dict():
    """{관리코드(NFC): {채널: 기준마진}}."""
    code, text = _gh_raw("reference/baseline_margin.csv")
    if code != 200:
        return {}
    lines = text.decode("utf-8-sig").splitlines()
    rd = csv.DictReader(lines)
    chans = [c for c in (rd.fieldnames or []) if c != "관리코드"]
    out = {}
    for row in csv.DictReader(lines):
        k = _nfc(row["관리코드"])
        out[k] = {c: float(row[c]) for c in chans if row[c] not in ("", "None")}
    return out


def _to_xlsx(df: pd.DataFrame, title: str) -> bytes:
    from openpyxl import Workbook

    def _cell(v):
        if pd.isna(v):
            return None
        if hasattr(v, "isoformat"):
            return v.isoformat()
        if hasattr(v, "item"):
            return v.item()
        return v

    wb = Workbook(); ws = wb.active; ws.title = title[:31]
    ws.append(list(df.columns))
    for _, r in df.iterrows():
        ws.append([_cell(v) for v in r])
    if "관리코드" in df.columns:
        col = df.columns.get_loc("관리코드") + 1
        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            for c in row:
                c.number_format = "@"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def _slot_ui(slot: str, label: str, key: str):
    """슬롯 1개: 자동 인계 상태 표시 + 수동 갱신 업로더(override). 유효 바이트 반환(없으면 None)."""
    cur = inbox.get(st.session_state, slot)
    c = st.columns([3, 2])
    if cur:
        c[0].success(f"🟢 {label} — 자동 인계됨: {cur['name']} ({cur['ts']})")
    else:
        c[0].info(f"⚪ {label} — 자동 인계 없음. 파일처리에서 실행했거나, 오른쪽에서 직접 올리세요.")
    up = c[1].file_uploader(f"{label} 수동 갱신", type=["xlsx"], key=key,
                            label_visibility="collapsed")
    if up is not None:
        data = up.getvalue()
        inbox.push(st.session_state, slot, data, up.name,
                   datetime.now(_KST).strftime("%m-%d %H:%M") + " 수동")
        return data
    return cur["bytes"] if cur else None


# ─────────────────────────────────────────────────────────
st.title("📅 데일리 대시보드")
st.caption("매일 반복 업무 산출물로 **오늘의 마진을 즉시 점검**합니다. "
           "파일처리에서 **오픈마켓(송장출력)·천년경영(output)**을 실행하면 이 세션에서 자동 인계돼 "
           "재업로드가 필요 없습니다. (상품관리는 reference 라이브 — 항상 최신)")

# ── 🚨 품절 알림판 (발주 품절목록 → 재입고 박스재고>0 자동해제 + 수동삭제) ──
st.subheader("🚨 품절 알림판")
_pat_d, _repo_d = _data_secret()
if not _pat_d:
    st.info("data repo 시크릿([data] pat)이 없어 품절 알림판을 쓸 수 없습니다.")
else:
    if st.button("🔄 상품관리 다시 읽기(재입고 반영)", key="sb_refresh"):
        st.cache_data.clear(); st.rerun()
    _box_stock = _box_stock_lookup()
    _today = datetime.now(_KST).strftime("%Y-%m-%d")
    _board = sb.read_board(_pat_d, _repo_d)
    if _board and _box_stock:   # 재입고 reconcile(박스재고>0) → 입고로그 + 자동삭제
        _board2, _restocked = sb.reconcile(_board, _box_stock, _today)
        if _restocked:
            sb.append_log(_pat_d, _repo_d, _restocked, f"restock +{len(_restocked)} ({_today})")
            sb.write_board(_pat_d, _repo_d, _board2, f"board: 재입고 {len(_restocked)}건 자동삭제 ({_today})")
            _board = _board2
            st.success("✅ 재입고 자동처리 " + str(len(_restocked)) + "건 — 입고로그 기록·알림판 제거: "
                       + ", ".join(f"{r['관리코드']}({r['품절일수']}일)" for r in _restocked))
    _bdf = sb.board_to_frame(_board, _box_stock, _today)
    if _bdf.empty:
        st.success("현재 품절(미입고) 상품이 없습니다. 발주 품절목록에 뜨면 여기 자동 등록됩니다.")
    else:
        st.caption(f"품절 {len(_bdf)}건 — 발주 품절목록 자동 등록 · 박스재고>0 들어오면 자동 입고처리. 🗑=수동 제거(로그 없음)")
        _h = st.columns([2, 5, 3, 2, 1])
        for _col, _t in zip(_h, ["관리코드", "상품명", "품절", "현재박스", ""]):
            _col.markdown(f"**{_t}**")
        for _r in _bdf.to_dict("records"):
            _cc = st.columns([2, 5, 3, 2, 1])
            _cc[0].write(_r["관리코드"])
            _cc[1].write(_r["상품명"])
            _since, _n = _r["품절시작일"], _r["N일째"]
            _cc[2].write(f"{pd.Timestamp(_since).strftime('%m월%d일')}부터 {_n}일째" if _since else "")
            _bs = _r["현재박스재고"]
            _cc[3].write("—" if _bs is None else f"{_bs:g}")
            if _cc[4].button("🗑", key=f"sb_rm_{_r['관리코드']}", help="수동 제거(로그 없음)"):
                sb.write_board(_pat_d, _repo_d, sb.manual_remove(_board, _r["관리코드"]),
                               f"board: 수동 제거 {_r['관리코드']} ({_today})")
                st.rerun()
    with st.expander("📥 최근 입고 로그"):
        _log = sb.read_log(_pat_d, _repo_d)
        if _log.empty:
            st.caption("입고 로그가 아직 없습니다.")
        else:
            st.dataframe(_log.tail(50).iloc[::-1], hide_index=True, use_container_width=True)

st.divider()
st.subheader("📊 당일 마진 점검")
box_lookup, master_price, name_lookup = _master_lookup()

st.subheader("오늘 파일")
b_chun = _slot_ui(inbox.SLOT_CHEONNYEON, "천년경영 output", "daily_chun")
b_inv = _slot_ui(inbox.SLOT_INVOICE, "송장출력", "daily_inv")
st.caption("🟢 상품관리(master) — reference 라이브(연동데이터관리에서 매일 갱신)." if box_lookup
           else "⚠️ 상품관리(master)를 불러올 수 없습니다.")

_DCH = sorted(set(dm.SHEET_TO_CMM.values()))
with st.expander("⚙️ 설정 (안전마진 여유 · 채널 flat 택배단가)", expanded=False):
    buffer = st.number_input("기준마진 미달 여유(%p)", 0, 10, 2, 1,
                             help="마진율이 채널 기준마진보다 이만큼↓이면 '미달' 표시") / 100.0
    _fc = st.columns(4)
    flat_by_channel = {ch: _fc[i % 4].number_input(ch, 0, 10000, dm.DEFAULT_FLAT, 100, key=f"d_flat_{ch}")
                       for i, ch in enumerate(_DCH)}

if not (b_chun and b_inv):
    _miss = [n for n, b in [("천년경영 output", b_chun), ("송장출력", b_inv)] if not b]
    st.info(f"아직 준비 안 된 파일: **{', '.join(_miss)}**. 파일처리에서 실행하거나 위에서 직접 올리면 "
            "바로 당일 마진을 계산합니다.")
elif not box_lookup:
    st.warning("product_master를 불러올 수 없습니다.")
else:
    try:
        _alloc, _chb = dm.parse_invoice_shipping(b_inv, box_lookup, _pc_lookup(), _hapo_codes())
        _sdf = dm.parse_cheonnyeon_sales(b_chun, box_lookup)
    except Exception as e:
        st.error(f"파싱 오류: {e}")
        _alloc, _chb, _sdf = {}, {}, pd.DataFrame()
    if _sdf.empty:
        st.warning("천년경영 output에서 대상 채널 데이터를 찾지 못했습니다 (시트명·형식 확인).")
    else:
        ddf = dm.compute_daily_margin(_sdf, _alloc, master_price, name_lookup,
                                      _baseline_dict(), flat_by_channel=flat_by_channel,
                                      buffer=float(buffer))
        anom = ddf[ddf["역마진"] | ddf["미달"]].reset_index(drop=True)
        _rev, _mar = ddf["매출"].sum(), ddf["마진"].sum()
        c = st.columns(4)
        c[0].metric("당일 매출(net)", f"{_rev:,.0f} 원")
        c[1].metric("당일 마진", f"{_mar:,.0f} 원", f"{100 * _mar / _rev:.1f}%" if _rev else None)
        c[2].metric("이상 건", f"{len(anom)} 건")
        c[3].metric("역마진", f"{int(ddf['역마진'].sum())} 건")
        _bx = " · ".join(f"{ch} {n}" for ch, n in sorted(_chb.items(), key=lambda x: -x[1]))
        st.caption(f"실제 박스(택배)수 — {_bx} · 합계 {sum(_chb.values())} · "
                   f"택배=물리 박스 배분(250/355 자동합포 + 175~200ml 30개입 수령자 ceil(팩/3))")
        _view = st.radio("보기", ["이상치만", "전체"], horizontal=True, key="d_view")
        _show = anom if _view == "이상치만" else ddf
        if _show.empty:
            st.success("✅ 당일 역마진·기준 미달 상품이 없습니다.")
        else:
            _d = _show.copy()
            for col in ("마진율", "기준마진"):
                _d[col] = (_d[col].astype(float) * 100).round(1)
            st.dataframe(_d, hide_index=True, use_container_width=True, height=460,
                         column_config={
                             "매출": st.column_config.NumberColumn("매출(net)", format="%d"),
                             "낱개수량": st.column_config.NumberColumn("낱개", format="%d"),
                             "박스": st.column_config.NumberColumn("박스(송장배분)", format="%.1f"),
                             "원가": st.column_config.NumberColumn(format="%d"),
                             "택배": st.column_config.NumberColumn(format="%d"),
                             "마진": st.column_config.NumberColumn(format="%d"),
                             "마진율": st.column_config.NumberColumn("마진%", format="%.1f"),
                             "기준마진": st.column_config.NumberColumn("기준%", format="%.1f"),
                             "역마진": st.column_config.CheckboxColumn("역마진"),
                             "미달": st.column_config.CheckboxColumn("미달"),
                         })
            st.download_button("📥 XLSX", _to_xlsx(anom, "당일마진이상"),
                               "당일마진_이상.xlsx", key="d_dl")
