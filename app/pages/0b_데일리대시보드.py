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
from core.intelligence import purchases as _buy
from core.intelligence import stock_history as shh
from core.workflows import channel_margin_monitor as cmm
from core.workflows import upload_monitor as um

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

@st.cache_data(ttl=1800, show_spinner="매입현황(최근입고·평균주기) 불러오는 중...")
def _buyin_cadence():
    """관리코드(NFC) → {최근입고일·평균주기·입고횟수} (발주일 기준 1년·최근 13개월 파티션). 품절 알림판 표시용."""
    pat, repo = _data_secret()
    if not pat:
        return {}
    try:
        mons = _buy.list_months(pat, repo)[-13:]
        parts = [_buy.read_partition(pat, repo, m) for m in mons]
        b = pd.concat(parts, ignore_index=True) if parts else None
        return _buy.cadence_by_code(b, months=12, now=datetime.now(_KST).replace(tzinfo=None))
    except Exception:
        return {}


@st.cache_data(ttl=1800, show_spinner="가격 변동 감지 중...")
def _price_changes(days: int, threshold: float):
    """최근 days일 내 ±threshold 이상 가격 변동(매입가·판매가). 1b 스냅샷 연속 비교."""
    pat, repo = _data_secret()
    if not pat:
        return pd.DataFrame()
    try:
        mons = shh.list_snapshot_months(pat, repo)[-3:]
        parts = [shh.read_snapshots(pat, repo, m) for m in mons]
        snaps = pd.concat([p for p in parts if len(p)], ignore_index=True) if parts else pd.DataFrame()
        chg = shh.detect_price_changes(snaps, threshold=threshold)
        if chg.empty:
            return chg
        cutoff = pd.Timestamp(datetime.now(_KST).date()) - pd.Timedelta(days=days)
        return chg[pd.to_datetime(chg["금일"]) >= cutoff].reset_index(drop=True)
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=1800, show_spinner="신규 입고·미업로드 감지 중...")
def _new_uploads(days: int):
    """최근 days일 재고 새로 생긴(입고·신규등재) 상품코드 ∩ 8채널 전부 미업로드 ∩ 재고>0.

    재고 신호 = stock_history.detect_new_stock(1b 스냅샷). 채널 미등록 판정·제외/skip은
    upload_monitor 코어 재사용(중복 0 — 업로드감시의 '최근 입고' 서브셋). 키=상품코드.
    반환 컬럼: 관리코드·상품명·박스재고·유형·이벤트일·_mc·올릴채널수.
    """
    pat, repo = _data_secret()
    if not pat:
        return pd.DataFrame()
    try:
        snaps = shh.read_all_snapshots(pat, repo)   # baseline floor용 전 기간(월경계 연속)
        since = pd.Timestamp(datetime.now(_KST).date()) - pd.Timedelta(days=days)
        ev = shh.detect_new_stock(snaps, since=since)
        if ev.empty:
            return pd.DataFrame()
        refs = um.load_references(str(_REF))
        rows = um.build_gap_table(str(_REF), refs)   # 비판매 제외·재고>0·채널별 skip 반영
        unup = {}                                    # 전채널 미업로드(모두 업로드필요/제외 & ≥1 업로드필요)
        for r in rows:
            chans = [r[k] for k in um.CHANNEL_KEYS]
            if all(c in (um.ST_NEED_UP, um.ST_SKIP_CH) for c in chans) and um.ST_NEED_UP in chans:
                unup[r["상품코드"]] = (r, sum(1 for c in chans if c == um.ST_NEED_UP))
        out = []
        for _, e in ev.iterrows():
            sc = e["상품코드"]
            if sc not in unup:
                continue
            r, need = unup[sc]
            mc = _nfc(r.get("관리코드"))
            out.append({
                "관리코드": mc or f"({sc})",
                "상품명": _nfc(r.get("상품명")),
                "박스재고": r.get("박스재고"),
                "유형": e["유형"],
                "이벤트일": pd.Timestamp(e["이벤트일"]).strftime("%m-%d"),
                "_mc": mc,
                "올릴채널수": need,
            })
        if not out:
            return pd.DataFrame()
        return (pd.DataFrame(out)
                .sort_values(["이벤트일", "박스재고"], ascending=[False, False])
                .reset_index(drop=True))
    except Exception:
        return pd.DataFrame()


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


@st.cache_data(ttl=600, show_spinner=False)
def _cmm_baseline_text():
    code, text = _gh_raw("reference/baseline_margin.csv")
    return text.decode("utf-8") if code == 200 else ""


def _cmm_key(channel):
    """데일리 채널명 → CHANNEL_CONFIG 키 보정. ESM은 cfg 키가 'esm'(소문자)라
    데일리 SHEET_TO_CMM 'ESM'(대문자)과 불일치 → 대소문자 무시 매칭으로 해소(전 채널 방어)."""
    if channel in cmm.CHANNEL_CONFIG:
        return channel
    low = str(channel).lower()
    for k in cmm.CHANNEL_CONFIG:
        if k.lower() == low:
            return k
    return channel


@st.cache_data(ttl=600, show_spinner="채널 권장가 불러오는 중...")
def _cmm_listing(channel: str):
    """채널 listing(reference/listing_<key>.csv) → (recs, compute_listing rows). 없으면 None."""
    ck = _cmm_key(channel)
    cfg = cmm.CHANNEL_CONFIG.get(ck)
    if not cfg:
        return None
    code, text = _gh_raw(f"reference/listing_{cfg['key']}.csv")
    if code != 200 or not text:
        return None
    recs = cmm.csv_text_to_recs(text.decode("utf-8"))
    bl = _cmm_baseline_text()
    override = cmm.parse_baseline_dict(bl) if bl else None
    rows, _ = cmm.compute_listing(recs, ck, str(_REF), baseline_override=override)
    return recs, rows


def _rng(vals):
    vals = [int(v) for v in vals if v not in (None, "", 0)]
    if not vals:
        return None
    lo, hi = min(vals), max(vals)
    return f"{lo:,}" if lo == hi else f"{lo:,}~{hi:,}"


@st.cache_data(ttl=600, show_spinner=False)
def _cmm_refs():
    return cmm.load_references(str(_REF))


def _reco_from_master(channel, code, refs, bdict):
    """product_master 매입가 기준 권장가(채널 기준마진 달성가) — listing 불요. N=1·채널기본 배송비 추정.
    cmm 권장가 공식과 동일: ((매입가+2700)/(1-기준마진) - 배송비*0.967)/(1-수수료), 100원 올림."""
    cfg = cmm.CHANNEL_CONFIG.get(_cmm_key(channel)) or {}
    _typ, base, *_ = cmm.resolve_code(code, refs)
    if base is None:
        return None
    bm = (bdict.get(code) or {}).get(cfg.get("baseline_col"))
    b = cmm._num(bm, None) if bm not in (None, "") else None
    if b is None or b >= 1:
        return None
    comm = cfg.get("commission")
    if comm is None:                                  # 배민 등 상품별 수수료 → listing 없으면 대략치
        comm = cfg.get("commission_add", 0.0) + 0.045
    매입가 = base * 1.0                                # N 미상 → 1 가정(listing 있으면 실 N 사용)
    배송비 = cfg.get("ship_fee_const", 0) or 0
    reco = ((매입가 + cfg.get("real_ship", 2700)) / (1 - b)
            - 배송비 * cfg.get("ship_settle", 0.967)) / (1 - comm)
    return cmm._ceil100(reco)


def _reco_lookup(pairs, buffer=0.0):
    """{(채널,관리코드NFC):(권장가표시, 현재가표시, listing마진표시, listing미달여부)}.
    권장가 = listing에 있으면 정확값(실 N·실 배송비), 없으면 product_master 매입가 기준 추정(항상 표시).
    현재가 = 채널 저장 listing 기준(미등재면 None — 최신성 인정).
    listing미달(이중검수) = 그 채널 listing 마진율 < 기준마진 − buffer(당일과 동일 규칙):
      True=listing도 미달(구조적·가격조정) / False=listing 정상(당일 미달은 일시적) / None=listing 없음."""
    refs = _cmm_refs()
    bl = _cmm_baseline_text()
    bdict = cmm.parse_baseline_dict(bl) if bl else {}
    channels = {ch for ch, _ in pairs}
    lmap = {}
    for ch in channels:
        try:
            data = _cmm_listing(ch)
        except Exception:
            data = None
        m = {}
        if data:
            _, rows = data
            for r in rows:
                mc = _nfc(r.get("관리코드"))
                if not mc:
                    continue
                a = m.setdefault(mc, {"reco": [], "cur": [], "mar": [], "under": False})
                if r.get("권장가") is not None:
                    a["reco"].append(r["권장가"])
                if r.get("판매가") is not None:
                    a["cur"].append(cmm._num(r.get("판매가")))
                mr, bm = r.get("마진율"), r.get("기준마진율")
                if mr is not None:
                    a["mar"].append(mr)
                    if bm is not None and mr < bm - buffer:
                        a["under"] = True
        lmap[ch] = m
    out = {}
    for ch, code in pairs:
        mc = _nfc(code)
        a = lmap.get(ch, {}).get(mc)
        reco = _rng(a["reco"]) if (a and a["reco"]) else None
        cur = _rng(a["cur"]) if (a and a["cur"]) else None
        if a and a["mar"]:
            ms = [round(x * 100, 1) for x in a["mar"]]
            lmar = f"{ms[0]:.1f}%" if min(ms) == max(ms) else f"{min(ms):.1f}~{max(ms):.1f}%"
            lunder = a["under"]
        else:
            lmar, lunder = None, None
        if reco is None:                              # listing 권장가 없음 → product_master 기준 역산
            try:
                est = _reco_from_master(ch, mc, refs, bdict)
            except Exception:
                est = None
            reco = f"{est:,}" if est else None
        out[(ch, mc)] = (reco, cur, lmar, lunder)
    return out


def _supports_price_change(cfg) -> bool:
    """가격변경 시트 생성 가능 채널? price_form(append/filter) 또는 스마트스토어형 bulk(즉시할인 cols·consolidate 아님). 알리=불가."""
    if cfg.get("price_form"):
        return True
    cols = cfg.get("cols") or {}
    return ("즉시할인" in cols) and not cfg.get("consolidate")


def _gen_price_form(channel, cfg, pf, recs, rows, pids):
    """선택 채널·pids → 가격변경 시트 bytes (cmm 빌더 재사용). 반환 dict(channel/bytes/preview 또는 error)."""
    try:
        mode = (pf or {}).get("mode")
        prev = []
        if mode == "append":
            items, prev, _sk = cmm.build_append_items(pf, rows, recs, pids)
            if not items:
                return {"channel": channel, "error": "권장가 산출 가능 항목이 없습니다(미매칭/기준 미설정)."}
            out = cmm.build_price_form_append((_REF / pf["template"]).read_bytes(), items, pf)
        elif mode == "filter":
            rc, raw = _gh_raw(f"reference/listing_{cfg['key']}.xlsx")
            if rc != 200 or not raw:
                return {"channel": channel, "error": f"{channel} 원본양식(.xlsx)이 없습니다. 채널마진모니터에서 '상품관리 갱신 → 전체 교체'를 1회 실행하세요."}
            out, prev, _sk, _ms = cmm.build_filter_price_xlsx(raw, rows, pids, cfg)
            if not prev:
                return {"channel": channel, "error": "권장가 산출 가능 항목이 없습니다(미매칭/기준 미설정)."}
        else:  # 스마트스토어 bulk(원본 filter)
            new_prices, _sk = cmm.compute_new_prices(rows, recs, set(pids))
            if not new_prices:
                return {"channel": channel, "error": "권장가 산출 가능 항목이 없습니다(미매칭/기준 미설정)."}
            rc, raw = _gh_raw(f"reference/listing_{cfg['key']}.xlsx")
            if rc != 200 or not raw:
                return {"channel": channel, "error": f"{channel} 원본양식(.xlsx)이 없습니다. 채널마진모니터에서 '상품관리 갱신 → 전체 교체'를 1회 실행하세요."}
            out, _kept, _ms = cmm.build_bulk_price_xlsx(raw, new_prices, cfg)
            rb = {r["상품번호"]: r for r in recs}
            ro = {r["상품번호"]: r for r in rows}
            prev = [{"상품명": ro[p]["상품명"], "현재판매가": int(rb[p]["판매가"]),
                     "새판매가": v[0], "권장가": ro[p]["권장가"]} for p, v in new_prices.items()]
        return {"channel": channel, "bytes": out, "preview": prev,
                "name": f"{channel}_가격변경_{datetime.now(_KST):%Y%m%d}.xlsx"}
    except Exception as e:
        return {"channel": channel, "error": f"생성 오류: {e}"}


def _do_price_change(channel, codes):
    """단일 채널 + 선택 관리코드 set → 그 채널 가격변경 시트 생성/다운로드 UI."""
    cfg = cmm.CHANNEL_CONFIG.get(_cmm_key(channel)) or {}
    pf = cfg.get("price_form")
    if not _supports_price_change(cfg):
        st.info(f"**{channel}**는 가격변경 양식이 아직 없습니다(예: 알리). "
                "지원: 스마트스토어·식봄·캐시노트·배민상회·쿠팡·올웨이즈·ESM.")
        return
    data = _cmm_listing(channel)
    if not data:
        st.warning(f"**{channel}** 저장 listing이 없습니다. 채널마진모니터에서 '상품관리 갱신'을 1회 실행하세요.")
        return
    recs, rows = data
    pids = [r["상품번호"] for r in rows if _nfc(r.get("관리코드")) in codes]
    if not pids:
        st.warning("선택 상품이 해당 채널 listing에 없습니다(listing 갱신 필요할 수 있음).")
        return
    if st.button(f"🛠️ {channel} 가격변경 시트 생성 — 관리코드 {len(codes)}개 → listing {len(pids)}건",
                 type="primary", key="d_pc_gen"):
        st.session_state["d_pcform"] = _gen_price_form(channel, cfg, pf, recs, rows, pids)
    _form = st.session_state.get("d_pcform")
    if _form and _form.get("channel") == channel:
        if _form.get("error"):
            st.warning(_form["error"])
        else:
            st.download_button(f"⬇️ {channel} 가격변경 시트 다운로드 (.xlsx)",
                               _form["bytes"], _form["name"], type="primary", key="d_pc_dl")
            if _form.get("preview"):
                st.dataframe(pd.DataFrame(_form["preview"]), hide_index=True, use_container_width=True)


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
    _bdf = sb.board_to_frame(_board, _box_stock, _today, cadence=_buyin_cadence())
    if _bdf.empty:
        st.success("현재 품절(미입고) 상품이 없습니다. 발주 품절목록에 뜨면 여기 자동 등록됩니다.")
    else:
        st.caption(f"품절 {len(_bdf)}건 — 발주 품절목록 자동 등록 · 박스재고>0 들어오면 자동 입고처리. 최근입고·평균주기·입고횟수=최근 1년 매입현황. 🗑=수동 제거(로그 없음)")
        _w = [2, 4, 3, 1.4, 2, 1.4, 1.4, 1]
        _h = st.columns(_w)
        for _col, _t in zip(_h, ["관리코드", "상품명", "품절", "현재박스", "최근입고", "평균주기", "입고(1년)", ""]):
            _col.markdown(f"**{_t}**")
        for _r in _bdf.to_dict("records"):
            _cc = st.columns(_w)
            _cc[0].write(_r["관리코드"])
            _cc[1].write(_r["상품명"])
            _since, _n = _r["품절시작일"], _r["N일째"]
            _cc[2].write(f"{pd.Timestamp(_since).strftime('%m월%d일')}부터 {_n}일째" if _since else "")
            _bs = _r["현재박스재고"]
            _cc[3].write("—" if _bs is None else f"{_bs:g}")
            _cc[4].write(_r.get("최근입고일") or "—")
            _avg = _r.get("평균매입주기")
            _cc[5].write("—" if _avg is None else f"{_avg:g}일")
            _cnt = _r.get("입고횟수(1년)")
            _cc[6].write("—" if _cnt is None else f"{_cnt:g}회")
            if _cc[7].button("🗑", key=f"sb_rm_{_r['관리코드']}", help="수동 제거(로그 없음)"):
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
st.subheader("🆕 신규 업로드 대상")
st.caption("최근 재고가 **새로 들어왔는데(입고·신규등재) 아직 어느 채널에도 안 올라간** 상품 — 신규 업로드 후보. "
           "재고 스냅샷(2026-06-15~ 적립)으로 감지 · 채널 등록현황은 업로드감시와 동일 기준(업로드감시의 *최근 입고* 서브셋).")
if not _pat_d:
    st.info("data repo 시크릿([data] pat)이 없어 신규 업로드 대상을 쓸 수 없습니다.")
else:
    _u1, _u2 = st.columns([3, 1])
    _udays = _u1.slider("최근 며칠", 1, 30, 7, key="nu_days")
    _u2.write("")
    if _u2.button("🔄 다시 읽기", key="nu_refresh"):
        st.cache_data.clear(); st.rerun()
    _nu = _new_uploads(int(_udays))
    if _nu is None or _nu.empty:
        st.success(f"최근 {_udays}일 내 새로 들어온 미업로드 상품이 없습니다. "
                   "(재고 스냅샷은 상품관리 업로드일마다 적립 — 2026-06-15부터 누적)")
    else:
        _cad = _buyin_cadence()

        def _last_buy(mc):
            d = (_cad.get(mc) or {}).get("최근입고일")
            return pd.Timestamp(d).strftime("%Y-%m-%d") if d is not None else "—"

        _nu = _nu.copy()
        _nu["최근매입일"] = _nu["_mc"].map(_last_buy)
        st.caption(f"{len(_nu)}건 — 8채널 어디에도 미등록 · 현재 박스재고>0. "
                   "최근매입일=매입현황(월1회 적재라 당일 매입은 늦게 반영).")
        _ucols = ["관리코드", "상품명", "박스재고", "유형", "이벤트일", "최근매입일", "올릴채널수"]
        _sty = (_nu[_ucols].style
                .format({"박스재고": "{:,.0f}", "올릴채널수": "{:.0f}"}, na_rep="—"))
        st.dataframe(_sty, hide_index=True, use_container_width=True, height=320)
        st.download_button("📥 XLSX", _to_xlsx(_nu[_ucols], "신규업로드대상"),
                           "신규업로드대상.xlsx", key="nu_dl")
        st.caption("→ 채널별 상세 업로드 현황·등록 인계는 **업로드감시** 페이지에서.")

st.divider()
st.subheader("💲 가격 변동 알림")
if not _pat_d:
    st.info("data repo 시크릿([data] pat)이 없어 가격 변동 알림을 쓸 수 없습니다.")
else:
    _p1, _p2, _p3 = st.columns([2, 2, 1])
    _days = _p1.slider("최근 며칠", 1, 30, 7, key="pc_days")
    _thr = _p2.slider("변동 임계값(%)", 1, 20, 2, key="pc_thr") / 100.0
    _p3.write("")
    if _p3.button("🔄 다시 읽기", key="pc_refresh"):
        st.cache_data.clear(); st.rerun()
    _chg = _price_changes(int(_days), float(_thr))
    if _chg is None or _chg.empty:
        st.success(f"최근 {_days}일 내 ±{_thr * 100:g}% 이상 가격 변동이 없습니다. "
                   "(스냅샷은 상품관리 업로드일마다 적립 — 2026-06-15부터 누적)")
    else:
        _, _, _nm = _master_lookup()
        _chg = _chg.copy()
        _chg["상품명"] = _chg["관리코드"].map(lambda x: _nm.get(_nfc(x), ""))
        _kinds = st.multiselect("구분", ["판매가", "매입가"], default=["판매가", "매입가"],
                                key="pc_kind")
        _chg = _chg[_chg["구분"].isin(_kinds)].reset_index(drop=True)
        _up = int((_chg["방향"] == "인상").sum())
        _dn = int((_chg["방향"] == "인하").sum())
        _m = st.columns(3)
        _m[0].metric("변동 상품", f"{len(_chg)} 건")
        _m[1].metric("🔺 인상", f"{_up} 건")
        _m[2].metric("🔻 인하", f"{_dn} 건")
        if _chg.empty:
            st.caption("선택한 구분에 해당하는 변동이 없습니다.")
        else:
            _bs = _box_stock_lookup()
            _chg["박스재고"] = _chg["관리코드"].map(lambda x: _bs.get(_nfc(x)))
            _disp = _chg.copy()
            _disp["방향"] = _disp["방향"].map({"인상": "▲ 인상", "인하": "▼ 인하"})
            _disp["변동일"] = pd.to_datetime(_disp["금일"]).dt.strftime("%m-%d")
            _disp = _disp.rename(columns={"변동률": "변동률(%)"})
            _disp = _disp[["관리코드", "상품명", "박스재고", "구분", "방향",
                           "전일가", "금일가", "변동률(%)", "변동일"]]

            def _dir_color(v):
                s = str(v)
                if "인상" in s:
                    return "color:#d11; font-weight:600"
                if "인하" in s:
                    return "color:#1565c0; font-weight:600"
                return ""

            _sty = (_disp.style
                    .map(_dir_color, subset=["방향"])
                    .format({"전일가": "{:,.0f}", "금일가": "{:,.0f}",
                             "변동률(%)": "{:.2f}", "박스재고": "{:,.0f}"}, na_rep="—"))
            st.dataframe(_sty, hide_index=True, use_container_width=True, height=360)
            st.download_button("📥 XLSX", _to_xlsx(_chg, "가격변동"),
                               "가격변동알림.xlsx", key="pc_dl")

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
        _g = (ddf.groupby("채널", as_index=False)
              .agg(매출=("매출", "sum"), 원가=("원가", "sum"), 택배=("택배", "sum"),
                   마진=("마진", "sum"), 품목수=("관리코드", "nunique")))
        _g["마진율"] = (_g["마진"] / _g["매출"].where(_g["매출"] > 0) * 100).round(1)
        _g = _g.sort_values("매출", ascending=False).reset_index(drop=True)
        st.markdown("**채널별 요약 (당일 매출·마진율)**")
        st.dataframe(_g[["채널", "매출", "원가", "택배", "마진", "마진율", "품목수"]],
                     hide_index=True, use_container_width=True,
                     column_config={
                         "매출": st.column_config.NumberColumn("매출(net)", format="%d"),
                         "원가": st.column_config.NumberColumn(format="%d"),
                         "택배": st.column_config.NumberColumn(format="%d"),
                         "마진": st.column_config.NumberColumn(format="%d"),
                         "마진율": st.column_config.NumberColumn("마진%", format="%.1f"),
                         "품목수": st.column_config.NumberColumn(format="%d"),
                     })
        _view = st.radio("보기", ["이상치만", "전체"], horizontal=True, key="d_view")
        _show = (anom if _view == "이상치만" else ddf).reset_index(drop=True)
        if _show.empty:
            st.success("✅ 당일 역마진·기준 미달 상품이 없습니다.")
        else:
            _lk = _reco_lookup({(ch, _nfc(mc)) for ch, mc in zip(_show["채널"], _show["관리코드"])}, buffer=float(buffer))
            _disp = _show.copy()
            _NA4 = (None, None, None, None)
            _disp["현재가"] = [_lk.get((ch, _nfc(mc)), _NA4)[1]
                            for ch, mc in zip(_show["채널"], _show["관리코드"])]
            _disp["권장가"] = [_lk.get((ch, _nfc(mc)), _NA4)[0]
                            for ch, mc in zip(_show["채널"], _show["관리코드"])]
            _disp["listing마진"] = [_lk.get((ch, _nfc(mc)), _NA4)[2]
                                  for ch, mc in zip(_show["채널"], _show["관리코드"])]
            _verdict = {True: "⚠️ listing도 미달", False: "listing 정상(일시적)", None: "listing 없음"}
            _disp["판정"] = [_verdict.get(_lk.get((ch, _nfc(mc)), _NA4)[3], "listing 없음")
                           for ch, mc in zip(_show["채널"], _show["관리코드"])]
            for col in ("마진율", "기준마진"):
                _disp[col] = (_disp[col].astype(float) * 100).round(1)
            _order = ["채널", "관리코드", "상품명", "매출", "낱개수량", "박스", "원가", "택배",
                      "마진", "마진율", "기준마진", "현재가", "권장가", "listing마진", "판정",
                      "역마진", "미달"]
            st.caption("왼쪽 체크박스로 상품 선택 → 아래에서 **그 채널** 가격변경 시트 다운로드. "
                       "권장가 = 채널 기준마진율 달성 판매가(매입가 기준 역산 — 항상 표시). "
                       "현재가 = 채널 저장 listing 기준(미등재면 빈칸). "
                       "**판정(이중검수)**: 당일 미달이라도 listing 마진이 기준 이상이면 '일시적'(쿠폰·실박스 택배 등 — 가격 손댈 필요 없음), "
                       "listing도 미달이면 '구조적'(가격 조정 검토).")
            _ev = st.dataframe(_disp[_order], hide_index=True, use_container_width=True, height=460,
                               on_select="rerun", selection_mode="multi-row", key="d_table",
                               column_config={
                                   "매출": st.column_config.NumberColumn("매출(net)", format="%d"),
                                   "낱개수량": st.column_config.NumberColumn("낱개", format="%d"),
                                   "박스": st.column_config.NumberColumn("박스(송장배분)", format="%.1f"),
                                   "원가": st.column_config.NumberColumn(format="%d"),
                                   "택배": st.column_config.NumberColumn(format="%d"),
                                   "마진": st.column_config.NumberColumn(format="%d"),
                                   "마진율": st.column_config.NumberColumn("마진%", format="%.1f"),
                                   "기준마진": st.column_config.NumberColumn("기준%", format="%.1f"),
                                   "현재가": st.column_config.TextColumn("현재가"),
                                   "권장가": st.column_config.TextColumn("권장가(채널기준)"),
                                   "listing마진": st.column_config.TextColumn("listing마진"),
                                   "판정": st.column_config.TextColumn(
                                       "판정", help="당일 미달 + listing도 미달 = 구조적(가격조정). "
                                       "당일만 미달 = 일시적(쿠폰·택배 등, 가격 OK)."),
                                   "역마진": st.column_config.CheckboxColumn("역마진"),
                                   "미달": st.column_config.CheckboxColumn("미달"),
                               })
            st.download_button("📥 XLSX (전체 이상치)", _to_xlsx(anom, "당일마진이상"),
                               "당일마진_이상.xlsx", key="d_dl")
            try:
                _sel = list(_ev.selection["rows"])
            except Exception:
                _sel = []
            if _sel:
                _selrows = _show.iloc[_sel]
                _selch = sorted(set(_selrows["채널"]))
                st.markdown("---")
                st.markdown("#### 🧾 선택 상품 → 채널 가격변경 시트")
                if len(_selch) > 1:
                    st.warning("⚠️ 가격변경 시트는 **한 채널만** 가능합니다. 한 채널 상품만 선택하세요. "
                               f"(현재 선택 채널: {', '.join(_selch)})")
                else:
                    _do_price_change(_selch[0], {_nfc(x) for x in _selrows["관리코드"]})
