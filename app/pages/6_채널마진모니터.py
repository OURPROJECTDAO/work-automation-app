"""채널 가격·마진 모니터 (channel-margin-monitor).

저장된 상품관리(listing) 스냅샷을 자동 로드해 상품별 마진율·기준마진 대비 탐지·권장가(또는
제한 텍스트)·재고를 계산. 표에서 상품을 선택해 CSV 또는 '가격 일괄변경 양식'으로 내보낸다.
공식·근거 = KB workflows/channel-margin-monitor.md.
"""
import base64
import json
import sys
import urllib.request
from datetime import datetime, timedelta, timezone
from io import StringIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # repo root

import pandas as pd
import streamlit as st

from core.workflows import channel_margin_monitor as cmm

_REF = Path(__file__).parent.parent.parent / "reference"
_APP_REPO = "OURPROJECTDAO/work-automation-app"
_KST = timezone(timedelta(hours=9))
_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.title("💹 채널 가격·마진 모니터")
st.caption(
    "저장된 상품관리 기준으로 마진율·기준마진 대비 이탈·권장가를 계산합니다. "
    "매입가/재고는 상품관리(product_master), 기준마진율은 baseline_margin 기준. "
    "표에서 선택한 상품을 CSV 또는 가격 일괄변경 양식으로 내보낼 수 있습니다."
)


def _pat() -> str:
    return st.secrets.get("GITHUB_PAT", "")


def _gh(path, method="GET", data=None, raw=False):
    url = f"https://api.github.com/repos/{_APP_REPO}/contents/{path}"
    req = urllib.request.Request(url, method=method)
    if _pat():
        req.add_header("Authorization", f"Bearer {_pat()}")
    req.add_header("Accept", "application/vnd.github.raw" if raw else "application/vnd.github+json")
    if data is not None:
        req.data = json.dumps(data).encode()
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req) as r:
            return r.status, (r.read().decode() if raw else json.load(r))
    except urllib.error.HTTPError as e:
        return e.code, (e.read().decode() if raw else json.loads(e.read().decode() or "{}"))


def _gh_bytes(path):
    """바이너리(원본 .xlsx) raw 다운로드 → (status, bytes)."""
    url = f"https://api.github.com/repos/{_APP_REPO}/contents/{path}"
    req = urllib.request.Request(url)
    if _pat():
        req.add_header("Authorization", f"Bearer {_pat()}")
    req.add_header("Accept", "application/vnd.github.raw")
    try:
        with urllib.request.urlopen(req) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, b""


def _listing_path(key): return f"reference/listing_{key}.csv"
def _meta_path(key): return f"reference/listing_{key}.meta.json"
def _raw_path(key): return f"reference/listing_{key}.xlsx"


@st.cache_data(ttl=600, show_spinner="저장된 상품관리 불러오는 중...")
def _load_listing(key: str):
    code, text = _gh(_listing_path(key), raw=True)
    if code != 200:
        return None, {}
    recs = cmm.csv_text_to_recs(text)
    mcode, mtext = _gh(_meta_path(key), raw=True)
    meta = json.loads(mtext) if mcode == 200 else {}
    return recs, meta


def _commit_listing(key: str, recs: list) -> dict:
    meta = {"updated_at": datetime.now(_KST).isoformat(timespec="seconds"), "rows": len(recs)}
    for path, body in [(_listing_path(key), cmm.recs_to_csv(recs)),
                       (_meta_path(key), json.dumps(meta, ensure_ascii=False))]:
        code, m = _gh(path)
        payload = {"message": f"data(listing): {key} 갱신 ({len(recs)}건)",
                   "content": base64.b64encode(body.encode("utf-8")).decode()}
        if code == 200:
            payload["sha"] = m["sha"]
        _gh(path, "PUT", payload)
    return meta


def _commit_raw(key: str, xlsx_bytes: bytes):
    """원본 일괄변경 양식(.xlsx) 저장 — 가격변경 양식 출력의 원천(전체 컬럼 보존)."""
    path = _raw_path(key)
    code, m = _gh(path)
    payload = {"message": f"data(listing-raw): {key} 원본양식 갱신",
               "content": base64.b64encode(xlsx_bytes).decode()}
    if code == 200:
        payload["sha"] = m["sha"]
    _gh(path, "PUT", payload)


channel = st.selectbox("채널", list(cmm.CHANNEL_CONFIG.keys()), index=0)
cfg = cmm.CHANNEL_CONFIG[channel]
key = cfg["key"]
st.caption(
    f"수수료 {cfg['commission']*100:.0f}% · 배송비 정산 ×{cfg['ship_settle']} · "
    f"실택배비 {cfg['real_ship']:,}원 · 기준마진 '{cfg['baseline_col']}'"
    + (" · 마진제한 적용" if cfg.get("apply_floor") else "")
)

# ── 상품관리 갱신 ─────────────────────────────────────────────────────────────
committed = None
flash = None
with st.expander("📥 상품관리 갱신 (새 다운로드 업로드)"):
    up = st.file_uploader(f"{channel} 상품관리 다운로드 (.xlsx 전체 업로드)", type=["xlsx"], key=f"up_{key}")
    if up is not None:
        up_bytes = up.getvalue()
        new_recs = cmm.parse_download(up_bytes, cfg)
        st.write(f"업로드 파싱: **{len(new_recs):,}건**")
        if not _pat():
            st.error("저장용 PAT(st.secrets GITHUB_PAT)가 없어 커밋할 수 없습니다.")
        else:
            b1, b2 = st.columns(2)
            if b1.button("전체 교체 저장", type="primary", use_container_width=True,
                         help="최신 전체 다운로드로 덮어쓰기 (신규+가격변동 반영) + 원본양식 저장"):
                meta = _commit_listing(key, new_recs)
                _commit_raw(key, up_bytes)          # 원본 양식(전체 컬럼) 저장
                _load_listing.clear()
                committed = new_recs
                flash = f"전체 교체 완료 — {meta['rows']:,}건 ({meta['updated_at']})"
            if b2.button("신규만 추가", use_container_width=True,
                         help="기존 유지 + 새 상품번호만 병합 (원본양식에도 신규 행 추가)"):
                cur, _ = _load_listing(key)
                merged, added = cmm.merge_listing(cur or [], new_recs)
                meta = _commit_listing(key, merged)
                added_pids = {r["상품번호"] for r in new_recs} - {r["상품번호"] for r in (cur or [])}
                rcode, rawb = _gh_bytes(_raw_path(key))
                newraw = cmm.append_rows_to_raw(rawb, up_bytes, added_pids, cfg) if (rcode == 200 and rawb) else up_bytes
                _commit_raw(key, newraw)
                _load_listing.clear()
                committed = merged
                flash = f"신규 {added:,}건 추가 — 총 {meta['rows']:,}건"

if committed is not None:
    recs = committed
    meta = {"updated_at": datetime.now(_KST).isoformat(timespec="seconds"), "rows": len(committed)}
    st.success(flash)
else:
    recs, meta = _load_listing(key)

if not recs:
    st.info("저장된 상품관리가 없습니다. 위 '📥 상품관리 갱신'에서 다운로드를 올려 저장해 주세요.")
    st.stop()

st.caption(f"📦 저장된 상품관리 기준 · 최종 갱신 **{meta.get('updated_at', '?')}** · {len(recs):,}건")

rows, stats = cmm.compute_listing(recs, channel, str(_REF))

# ── KPI ──────────────────────────────────────────────────────────────────────
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("총 상품", f"{stats['총건수']:,}")
c2.metric("평균 마진율", f"{stats['평균마진율']*100:.2f}%" if stats["평균마진율"] is not None else "—")
c3.metric("마진 미달", f"{stats['마진미달']:,}", help="기준마진율보다 1%p 이상 낮음")
c4.metric("제한 상품", f"{stats['제한상품']:,}")
c5.metric("기준 미설정", f"{stats['미설정']:,}")
c6.metric("미매칭", f"{stats['미매칭']:,}")

# ── 필터 ──────────────────────────────────────────────────────────────────────
df = pd.DataFrame(rows)


def _rec_disp(r):
    if r["제한"]:
        return str(r["제한"])
    if r["매입가"] is None:
        return str(r["비고"]) or "미매칭"
    if r["기준마진율"] is None:
        return "기준 미설정"
    return f"{int(r['권장가']):,}" if pd.notna(r["권장가"]) else ""


df["권장가/제한"] = df.apply(_rec_disp, axis=1)

types = sorted(df["코드유형"].unique().tolist())
fc1, fc2 = st.columns([2, 3])
with fc1:
    pick = st.multiselect("코드유형", types, default=types)
with fc2:
    f1, f2, f3, f4 = st.columns(4)
    only_under = f1.checkbox("마진미달만", help="기준마진율보다 1%p 이상 낮은 상품")
    only_zero = f2.checkbox("재고 0")
    only_floor = f3.checkbox("제한상품만")
    only_miss = f4.checkbox("미매칭만")

view = df[df["코드유형"].isin(pick)].copy()
if only_under:
    view = view[view["탐지"].notna() & (view["탐지"] < cmm.MARGIN_UNDER_THRESHOLD)]
if only_zero:
    view = view[view["재고"].fillna(-1) == 0]
if only_floor:
    view = view[view["제한"].astype(str) != ""]
if only_miss:
    view = view[view["매입가"].isna()]

DISPLAY = ["상품번호", "관리코드", "상품명", "규격", "코드유형", "N", "재고",
           "매입가", "판매가", "배송비", "정산액", "마진율", "기준마진율", "탐지", "권장가/제한", "비고"]

# ── 선택(체크박스) — 필터 넘어 상품번호 기준으로 유지 ─────────────────────────
sel = st.session_state.setdefault(f"cmm_sel_{key}", set())
view_pids = set(view["상품번호"])
nonce = st.session_state.setdefault(f"cmm_nonce_{key}", 0)

s1, s2, s3 = st.columns([1, 1, 4])
if s1.button("전체 선택", use_container_width=True, help="현재 필터에 보이는 상품 전체 선택"):
    sel |= view_pids
    st.session_state[f"cmm_nonce_{key}"] = nonce + 1
    st.rerun()
if s2.button("전체 해제", use_container_width=True, help="현재 필터에 보이는 상품 선택 해제"):
    sel -= view_pids
    st.session_state[f"cmm_nonce_{key}"] = nonce + 1
    st.rerun()

edit_df = view[DISPLAY].copy()
edit_df.insert(0, "선택", view["상품번호"].isin(sel).values)
filter_sig = hash((tuple(sorted(pick)), only_under, only_zero, only_floor, only_miss))

edited = st.data_editor(
    edit_df,
    use_container_width=True,
    hide_index=True,
    disabled=DISPLAY,
    key=f"cmm_ed_{key}_{filter_sig}_{nonce}",
    column_config={
        "선택": st.column_config.CheckboxColumn("선택", help="가격변경/CSV로 내보낼 상품 선택"),
        "N": st.column_config.NumberColumn("N", format="%.4g", help="판매단위 배수(판매자바코드, 빈값→1, 분수 가능)"),
        "재고": st.column_config.NumberColumn("재고", format="localized"),
        "매입가": st.column_config.NumberColumn("매입가", format="localized"),
        "판매가": st.column_config.NumberColumn("판매가", format="localized"),
        "배송비": st.column_config.NumberColumn("배송비", format="localized"),
        "정산액": st.column_config.NumberColumn("정산액", format="localized"),
        "마진율": st.column_config.NumberColumn("마진율", format="percent"),
        "기준마진율": st.column_config.NumberColumn("기준마진율", format="percent"),
        "탐지": st.column_config.NumberColumn("탐지(현-기준)", format="percent"),
        "권장가/제한": st.column_config.TextColumn("권장가 / 제한", help="기준마진 달성 판매가(올림). 제한상품은 제한 텍스트."),
    },
)
# 현재 화면 체크 결과를 세션 선택집합에 반영 (화면 밖 선택은 유지)
checked_now = set(edited.loc[edited["선택"], "상품번호"])
st.session_state[f"cmm_sel_{key}"] = (sel - view_pids) | checked_now
sel = st.session_state[f"cmm_sel_{key}"]

st.caption(f"표시 {len(view):,} / 전체 {len(df):,}건 · ✅ 선택 **{len(sel):,}건**")

# ── 내보내기: CSV / 가격 일괄변경 양식 ───────────────────────────────────────
dc1, dc2 = st.columns(2)

csv_src = df[df["상품번호"].isin(sel)] if sel else view
csv_buf = StringIO()
csv_src[DISPLAY].to_csv(csv_buf, index=False)
dc1.download_button(
    f"📄 CSV 다운로드 ({'선택 '+str(len(sel)) if sel else '현재필터 '+str(len(view))}건)",
    data=csv_buf.getvalue().encode("utf-8-sig"),
    file_name=f"{channel}_마진모니터.csv",
    mime="text/csv",
    use_container_width=True,
)

if dc2.button(f"🛠️ 가격 일괄변경 양식 생성 (선택 {len(sel)}건)",
              type="primary", use_container_width=True, disabled=(len(sel) == 0)):
    new_prices, skipped = cmm.compute_new_prices(rows, recs, sel)
    if not new_prices:
        st.session_state[f"form_{key}"] = {"error": "선택 상품 중 권장가 산출 가능 항목이 없습니다(미매칭/기준 미설정)."}
    else:
        rcode, raw = _gh_bytes(_raw_path(key))
        if rcode != 200 or not raw:
            st.session_state[f"form_{key}"] = {"error": "원본 양식(.xlsx)이 저장돼 있지 않습니다. '상품관리 갱신 → 전체 교체'를 1회 실행해 주세요."}
        else:
            out, kept, missing = cmm.build_bulk_price_xlsx(raw, new_prices, cfg)
            rec_by = {r["상품번호"]: r for r in recs}
            row_by = {r["상품번호"]: r for r in rows}
            prev = []
            for pid, (np_, nd_) in new_prices.items():
                rc, ro = rec_by[pid], row_by[pid]
                cur_net = rc["판매가"] - rc["즉시할인"] - rc["포인트"]
                new_net = np_ - nd_ - rc["포인트"]
                prev.append({
                    "상품명": ro["상품명"], "현재판매가": int(rc["판매가"]), "현재할인": int(rc["즉시할인"]),
                    "새판매가": np_, "새할인": nd_, "권장가(net)": ro["권장가"],
                    "방향": "인상" if new_net > cur_net else ("인하" if new_net < cur_net else "유지"),
                })
            st.session_state[f"form_{key}"] = {
                "bytes": out, "kept": kept, "skipped": skipped, "missing": missing, "preview": prev,
                "name": f"{channel}_가격일괄변경_{datetime.now(_KST):%Y%m%d}.xlsx",
            }

form = st.session_state.get(f"form_{key}")
if form:
    if form.get("error"):
        st.warning(form["error"])
    else:
        msg = f"양식 생성 완료 — **{form['kept']}건**"
        if form["skipped"]:
            msg += f" · 제외(권장가 없음) {len(form['skipped'])}건"
        if form["missing"]:
            msg += f" · 원본에 없어 누락 {len(form['missing'])}건(상품관리 전체 교체 필요)"
        st.success(msg)
        st.download_button(
            "⬇️ 가격 일괄변경 양식 다운로드 (.xlsx)",
            data=form["bytes"], file_name=form["name"], mime=_XLSX_MIME,
            use_container_width=True,
        )
        with st.expander(f"변경 미리보기 ({form['kept']}건)", expanded=True):
            st.dataframe(pd.DataFrame(form["preview"]), use_container_width=True, hide_index=True,
                         column_config={
                             "현재판매가": st.column_config.NumberColumn(format="localized"),
                             "현재할인": st.column_config.NumberColumn(format="localized"),
                             "새판매가": st.column_config.NumberColumn(format="localized"),
                             "새할인": st.column_config.NumberColumn(format="localized"),
                             "권장가(net)": st.column_config.NumberColumn(format="localized"),
                         })
        st.caption("★ 가격은 net(판매가−즉시할인−포인트) 기준으로 기준마진 달성가에 맞춥니다. "
                   "할인 우선: 인상 시 즉시할인을 먼저 줄이고 모자라면 판매가를 올립니다(인하는 반대).")
