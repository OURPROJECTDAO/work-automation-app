"""채널 가격·마진 모니터 (channel-margin-monitor).

저장된 상품관리(listing) 스냅샷을 자동 로드해 상품별 마진율·기준마진 대비 탐지·권장가(또는
제한 텍스트)·재고를 계산. 다운로드는 매번 올릴 필요 없이 '상품관리 갱신'에서 전체 교체/신규 추가.
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

st.title("💹 채널 가격·마진 모니터")
st.caption(
    "저장된 상품관리 기준으로 마진율·기준마진 대비 이탈·권장가를 계산합니다. "
    "매입가/재고는 상품관리(product_master), 기준마진율은 baseline_margin 기준. "
    "다운로드는 '상품관리 갱신'에서 새로 올릴 때만 갱신."
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


def _listing_path(key: str) -> str:
    return f"reference/listing_{key}.csv"


def _meta_path(key: str) -> str:
    return f"reference/listing_{key}.meta.json"


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
        new_recs = cmm.parse_download(up.getvalue(), cfg)
        st.write(f"업로드 파싱: **{len(new_recs):,}건**")
        if not _pat():
            st.error("저장용 PAT(st.secrets GITHUB_PAT)가 없어 커밋할 수 없습니다.")
        else:
            b1, b2 = st.columns(2)
            if b1.button("전체 교체 저장", type="primary", use_container_width=True,
                         help="최신 전체 다운로드로 덮어쓰기 (신규+가격변동 반영)"):
                meta = _commit_listing(key, new_recs)
                _load_listing.clear()
                committed = new_recs
                flash = f"전체 교체 완료 — {meta['rows']:,}건 ({meta['updated_at']})"
            if b2.button("신규만 추가", use_container_width=True,
                         help="기존 유지 + 새 상품번호만 병합"):
                cur, _ = _load_listing(key)
                merged, added = cmm.merge_listing(cur or [], new_recs)
                meta = _commit_listing(key, merged)
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
c3.metric("마진 미달", f"{stats['마진미달']:,}")
c4.metric("제한 상품", f"{stats['제한상품']:,}")
c5.metric("기준 미설정", f"{stats['미설정']:,}")
c6.metric("미매칭", f"{stats['미매칭']:,}")

# ── 필터 ──────────────────────────────────────────────────────────────────────
df = pd.DataFrame(rows)
types = sorted(df["코드유형"].unique().tolist())
fc1, fc2 = st.columns([2, 3])
with fc1:
    pick = st.multiselect("코드유형", types, default=types)
with fc2:
    f1, f2, f3, f4 = st.columns(4)
    only_under = f1.checkbox("마진미달만")
    only_zero = f2.checkbox("재고 0")
    only_floor = f3.checkbox("제한상품만")
    only_miss = f4.checkbox("미매칭만")

view = df[df["코드유형"].isin(pick)].copy()
if only_under:
    view = view[view["탐지"].notna() & (view["탐지"] < 0)]
if only_zero:
    view = view[view["재고"].fillna(-1) == 0]
if only_floor:
    view = view[view["제한"].astype(str) != ""]
if only_miss:
    view = view[view["매입가"].isna()]


def _rec_disp(r):
    if r["제한"]:
        return str(r["제한"])
    if r["매입가"] is None:
        return str(r["비고"]) or "미매칭"
    if r["기준마진율"] is None:
        return "기준 미설정"
    return f"{int(r['권장가']):,}" if pd.notna(r["권장가"]) else ""


view["권장가/제한"] = view.apply(_rec_disp, axis=1)

DISPLAY = ["상품번호", "관리코드", "상품명", "규격", "코드유형", "N", "재고",
           "매입가", "판매가", "배송비", "정산액", "마진율", "기준마진율", "탐지", "권장가/제한", "비고"]

st.dataframe(
    view[DISPLAY],
    use_container_width=True,
    hide_index=True,
    column_config={
        "N": st.column_config.NumberColumn("N", format="%.4g", help="판매단위 배수(판매자바코드, 빈값→1, 분수 가능)"),
        "재고": st.column_config.NumberColumn("재고", format="localized"),
        "매입가": st.column_config.NumberColumn("매입가", format="localized"),
        "판매가": st.column_config.NumberColumn("판매가", format="localized"),
        "배송비": st.column_config.NumberColumn("배송비", format="localized"),
        "정산액": st.column_config.NumberColumn("정산액", format="localized"),
        "마진율": st.column_config.NumberColumn("마진율", format="percent"),
        "기준마진율": st.column_config.NumberColumn("기준마진율", format="percent"),
        "탐지": st.column_config.NumberColumn("탐지(현-기준)", format="percent"),
        "권장가/제한": st.column_config.TextColumn("권장가 / 제한", help="기준마진 달성 판매가. 제한상품은 제한 텍스트."),
    },
)
st.caption(f"표시 {len(view):,} / 전체 {len(df):,}건")

buf = StringIO()
view[DISPLAY].to_csv(buf, index=False)
st.download_button(
    "CSV 다운로드",
    data=buf.getvalue().encode("utf-8-sig"),
    file_name=f"{channel}_마진모니터.csv",
    mime="text/csv",
)
