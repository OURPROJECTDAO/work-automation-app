"""🩸 마진 침식 — 최근 매입가가 올라 채널 기준마진 아래로 떨어진(아직 미수정) 상품.

두뇌① (intelligence-layer.md §6 ①). 가격이력(1a) '최근 N개월 매입가 인상' ∩ 채널 listing 마진 미달.
사용자가 이미 가격 재설정한 건 마진이 기준 이상 → 자동 제외. 8채널 통합 작업목록.
"""
import io
import json
import sys
import urllib.error
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # repo root

import pandas as pd
import streamlit as st

from core.workflows import channel_margin_monitor as cmm
from core.intelligence import margin_erosion as me
from core.intelligence import price_history

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


@st.cache_data(ttl=600, show_spinner=False)
def _listing(key: str):
    code, text = _gh_raw(f"reference/listing_{key}.csv")
    if code != 200:
        return None, {}
    recs = cmm.csv_text_to_recs(text.decode("utf-8-sig"))
    mcode, mtext = _gh_raw(f"reference/listing_{key}.meta.json")
    meta = json.loads(mtext) if mcode == 200 else {}
    return recs, meta


@st.cache_data(ttl=600, show_spinner=False)
def _baseline_override():
    code, text = _gh_raw("reference/baseline_margin.csv")
    return cmm.parse_baseline_dict(text.decode("utf-8-sig")) if code == 200 else None


@st.cache_data(ttl=600, show_spinner="가격이력 불러오는 중...")
def _price_changes() -> pd.DataFrame:
    pat, repo = _data_secret()
    if not pat:
        return pd.DataFrame()
    return price_history.read_history(pat, repo)


def _to_xlsx(df: pd.DataFrame) -> bytes:
    from openpyxl import Workbook

    def _cell(v):
        if pd.isna(v):
            return None
        if hasattr(v, "isoformat"):
            return v.isoformat()
        if hasattr(v, "item"):
            return v.item()
        return v

    wb = Workbook()
    ws = wb.active
    ws.title = "마진침식"
    ws.append(list(df.columns))
    for _, r in df.iterrows():
        ws.append([_cell(v) for v in r])
    if "관리코드" in df.columns:
        col = df.columns.get_loc("관리코드") + 1
        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            for c in row:
                c.number_format = "@"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────
st.title("🩸 마진 침식")
st.caption("최근 매입가가 올라 채널 기준마진(baseline) 아래로 떨어졌는데 아직 가격을 안 고친 상품 — 8채널 통합. "
           "이미 기준에 맞게 재설정한 상품은 안 뜹니다.")

pc = _price_changes()
if pc.empty:
    st.warning("가격이력(price_changes)이 없습니다. secrets `[data] pat` 또는 수정로그 적재를 확인해 주세요.")
    st.stop()

_chan_keys = {ch: cfg["key"] for ch, cfg in cmm.CHANNEL_CONFIG.items()}

c1, c2, c3 = st.columns([1.2, 1.3, 2.5])
with c1:
    months = st.number_input("매입가 인상 기준 기간(개월)", min_value=1, max_value=12, value=3, step=1,
                             help="최근 몇 개월 안에 매입가가 오른 것을 침식 후보로 볼지")
with c2:
    sort_label = st.radio("정렬", ["미달폭 큰 순", "매입 상승폭 큰 순"])
with c3:
    chans = st.multiselect("채널", list(_chan_keys.keys()), default=list(_chan_keys.keys()))

sort_key = "미달폭" if sort_label.startswith("미달폭") else "매입Δ%"

_pc_last = pd.to_datetime(pc["수정일자"]).max()
raises = me.recent_buy_raises(pc, months=int(months), now=pd.Timestamp.now())
baseline_override = _baseline_override()

all_rows, fresh = [], []
for ch in chans:
    key = _chan_keys[ch]
    recs, meta = _listing(key)
    if recs is None:
        fresh.append((ch, "—(미저장)"))
        continue
    fresh.append((ch, (meta.get("updated_at") or "?")[:10]))
    rows, _ = cmm.compute_listing(recs, ch, str(_REF), baseline_override=baseline_override)
    all_rows += me.erosion_rows(rows, ch, raises)

df = me.to_frame(all_rows, sort=sort_key)

k1, k2, k3, k4 = st.columns(4)
k1.metric("침식 경보", f"{len(df)} 건")
k2.metric("관리코드", f"{df['관리코드'].nunique() if not df.empty else 0} 개")
k3.metric("채널", f"{df['채널'].nunique() if not df.empty else 0} 개")
k4.metric("가격이력 최신", _pc_last.strftime("%Y-%m-%d") if pd.notna(_pc_last) else "—")

with st.expander("채널 listing 신선도 (이 날짜의 판매가로 마진 계산 — 오래되면 갱신 권장)"):
    st.dataframe(pd.DataFrame(fresh, columns=["채널", "listing 갱신일"]),
                 hide_index=True, use_container_width=True)

if df.empty:
    st.success("✅ 최근 매입가 인상으로 기준마진 아래로 떨어진(미수정) 상품이 없습니다.")
    st.stop()

st.caption(f"최근 **{int(months)}개월** 매입가 인상 상품 중, 각 채널에서 아직 기준마진 미달인 건만. "
           f"권장가 = 기준마진 회복가(100원 올림). 정렬: {sort_label}.")

# 표시용: 비율 컬럼 ×100 (% 표기)
_disp = df.copy()
for c in ("매입Δ%", "마진율", "기준마진율", "미달폭"):
    _disp[c] = (_disp[c].astype(float) * 100).round(1)

event = st.dataframe(
    _disp,
    hide_index=True, use_container_width=True, height=460,
    on_select="rerun", selection_mode="single-row",
    column_config={
        "채널": st.column_config.TextColumn(width="small"),
        "관리코드": st.column_config.TextColumn(width="small"),
        "상품명": st.column_config.TextColumn(width="medium"),
        "과거매입": st.column_config.NumberColumn(format="%d"),
        "현재매입": st.column_config.NumberColumn(format="%d"),
        "매입Δ%": st.column_config.NumberColumn("매입↑%", format="%.0f", help="과거→현재 매입가 상승률(%)"),
        "마지막인상일": st.column_config.DatetimeColumn("마지막인상", format="YYYY-MM-DD"),
        "현재판매가": st.column_config.NumberColumn(format="%d"),
        "마진율": st.column_config.NumberColumn("현재마진%", format="%.1f"),
        "기준마진율": st.column_config.NumberColumn("기준마진%", format="%.1f"),
        "미달폭": st.column_config.NumberColumn("미달폭%p", format="%.1f", help="현재마진−기준마진(음수=미달)"),
        "권장가": st.column_config.NumberColumn(format="%d"),
        "재고": st.column_config.NumberColumn("박스재고", format="%d"),
        "상품수": st.column_config.NumberColumn("코드수", format="%d", width="small",
                                             help="이 관리코드의 매입 인상 상품코드 수(>1=충돌)"),
    },
)

sel = []
try:
    sel = list(event.selection["rows"])
except Exception:
    try:
        sel = list(event["selection"]["rows"])
    except Exception:
        sel = []

if sel:
    r = df.iloc[sel[0]]
    code = str(r["관리코드"])
    h = pc[(pc["관리코드"].astype(str) == code) & (pc["수정항목"] == "매입단가")].sort_values("수정일자")
    st.markdown(f"**{code} · {r['상품명']}** — 매입가 변경 이력")
    if h.empty:
        st.caption("이 관리코드의 매입가 변경 이력이 없습니다(상품코드 키 차이일 수 있음).")
    else:
        ser = pd.concat([
            h[["수정일자", "수정전"]].rename(columns={"수정전": "매입가"}).head(1),
            h[["수정일자", "수정후"]].rename(columns={"수정후": "매입가"}),
        ]).set_index("수정일자")["매입가"]
        st.line_chart(ser, height=220)

st.download_button("📥 XLSX 다운로드", data=_to_xlsx(df), file_name="마진침식.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
