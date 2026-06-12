"""업로드감시 (upload-monitor) — 채널 등록 갭/품절 정합.

박스재고가 있는데 각 채널에 아직 등록(업로드) 안 된 관리코드를 탐지
(+ 재고0인데 채널엔 살아있는 = 품절처리 대상). 등록현황 = 채널마진모니터 listing 스냅샷 공유,
재고·매입가 = 상품관리(product_master). 키=상품코드, 우선순위=재고금액(박스재고×박스매입가).
근거 = KB workflows/upload-monitor.md (ADR 0017).
"""
import sys
from io import StringIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # repo root

import pandas as pd
import streamlit as st

from core.workflows import upload_monitor as um

_REF = Path(__file__).parent.parent.parent / "reference"

st.title("📦 업로드감시")
st.caption(
    "박스재고가 있는데 채널에 아직 등록 안 된 상품(업로드필요)과, 재고0인데 채널엔 살아있는 "
    "상품(품절처리필요)을 채널별로 점검합니다. 등록현황은 채널마진모니터 상품관리 스냅샷 기준, "
    "재고·매입가는 상품관리(product_master) 기준. 우선순위는 재고금액(박스재고×박스매입가)입니다. "
    "낱개·박스·소분·합포 중 하나라도 올라가 있으면 등록된 것으로 봅니다."
)


@st.cache_data(ttl=600, show_spinner="재고 · 채널 등록현황 대조 중...")
def _load():
    return um.build_gap_table(str(_REF))


rows = _load()
if not rows:
    st.warning("데이터가 없습니다. 상품관리(product_master)와 채널 listing 저장본을 확인하세요.")
    st.stop()

df = pd.DataFrame(rows)
KEYS = um.CHANNEL_KEYS

# ── KPI ──────────────────────────────────────────────────────────────────────
need_any = int((df[KEYS] == um.ST_NEED_UP).any(axis=1).sum())
sold_any = int((df[KEYS] == um.ST_NEED_SOLD).any(axis=1).sum())
k1, k2, k3 = st.columns(3)
k1.metric("감시대상 상품", f"{len(df):,}")
k2.metric("업로드필요", f"{need_any:,}", help="한 채널 이상에서 업로드필요")
k3.metric("품절처리필요", f"{sold_any:,}", help="재고0인데 채널엔 등록·판매 중")

st.markdown("**채널별 미업로드 / 품절처리 건수**")
sdf = pd.DataFrame(um.channel_summary(rows))[["label", "업로드필요", "품절처리필요"]].rename(
    columns={"label": "채널"})
st.dataframe(sdf, hide_index=True, use_container_width=True,
             column_config={"업로드필요": st.column_config.NumberColumn(format="localized"),
                            "품절처리필요": st.column_config.NumberColumn(format="localized")})

# ── 채널 선택 (체크박스 + 전체 선택/해제) ────────────────────────────────────
ALL_STATUS = ["(전체)", um.ST_NEED_UP, um.ST_OK, um.ST_NEED_SOLD, um.ST_SKIP]


def _toggle_all_channels():
    all_on = all(st.session_state.get(f"um_ch_{k}", True) for k in KEYS)
    for k in KEYS:
        st.session_state[f"um_ch_{k}"] = not all_on


tcol, _ = st.columns([1, 5])
tcol.button("전체 선택/해제", on_click=_toggle_all_channels, use_container_width=True)
ch_boxes = st.columns(len(um.CHANNELS))
for (k, label, _au), col in zip(um.CHANNELS, ch_boxes):
    col.checkbox(label, value=True, key=f"um_ch_{k}")
selected = [k for k in KEYS if st.session_state.get(f"um_ch_{k}", True)]

# ── 컬럼별(헤더) 상태 필터 — 선택 채널만 ──────────────────────────────────────
col_status: dict[str, str] = {}
if selected:
    st.caption("컬럼 필터 (선택 채널별 상태)")
    fcols = st.columns(len(selected))
    for k, col in zip(selected, fcols):
        col_status[k] = col.selectbox(um.CHANNEL_LABEL[k], ALL_STATUS, index=0, key=f"um_st_{k}")
else:
    st.caption("채널을 하나 이상 선택하면 해당 상태 컬럼과 컬럼 필터가 표시됩니다.")

search = st.text_input("🔍 검색", placeholder="관리코드 · 상품코드 · 상품명 (부분일치)",
                       label_visibility="collapsed")

# ── 행 필터 (선택 채널 상태 AND + 검색) ───────────────────────────────────────
view = df.copy()
for k in selected:
    s = col_status.get(k, "(전체)")
    if s != "(전체)":
        view = view[view[k] == s]
if search:
    q = um._nfc(search).lower()
    hay = (view["관리코드"].astype(str) + " ||| " + view["상품코드"].astype(str)
           + " ||| " + view["상품명"].astype(str)).str.lower()
    view = view[hay.str.contains(q, regex=False, na=False)]

# 표시: 기본정보 + 선택 채널 컬럼(재고금액 옆). 채널키→라벨 헤더.
base_cols = ["상품코드", "관리코드", "상품명", "박스재고", "박스매입가", "재고금액"]
view_reset = view.reset_index(drop=True)
disp = view_reset[base_cols + selected].rename(columns=um.CHANNEL_LABEL)

filter_sig = hash((tuple(selected),
                   tuple(col_status.get(k, "(전체)") for k in selected),
                   um._nfc(search) if search else "", len(view_reset)))
event = st.dataframe(
    disp,
    use_container_width=True,
    hide_index=True,
    on_select="rerun",
    selection_mode="multi-row",
    key=f"um_df_{filter_sig}",
    column_config={
        "박스재고": st.column_config.NumberColumn(format="localized"),
        "박스매입가": st.column_config.NumberColumn(format="localized"),
        "재고금액": st.column_config.NumberColumn(format="localized"),
    },
)
_sel = event.selection.rows if event and getattr(event, "selection", None) else []
sel_rows = [i for i in _sel if 0 <= i < len(view_reset)]
sel_codes = view_reset.iloc[sel_rows]["상품코드"].tolist() if sel_rows else []

amt = int(view["재고금액"].sum()) if len(view) else 0
st.caption(f"표시 {len(view):,} / 전체 {len(df):,}건 · 재고금액합 {amt:,}원 · "
           f"✅ 선택 **{len(sel_codes):,}건** (표 왼쪽 체크박스로 개별, 헤더로 화면 전체)")

# ── 내보내기 (CSV) ────────────────────────────────────────────────────────────
csv_src = view_reset.iloc[sel_rows] if sel_rows else view_reset
exp = csv_src[base_cols + selected].rename(columns=um.CHANNEL_LABEL)
buf = StringIO()
exp.to_csv(buf, index=False)
if len(selected) == 1:
    tag = um.CHANNEL_LABEL[selected[0]]
elif len(selected) == len(KEYS) or not selected:
    tag = "전체"
else:
    tag = f"{len(selected)}채널"
st.download_button("📥 CSV 다운로드 (선택 또는 현재 화면)",
                   buf.getvalue().encode("utf-8-sig"),
                   file_name=f"업로드감시_{tag}.csv", mime="text/csv")

st.info("스마트스토어·ESM **등록폼 자동생성(L4)** 과 비판매 제외목록 편집은 다음 단계에서 추가됩니다. "
        "지금은 채널별 갭을 CSV로 받아 활용할 수 있습니다.")
