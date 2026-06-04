"""송장처리 — 채널 송장번호 일괄입력 템플릿에 공통 송장 마스터의 송장번호 채우기.

흐름: 공통 송장 마스터 업로드(세션) → 채널 처리전 업로드 → VLOOKUP →
      N/A 합포장 사용자 확인 → 잔존 N/A 삭제 → 원본 .xls 양식으로 다운로드.

PII 주의: 송장 마스터·채널 파일은 고객정보(수령자·주소·연락처)를 포함하므로
          서버/저장소에 저장하지 않고 이 세션에서만 사용한다.
"""
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # repo root (root page)

import streamlit as st

from core.workflows.invoice_fill import (
    CHANNEL_CONFIG, parse_template_xls, parse_master, build_master_lookup,
    vlookup_fill, find_consolidation_candidates, apply_decisions, finalize,
    write_template_xls,
)

st.title("🏷️ 송장처리")
st.caption("채널 '송장번호 일괄입력 템플릿'에 공통 송장 마스터의 송장번호를 채웁니다. "
           "원리는 전 채널 동일, 템플릿 양식만 다릅니다.")

# ── 1) 공통 송장 마스터 (세션 적재 · 미저장: 고객 PII 포함) ──────────────
st.subheader("1) 공통 송장 마스터")

if "invoice_master" in st.session_state:
    m = st.session_state["invoice_master"]
    st.success(f"📅 송장 마스터: **{m['time']}** 업로드 · 총 **{m['count']}**건")
    st.caption("판매처 분포 — " + " · ".join(f"{k} {v}" for k, v in m["dist"]))
    if st.button("🔄 마스터 다시 업로드"):
        st.session_state.pop("invoice_master", None)
        st.session_state.pop("if_work", None)
        st.session_state.pop("if_result", None)
        st.rerun()
else:
    st.warning("⚠️ 오늘자 공통 송장 마스터(.xlsx)를 먼저 업로드하세요. "
               "고객정보가 포함되어 서버에 저장하지 않고 이 세션에서만 사용합니다.")
    up = st.file_uploader("송장 마스터 .xlsx (시트: 송장출력)", type=["xlsx"], key="master_up")
    if up is not None:
        try:
            rows = parse_master(up.getvalue())
            dist = Counter(str(r.get("판매처")) for r in rows).most_common()
            st.session_state["invoice_master"] = {
                "rows": rows, "count": len(rows),
                "time": datetime.now().strftime("%H:%M"), "dist": dist,
            }
            st.rerun()
        except Exception as e:
            st.error(f"마스터 읽기 오류: {e}")
            st.exception(e)

# 게이트 — 마스터 없으면 중단 (단일 페이지라 st.stop 안전)
if "invoice_master" not in st.session_state:
    st.stop()

master = st.session_state["invoice_master"]["rows"]

# ── 2) 채널 처리 ──────────────────────────────────────────────────────
st.divider()
st.subheader("2) 채널 처리")

channel = st.selectbox("채널", list(CHANNEL_CONFIG.keys()))
cfg = CHANNEL_CONFIG[channel]
st.caption(f"매칭 규칙 — 처리전.`{cfg['match_col']}` ⟷ 마스터.`{cfg['master_key']}`")

before_up = st.file_uploader(f"{channel} 처리전 템플릿 (.xls)", type=["xls"],
                             key=f"before_{channel}")

if before_up is not None and st.button("🔍 분석", type="primary"):
    try:
        parsed = parse_template_xls(before_up.getvalue())
        lk = build_master_lookup(master, cfg["master_key"])
        rows = vlookup_fill(parsed["rows"], lk, channel)
        cands, indep = find_consolidation_candidates(rows)
        st.session_state["if_work"] = {
            "channel": channel, "parsed": parsed,
            "cands": cands, "indep": indep, "matched": sum(
                1 for r in rows if r["_status"] == "matched"),
            "total": len(rows),
        }
        st.session_state.pop("if_result", None)
    except Exception as e:
        st.error(f"분석 오류: {e}")
        st.exception(e)

work = st.session_state.get("if_work")
if work and work["channel"] == channel:
    cands, indep = work["cands"], work["indep"]
    st.info(f"VLOOKUP 매칭 **{work['matched']}** / N/A **{work['total'] - work['matched']}** "
            f"— 합포장 후보 **{len(cands)}**건 · 독립 N/A **{len(indep)}**건")

    decisions = {}
    if cands:
        st.markdown("##### 📦 합포장 확인")
        st.caption("같은 주소에 송장(박스)이 있는 N/A 건입니다. 어느 박스로 합칠지 선택하세요. "
                   "(자동 선택 불가 — 어느 박스에 담겼는지는 사람만 압니다)")
        for c in cands:
            na = c["na_row"]
            recv = na.get("수취인명(받는사람)") or na.get("수취인") or ""
            st.markdown(f"**N/A**: {recv} · {na.get('상품명')}  \n"
                        f"`{na.get('배송지')}`")
            opts = ["❌ 합포장 아님 (N/A 유지 → 삭제)"] + [
                f"📦 {b['송장']} · {b['수취인']} · {b['상품명']}" for b in c["boxes"]]
            sel = st.radio("합칠 박스 선택", options=list(range(len(opts))),
                           format_func=lambda i, o=opts: o[i],
                           key=f"cons_{channel}_{c['na_index']}")
            decisions[c["na_index"]] = None if sel == 0 else c["boxes"][sel - 1]["송장"]
            st.markdown("---")

    if st.button("✅ 확정 & 결과 생성", key=f"finalize_{channel}", type="primary"):
        lk = build_master_lookup(master, cfg["master_key"])
        rows = vlookup_fill(work["parsed"]["rows"], lk, channel)   # 깨끗이 재계산(멱등)
        apply_decisions(rows, decisions)
        keep, na_count, na_rows = finalize(rows)
        out_bytes = write_template_xls(work["parsed"], rows, keep)
        st.session_state["if_result"] = {
            "channel": channel, "bytes": out_bytes, "keep": len(keep),
            "na_count": na_count,
            "na_rows": [((r.get("수취인명(받는사람)") or r.get("수취인") or ""),
                         r.get("상품명"), r.get("배송지")) for r in na_rows],
        }

# ── 3) 결과 다운로드 (버튼 블록 밖 — rerun 시 위젯 소멸 방지) ────────────
res = st.session_state.get("if_result")
if res and res["channel"] == channel:
    st.divider()
    st.success(f"✅ 완료 · 최종 **{res['keep']}**행 · N/A 삭제 **{res['na_count']}**건")
    if res["na_count"]:
        with st.expander(f"삭제된 N/A {res['na_count']}건 보기"):
            for recv, prod, addr in res["na_rows"]:
                st.write(f"- {recv} · {prod} · `{addr}`")
    fname = f"{channel}배송{datetime.now().strftime('%Y%m%d')}_처리후_.xls"
    st.download_button("⬇️ 처리후 .xls 다운로드", data=res["bytes"], file_name=fname,
                       mime="application/vnd.ms-excel", key=f"dl_{channel}", type="primary")
