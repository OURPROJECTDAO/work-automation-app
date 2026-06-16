"""파일 처리 페이지: 발주 파일 업로드 → 워크플로우 실행 → 결과 다운로드."""
import sys, tempfile
from datetime import datetime, timezone, timedelta

_KST = timezone(timedelta(hours=9))  # 한국 표준시 UTC+9
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import streamlit as st
import pandas as pd
import core.workflows.openmarket_merge  # noqa: F401  ← @register 트리거
import core.workflows.onnuri_order      # noqa: F401  ← @register 트리거
from core.workflows.registry import list_workflows, get_workflow
import core.workflows.logistics_order as lo
import core.workflows.cheonnyeon_upload as cy
import core.intelligence.daily_inbox as _inbox
import core.intelligence.stockout_board as _sb

def _seed_stockout_board(so_df):
    """발주 품절목록 → 품절 알림판 자동 등록(데일리 대시보드·data repo 영속). 실패는 발주 흐름 비차단."""
    try:
        if so_df is None or len(so_df) == 0:
            return
        d = st.secrets["data"]
        pat = d["pat"]; repo = d.get("repo", "OURPROJECTDAO/work-automation-data")
        today = datetime.now(_KST).strftime("%Y-%m-%d")
        bd = _sb.read_board(pat, repo)
        bd, added = _sb.seed_from_stockout(bd, so_df, today)
        if added:
            _sb.write_board(pat, repo, bd, f"board: 발주 품절 {len(added)}건 등록 ({today})")
    except Exception:
        pass


st.title("📂 파일 처리")

tab_basic, tab_order, tab_cy = st.tabs(["📦 기존 워크플로우", "🚚 발주서 출력업무", "🏪 천년경영 업로드"])

# ═══════════════════════════════════════════════
# 탭 1 : 기존 워크플로우
# ═══════════════════════════════════════════════
with tab_basic:
    st.caption("마켓플레이스 발주 파일(.xls)을 업로드하면 합포·도서산간·필터링 결과를 한 번에 만들어드립니다.")
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded = st.file_uploader(
            "발주 파일 업로드",
            type=["xls", "xlsx", "xlsm"],
            help="스마트스토어·쿠팡·G마켓 등에서 다운로드한 .xls 발주 파일",
            key="basic_upload",
        )
    with col2:
        workflows = list_workflows()
        workflow_name = st.selectbox("워크플로우", workflows, key="basic_wf") if workflows else None
        if not workflows:
            st.warning("등록된 워크플로우 없음")

    if uploaded and workflow_name:
        if st.button("▶ 실행", type="primary", use_container_width=True, key="basic_run"):
            with st.spinner("처리 중..."):
                try:
                    with tempfile.TemporaryDirectory() as tmp:
                        tmp_path = Path(tmp)
                        input_path = tmp_path / uploaded.name
                        input_path.write_bytes(uploaded.getvalue())
                        output_dir = tmp_path / "output"
                        output_dir.mkdir()
                        result = get_workflow(workflow_name).run(input_path, output_dir)
                        result_bytes = result.read_bytes()
                        result_name  = result.name

                    import io as _io
                    sheets = pd.read_excel(_io.BytesIO(result_bytes), sheet_name=None, dtype=str)
                    labels = {"합포확인": "합포", "지역확인": "도서산간",
                              "필터링확인": "필터링", "미배송지역확인": "미배송", "송장출력": "전체 송장"}
                    # 송장출력 단독 파일 (VBA SaveSheetToNewFile 복원)
                    invoice_bytes = (
                        core.workflows.openmarket_merge.generate_invoice_xlsx(sheets["송장출력"])
                        if "송장출력" in sheets else None
                    )
                    # 결과를 session_state에 보관 → download_button 클릭(rerun) 후에도 유지
                    st.session_state["basic_result"] = {
                        "result_bytes": result_bytes,
                        "result_name":  result_name,
                        "invoice_bytes": invoice_bytes,
                        "mmdd": datetime.now(_KST).strftime("%m%d"),
                        "stats": {labels.get(s, s): len(df) for s, df in sheets.items()},
                    }
                    if invoice_bytes:   # 데일리 대시보드 자동 인계(송장출력)
                        _inbox.push(st.session_state, _inbox.SLOT_INVOICE, invoice_bytes,
                                    f"★★송장{datetime.now(_KST).strftime('%m%d')}.xlsx",
                                    datetime.now(_KST).strftime("%m-%d %H:%M"))
                except Exception as e:
                    st.session_state.pop("basic_result", None)
                    st.error(f"오류: {e}")
                    st.exception(e)
    else:
        st.info("파일을 업로드하고 워크플로우를 선택하세요.")

    # ── 결과 렌더 (실행 블록 밖에서: download_button rerun에도 버튼 유지) ──
    res = st.session_state.get("basic_result")
    if res:
        st.success("✅ 처리 완료!")
        cols = st.columns(len(res["stats"]))
        for col, (label, n) in zip(cols, res["stats"].items()):
            col.metric(label, f"{n}건")

        st.download_button(
            label="📥 결과 파일 다운로드",
            data=res["result_bytes"],
            file_name=res["result_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="basic_result_dl",
        )
        if res["invoice_bytes"]:
            st.download_button(
                label="📥 송장 파일 다운로드 (★★송장)",
                data=res["invoice_bytes"],
                file_name=f"★★송장{res['mmdd']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="basic_invoice_dl",
            )


# ═══════════════════════════════════════════════
# 탭 2 : 천년경영 업로드
# ═══════════════════════════════════════════════
with tab_cy:
    st.caption("발주자료·배민주문·스스주문(암호 1323) 3개를 올리면 "
               "마켓플레이스별 업로드 시트(전체/낱개)를 만들어드립니다.")
    c1, c2, c3 = st.columns(3)
    f_baeju  = c1.file_uploader("① ★★발주자료 (.xlsx)", type=["xlsx"], key="cy_baeju",
                                help="발주서출력업무 Phase1에서 받은 발주자료 아카이브")
    f_baemin = c2.file_uploader("② 배민주문 (.xlsx)", type=["xlsx"], key="cy_baemin")
    f_sss    = c3.file_uploader("③ 스스주문 (.xlsx · 암호1323)", type=["xlsx"], key="cy_sss")

    if f_baeju and f_baemin and f_sss:
        if st.button("▶ 실행", type="primary", use_container_width=True, key="cy_run"):
            with st.spinner("처리 중..."):
                try:
                    out, stats, sheets, _ = cy.run(
                        f_baeju.getvalue(), f_baemin.getvalue(), f_sss.getvalue())
                    st.session_state["cy_result"] = {
                        "bytes": out, "stats": stats,
                        "anomalies": cy.detect_box_anomalies(sheets),
                        "name": datetime.now(_KST).strftime("%y%m%d") + ".xlsx",
                    }
                    _inbox.push(st.session_state, _inbox.SLOT_CHEONNYEON, out,   # 데일리 대시보드 자동 인계
                                datetime.now(_KST).strftime("%y%m%d") + ".xlsx",
                                datetime.now(_KST).strftime("%m-%d %H:%M"))
                except Exception as e:
                    st.session_state.pop("cy_result", None)
                    st.error(f"오류: {e}")
                    st.exception(e)
    else:
        st.info("3개 파일을 모두 업로드하세요.")

    res = st.session_state.get("cy_result")
    if res:
        st.success("✅ 처리 완료!")
        items = list(res["stats"].items())
        for i in range(0, len(items), 6):
            cols = st.columns(6)
            for col, (label, n) in zip(cols, items[i:i + 6]):
                col.metric(label, f"{n}건")

        anomalies = res.get("anomalies") or []
        if anomalies:
            st.error(
                f"⚠️ 전체(박스) 시트에 **박스 코드가 아닌 상품 {len(anomalies)}건**이 "
                "있습니다. **소분목록 누락**이 의심됩니다 — 업로드 전 확인하세요.",
                icon="🚨",
            )
            st.dataframe(
                pd.DataFrame(anomalies)[["시트", "관리코드", "상품명", "신호", "확신"]],
                use_container_width=True, hide_index=True,
            )
            st.caption(
                "낱개로 파는 상품인데 소분목록(낱개코드→원코드)에 없으면 전체 시트에 "
                "그대로 남습니다(박스로 잘못 업로드). 👉 **기준데이터관리 → 천년경영업로드**에서 "
                "해당 코드를 소분목록에 추가한 뒤 다시 실행하세요."
            )

        st.download_button(
            label="📥 결과 파일 다운로드",
            data=res["bytes"],
            file_name=res["name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="cy_result_dl",
        )


# ═══════════════════════════════════════════════
# 탭 3 : 발주서 출력업무
# ═══════════════════════════════════════════════
with tab_order:

    # ── 상품관리 타임스탬프 확인 ────────────────
    updated = lo.get_product_master_updated()
    if updated:
        proceed = st.warning(
            f"📅 상품관리 마지막 업데이트: **{updated}**  \n계속 진행하시겠습니까?",
            icon="⚠️",
        )
        if "order_confirmed" not in st.session_state:
            st.session_state.order_confirmed = False
        confirmed = st.checkbox("네, 진행합니다", key="order_confirm_chk",
                                value=st.session_state.order_confirmed)
        st.session_state.order_confirmed = confirmed
    else:
        st.error("⚠️ 상품관리 데이터 없음 — 연동데이터관리에서 먼저 업로드해주세요.")
        confirmed = False

    if not confirmed:
        st.stop()

    st.divider()

    # ── 매출통계 업로드 ─────────────────────────
    sales_file = st.file_uploader(
        "📊 판매처상품매출통계 업로드 (.xls)",
        type=["xls"],
        key="order_sales_upload",
    )

    if sales_file:
        sales_bytes = sales_file.getvalue()
        st.session_state["order_sales_bytes"] = sales_bytes

    if "order_sales_bytes" not in st.session_state:
        st.info("매출통계 파일을 업로드하세요.")
        st.stop()

    # ─────────────────────────────────────────────
    # Phase 1
    # ─────────────────────────────────────────────
    st.subheader("Phase 1 — 정제 + 분류")

    if st.button("▶ Phase 1 실행", type="primary", key="p1_run"):
        with st.spinner("정제 중..."):
            try:
                result, unmatched, pre_cls, archive_df = lo.run_phase1(
                    st.session_state["order_sales_bytes"]
                )
                st.session_state["order_archive_df"] = archive_df
                if unmatched:
                    st.session_state["order_unmatched_cls"] = unmatched
                    st.session_state["order_pre_cls_df"]    = pre_cls
                    st.session_state.pop("order_phase1_df", None)
                else:
                    st.session_state["order_phase1_df"]     = result
                    st.session_state.pop("order_unmatched_cls", None)
                    st.session_state.pop("order_pre_cls_df", None)
            except Exception as e:
                st.error(f"오류: {e}"); st.exception(e)

    # ── GATE A : 미분류 코드 처리 ───────────────
    if "order_unmatched_cls" in st.session_state:
        unmatched = st.session_state["order_unmatched_cls"]
        st.warning(f"⚠️ {len(unmatched)}개 코드가 분류되지 않았습니다.")

        cls_df  = lo.load_classification()
        choices = {}
        with st.form("gate_a_form"):
            for row in unmatched:
                code  = row["erp관리코드"]
                _admin = row.get("어드민옵션", "")
                _empty = pd.isna(_admin) or str(_admin).strip() == ""
                label = "⚠️ 어드민옵션 비어있음 (코드로 분류)" if _empty else str(_admin)[:30]
                choices[code] = st.selectbox(
                    f"`{code}`  {label}",
                    ["음료", "식품", "선물세트"],
                    key=f"cls_{code}",
                )
            submitted = st.form_submit_button("✅ 분류 저장 후 재실행", type="primary")

        if submitted:
            import io as _io, base64 as _b64, urllib.request as _ur, json as _json

            token = st.secrets.get("GITHUB_PAT", "")
            # 분류표에 추가
            new_rows = pd.DataFrame([
                {"관리코드": c, "구분": g} for c, g in choices.items()
            ])
            updated_cls = pd.concat([cls_df, new_rows], ignore_index=True).drop_duplicates("관리코드")

            csv_bytes = updated_cls.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            b64 = _b64.b64encode(csv_bytes).decode()

            # sha 조회
            path = "reference/logistics_classification.csv"
            req = _ur.Request(
                f"https://api.github.com/repos/OURPROJECTDAO/work-automation-app/contents/{path}",
                headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"},
            )
            with _ur.urlopen(req) as r:
                sha = _json.loads(r.read())["sha"]

            body = _json.dumps({"message": "ref: 분류표 신규 코드 추가", "content": b64, "sha": sha}).encode()
            req2 = _ur.Request(
                f"https://api.github.com/repos/OURPROJECTDAO/work-automation-app/contents/{path}",
                data=body, method="PUT",
                headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json",
                         "Content-Type": "application/json"},
            )
            _ur.urlopen(req2)

            with st.spinner("재실행 중..."):
                result, unmatched2, pre_cls2 = lo.resume_phase1_after_gate(
                    st.session_state["order_pre_cls_df"], cls_df=updated_cls
                )
                if unmatched2:
                    st.session_state["order_unmatched_cls"] = unmatched2
                    st.session_state["order_pre_cls_df"]    = pre_cls2
                else:
                    st.session_state["order_phase1_df"] = result
                    st.session_state.pop("order_unmatched_cls", None)
                    st.session_state.pop("order_pre_cls_df", None)
                    st.success("✅ Phase 1 완료!")
            st.rerun()

    # ── Phase 1 완료 표시 ───────────────────────
    if "order_phase1_df" in st.session_state:
        p1_df = st.session_state["order_phase1_df"]
        c1, c2, c3 = st.columns(3)
        c1.metric("분류된 상품 수", len(p1_df))
        c2.metric("음료", len(p1_df[p1_df["구분"] == "음료"]))
        c3.metric("식품 + 선물세트",
                  len(p1_df[p1_df["구분"].isin(["식품", "선물세트"])]))
        st.success("✅ Phase 1 완료")

        archive_bytes = lo.generate_archive_xlsx(st.session_state["order_archive_df"])
        today = datetime.now(_KST).strftime("%m%d")
        st.download_button(
            "📥 발주자료 아카이브 다운로드",
            data=archive_bytes,
            file_name=f"★★발주자료{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_archive",
        )

    # ─────────────────────────────────────────────
    # Phase 2
    # ─────────────────────────────────────────────
    if "order_phase1_df" not in st.session_state:
        st.stop()

    st.divider()
    st.subheader("Phase 2 — 재고 대조 + 출력")

    if st.button("▶ Phase 2 실행", type="primary", key="p2_run"):
        with st.spinner("재고 대조 중..."):
            try:
                result2, unmatched_u, combined = lo.run_phase2(
                    st.session_state["order_phase1_df"]
                )
                if unmatched_u:
                    st.session_state["order_unmatched_units"]  = unmatched_u
                    st.session_state["order_combined_df"]       = combined
                    st.session_state.pop("order_phase2_df", None)
                    st.session_state.pop("order_stockout_df", None)
                else:
                    st.session_state["order_phase2_df"]   = result2
                    st.session_state["order_stockout_df"] = combined
                    _seed_stockout_board(combined)
                    st.session_state.pop("order_unmatched_units", None)
                    st.session_state.pop("order_combined_df", None)
            except Exception as e:
                st.error(f"오류: {e}"); st.exception(e)

    # ── GATE B : 낱개 원코드 미매칭 ─────────────
    if "order_unmatched_units" in st.session_state:
        unmatched_u = st.session_state["order_unmatched_units"]
        st.warning(f"⚠️ {len(unmatched_u)}개 낱개 코드의 원코드를 찾을 수 없습니다.")
        st.info("👉 **기준데이터관리 → 낱개처리목록**에서 원코드를 추가한 뒤 Phase 2를 재실행해주세요.")

        for row in unmatched_u:
            _a = row.get('어드민옵션', '')
            st.code(f"코드: {row['erp관리코드']}  |  {'' if pd.isna(_a) else _a}")

        if st.button("🔄 Phase 2 재실행", key="p2_retry"):
            with st.spinner("재실행 중..."):
                result2, unmatched_u2, combined2 = lo.run_phase2(
                    st.session_state["order_phase1_df"]
                )
                if unmatched_u2:
                    st.session_state["order_unmatched_units"] = unmatched_u2
                    st.session_state["order_combined_df"]      = combined2
                else:
                    st.session_state["order_phase2_df"]   = result2
                    st.session_state["order_stockout_df"] = combined2
                    _seed_stockout_board(combined2)
                    st.session_state.pop("order_unmatched_units", None)
                    st.rerun()

    # ── Phase 2 완료 표시 ───────────────────────
    if "order_phase2_df" in st.session_state and "order_stockout_df" in st.session_state:
        p2_df  = st.session_state["order_phase2_df"]
        so_df  = st.session_state["order_stockout_df"]

        c1, c2 = st.columns(2)
        c1.metric("물류팀 항목 수", len(p2_df))
        c2.metric("품절 품목", len(so_df))
        st.success("✅ Phase 2 완료")

        result_bytes = lo.generate_result_xlsx(p2_df, so_df)
        today = datetime.now(_KST).strftime("%m%d")
        st.download_button(
            "📥 최종결과물 다운로드 (물류팀 + 품절목록)",
            data=result_bytes,
            file_name=f"물류팀_{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_result",
        )


