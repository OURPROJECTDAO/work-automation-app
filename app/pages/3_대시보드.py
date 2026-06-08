"""대시보드 (Phase 4) — 매출 집계 + 데이터 추가 + 거래처 그룹 + 구분 분류.

탭:
 📊 대시보드   : 매출 KPI + 구분/그룹/거래처/상품/관리코드별 집계.
 ➕ 데이터 추가 : 영업이익현황 .xlsx → 날짜구간 교체로 월 파티션 누적.
 👥 거래처 그룹 : 상호명→그룹 배정(검색·인라인·일괄). private data repo 저장.
 🏷 구분 분류  : 미분류 관리코드 → 멸치쇼핑 분류표(app repo) 추가.
(차트·물류량·이익률은 추후 추가.)
"""
import base64
import io
import json
import sys
import unicodedata
import urllib.parse
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # repo root

import pandas as pd
import streamlit as st

from core.dashboard import store
from core.dashboard.sales_data import make_classifier, parse_sales

_REF = Path(__file__).parent.parent.parent / "reference"
_APP_REPO = "OURPROJECTDAO/work-automation-app"
_CLS_PATH = "reference/logistics_classification.csv"

st.title("📊 영업이익현황 대시보드")


def _nfc(s) -> str:
    return unicodedata.normalize("NFC", str(s)).strip()


def _data_secret() -> tuple[str, str]:
    """(pat, repo). secrets [data] 우선, 없으면 GITHUB_PAT 폴백. (private data repo)"""
    repo = "OURPROJECTDAO/work-automation-data"
    try:
        d = st.secrets["data"]
        return d["pat"], d.get("repo", repo)
    except Exception:
        return st.secrets.get("GITHUB_PAT", ""), repo


def _app_pat() -> str:
    """app repo(공개) 쓰기용 PAT — 발주 분류표 편집과 동일 GITHUB_PAT(없으면 data pat)."""
    pat = st.secrets.get("GITHUB_PAT", "")
    if pat:
        return pat
    try:
        return st.secrets["data"]["pat"]
    except Exception:
        return ""


@st.cache_data(ttl=3600, show_spinner="매출 데이터 불러오는 중...")
def load_sales(pat: str, repo: str) -> pd.DataFrame:
    df = store.load_master(pat, repo)
    if df.empty:
        return df
    cls = pd.read_csv(_REF / "logistics_classification.csv", dtype=str, encoding="utf-8-sig")
    pm = pd.read_csv(_REF / "product_master.csv", dtype=str, encoding="utf-8-sig")
    classify = make_classifier(cls, pm)
    df["구분"] = df["관리코드"].map(classify)
    df["연도"] = df["거래일자"].dt.year
    return df


@st.cache_data(ttl=3600, show_spinner=False)
def load_group_map(pat: str, repo: str) -> dict:
    """{NFC(상호명): 그룹}. 매칭은 NFC 정규화 키로."""
    g = store.read_groups(pat, repo)
    if g.empty:
        return {}
    out = {}
    for _, r in g.iterrows():
        nm, grp = r.get("상호명"), r.get("그룹")
        if pd.notna(nm) and pd.notna(grp) and str(grp).strip():
            out[_nfc(nm)] = str(grp).strip()
    return out


# ── app repo 분류표 R/W ──────────────────────────────────────────
def _gh_req(url, pat, accept="application/vnd.github+json", data=None, method="GET"):
    h = {"Authorization": "Bearer " + pat, "Accept": accept, "User-Agent": "wa-app"}
    if data is not None:
        h["Content-Type"] = "application/json"
    return urllib.request.Request(url, data=data, method=method, headers=h)


def _cls_url() -> str:
    return f"https://api.github.com/repos/{_APP_REPO}/contents/{urllib.parse.quote(_CLS_PATH)}"


def append_classification(pat: str, mapping: dict) -> int:
    """{관리코드: 구분} → logistics_classification.csv 갱신(기존 코드면 덮어, 신규면 추가)."""
    with urllib.request.urlopen(_gh_req(_cls_url(), pat, "application/vnd.github.raw")) as r:
        cur = pd.read_csv(io.BytesIO(r.read()), dtype=str, encoding="utf-8-sig")
    m = {}
    for _, row in cur.iterrows():
        m[_nfc(row["관리코드"])] = (str(row["관리코드"]), str(row["구분"]))
    for code, gub in mapping.items():
        m[_nfc(code)] = (str(code), gub)
    out = pd.DataFrame([(c, g) for (c, g) in m.values()], columns=["관리코드", "구분"])
    csv_bytes = out.to_csv(index=False).encode("utf-8-sig")
    with urllib.request.urlopen(_gh_req(_cls_url(), pat)) as r:
        sha = json.load(r)["sha"]
    body = {"message": f"dashboard: 구분 분류 {len(mapping)}건 추가", "branch": "main",
            "content": base64.b64encode(csv_bytes).decode(), "sha": sha}
    with urllib.request.urlopen(_gh_req(_cls_url(), pat, data=json.dumps(body).encode(), method="PUT")):
        pass
    return len(mapping)


def _save_groups(pat: str, repo: str, updates: dict, deletions=()) -> int:
    """기존 그룹 + updates({상호명:그룹}, 빈값=삭제) + deletions → write. canonical 상호명 보존."""
    base = store.read_groups(pat, repo)
    rec = {}  # NFC -> (상호명, 그룹)
    for _, r in base.iterrows():
        nm, g = str(r["상호명"]), str(r["그룹"]).strip()
        if g and g.lower() != "nan":
            rec[_nfc(nm)] = (nm, g)
    for nm, g in updates.items():
        k, g = _nfc(nm), str(g).strip()
        if g:
            rec[k] = (str(nm), g)
        elif k in rec:
            del rec[k]
    for nm in deletions:
        rec.pop(_nfc(nm), None)
    out = (pd.DataFrame([(nm, g) for (nm, g) in rec.values()], columns=["상호명", "그룹"])
           .sort_values(["그룹", "상호명"]).reset_index(drop=True))
    store.write_groups(pat, repo, out)
    return len(out)


pat, repo = _data_secret()
tab_dash, tab_add, tab_group, tab_cls = st.tabs(
    ["📊 대시보드", "➕ 데이터 추가", "👥 거래처 그룹", "🏷 구분 분류"])

# ── [데이터 추가] 탭 ───────────────────────────────────────────
with tab_add:
    st.subheader("영업이익현황 파일 추가")
    st.caption("일/주/월 다운로드한 영업이익현황(.xlsx)을 올리면 월별로 누적됩니다. "
               "같은 날짜구간을 다시 올리면 최신 파일로 교체됩니다(중복 안 쌓임).")
    if not pat:
        st.error("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    else:
        up = st.file_uploader("영업이익현황 .xlsx", type=["xlsx"], key="ingest_up")
        if up is not None:
            try:
                preview = parse_sales(io.BytesIO(up.getvalue()))
            except Exception as e:  # noqa: BLE001
                st.error(f"파일을 읽지 못했습니다: {e}")
                preview = None
            if preview is not None and len(preview):
                dmin, dmax = preview["거래일자"].min(), preview["거래일자"].max()
                months = sorted(preview["거래일자"].dt.strftime("%Y-%m").unique())
                st.info(f"행 **{len(preview):,}** · 기간 **{dmin:%Y-%m-%d} ~ {dmax:%Y-%m-%d}** · "
                        f"영향 월: {', '.join(months)}")
                st.caption("⚠ 위 날짜구간에 해당하는 기존 데이터는 이 파일 내용으로 교체됩니다.")
                if st.button("이 파일로 적재", type="primary", key="ingest_btn"):
                    with st.spinner("적재 중..."):
                        summary = store.ingest(pat, repo, io.BytesIO(up.getvalue()))
                    load_sales.clear()
                    st.success(f"적재 완료 · {summary['rows']:,}행 · 월 {len(summary['months'])}개 갱신 "
                               f"({summary['date_range'][0]} ~ {summary['date_range'][1]})")
                    st.rerun()
            elif preview is not None:
                st.warning("유효한 거래 데이터가 없습니다(합계행만 있거나 빈 파일).")

# ── [거래처 그룹] 탭 ───────────────────────────────────────────
with tab_group:
    st.subheader("거래처 그룹 관리")
    if not pat:
        st.error("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    else:
        df_g = load_sales(pat, repo)
        gmap = load_group_map(pat, repo)
        if df_g.empty:
            st.info("매출 데이터가 없습니다.")
        else:
            sales_by_store = df_g.groupby("상호명", observed=True)["판매금액"].sum()
            stores = sorted(sales_by_store.index, key=lambda s: -sales_by_store[s])
            canon = {_nfc(s): s for s in stores}
            assigned = {s: gmap.get(_nfc(s), "") for s in stores}
            cnt = pd.Series([g or "(미지정)" for g in assigned.values()]).value_counts()
            st.caption("현황 · " + " / ".join(f"{g} {c:,}" for g, c in cnt.items()))

            st.markdown("**① 검색 → 인라인 지정** (그룹칸을 비우면 해제)")
            q = st.text_input("거래처 검색", key="grp_q", placeholder="상호명 일부")
            if q.strip():
                shown = [s for s in stores if q.strip() in s]
            else:
                shown = stores[:50]
                st.caption("검색어 없음 — 매출 상위 50곳만 표시")
            tbl = pd.DataFrame({
                "상호명": shown,
                "매출": [int(sales_by_store[s]) for s in shown],
                "그룹": [assigned[s] for s in shown],
            })
            edited = st.data_editor(
                tbl, key="grp_editor", use_container_width=True, hide_index=True, num_rows="fixed",
                column_config={
                    "상호명": st.column_config.TextColumn(disabled=True),
                    "매출": st.column_config.NumberColumn(format="%d", disabled=True),
                    "그룹": st.column_config.TextColumn(help="온라인 / 오프라인 등 자유 입력"),
                },
                height=min(560, 80 + 36 * min(len(shown), 30)))
            if st.button("💾 지정 저장", key="grp_save_inline", type="primary"):
                updates = {r["상호명"]: str(r["그룹"]).strip() for _, r in edited.iterrows()}
                n = _save_groups(pat, repo, updates)
                load_group_map.clear()
                st.success(f"저장 완료 · 그룹 배정 {n:,}곳")
                st.rerun()

            st.divider()
            st.markdown("**② 목록 일괄 배정** (붙여넣은 거래처를 한 그룹으로)")
            bulk_text = st.text_area("거래처 목록 (한 줄에 1개)", key="grp_bulk_text", height=120)
            bg = st.text_input("배정할 그룹명", value="온라인", key="grp_bulk_group")
            if st.button("일괄 배정 저장", key="grp_bulk_btn"):
                lines = [ln.strip() for ln in bulk_text.splitlines() if ln.strip()]
                matched, unmatched = {}, []
                for ln in lines:
                    k = _nfc(ln)
                    if k in canon:
                        matched[canon[k]] = bg.strip()
                    else:
                        unmatched.append(ln)
                if matched:
                    n = _save_groups(pat, repo, matched)
                    load_group_map.clear()
                    st.success(f"{len(matched)}곳 '{bg.strip()}' 배정 저장 · 전체 {n:,}곳")
                else:
                    st.warning("매칭된 거래처가 없습니다.")
                if unmatched:
                    st.warning("매출 데이터에서 못 찾은 거래처(저장 안 됨): " + ", ".join(unmatched))

# ── [구분 분류] 탭 ─────────────────────────────────────────────
with tab_cls:
    st.subheader("미분류 구분 → 멸치쇼핑 분류표 추가")
    st.caption("미분류 관리코드를 음료/식품/선물세트로 지정하면 발주서출력업무와 공유하는 분류표에 반영됩니다. "
               "저장 후 재배포(1~2분) 뒤 대시보드에 적용됩니다.")
    apat = _app_pat()
    if not pat:
        st.error("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    elif not apat:
        st.error("분류표 쓰기용 `GITHUB_PAT` 시크릿이 없습니다.")
    else:
        df_c = load_sales(pat, repo)
        unc = df_c[df_c["구분"] == "미분류"] if not df_c.empty else df_c
        if df_c.empty:
            st.info("매출 데이터가 없습니다.")
        elif unc.empty:
            st.success("미분류 관리코드가 없습니다 🎉")
        else:
            g = (unc.groupby("관리코드", observed=True)
                 .agg(상품명=("상품명", "first"), 매출=("판매금액", "sum"), 건수=("판매금액", "size"))
                 .sort_values("매출", ascending=False).reset_index())
            st.caption(f"미분류 관리코드 **{len(g):,}개** · 매출합 **{g['매출'].sum()/1e8:.1f}억**")
            qc = st.text_input("관리코드·상품명 검색", key="cls_q")
            if qc.strip():
                mask = (g["관리코드"].str.contains(qc.strip(), na=False)
                        | g["상품명"].astype(str).str.contains(qc.strip(), na=False))
                shown = g[mask].copy()
            else:
                shown = g.head(50).copy()
                st.caption("검색어 없음 — 매출 상위 50개만 표시")
            shown["매출"] = shown["매출"].astype(int)
            shown["구분지정"] = ""
            edited = st.data_editor(
                shown, key="cls_editor", use_container_width=True, hide_index=True, num_rows="fixed",
                column_config={
                    "관리코드": st.column_config.TextColumn(disabled=True),
                    "상품명": st.column_config.TextColumn(disabled=True),
                    "매출": st.column_config.NumberColumn(format="%d", disabled=True),
                    "건수": st.column_config.NumberColumn(disabled=True),
                    "구분지정": st.column_config.SelectboxColumn(
                        options=["", "음료", "식품", "선물세트"], help="지정한 행만 분류표에 반영"),
                },
                height=min(560, 80 + 36 * min(len(shown), 30)))
            picks = {str(r["관리코드"]): str(r["구분지정"]).strip()
                     for _, r in edited.iterrows() if str(r["구분지정"]).strip()}
            st.caption(f"지정됨: {len(picks)}건")
            if st.button("💾 분류표에 추가", key="cls_save", type="primary", disabled=not picks):
                with st.spinner("분류표 갱신 중..."):
                    n = append_classification(apat, picks)
                load_sales.clear()
                st.success(f"{n}건 분류표에 반영. 재배포(1~2분) 후 대시보드에 적용됩니다.")

# ── [대시보드] 탭 ──────────────────────────────────────────────
with tab_dash:
    if not pat:
        st.warning("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    else:
        df = load_sales(pat, repo)
        if df.empty:
            st.info("적재된 매출 데이터가 없습니다. [➕ 데이터 추가] 탭에서 파일을 올려주세요.")
        else:
            gmap = load_group_map(pat, repo)
            years = sorted(df["연도"].unique())
            gubuns = ["음료", "식품", "선물세트", "미분류"]
            group_opts = sorted(set(gmap.values())) + ["(미지정)"]
            c1, c2, c3 = st.columns(3)
            with c1:
                sel_years = st.multiselect("연도", years, default=years)
            with c2:
                sel_gubun = st.multiselect("구분", gubuns, default=gubuns)
            with c3:
                sel_group = st.multiselect("그룹", group_opts, default=group_opts)

            view = df[df["연도"].isin(sel_years) & df["구분"].isin(sel_gubun)].copy()
            view["그룹"] = view["상호명"].map(lambda s: gmap.get(_nfc(s), "(미지정)"))
            view = view[view["그룹"].isin(sel_group)]
            if view.empty:
                st.info("선택한 조건에 해당하는 데이터가 없습니다.")
            else:
                total = view["판매금액"].sum()
                k1, k2, k3 = st.columns(3)
                k1.metric("총 매출", f"{total/1e8:,.1f}억")
                k2.metric("거래 건수", f"{len(view):,}건")
                k3.metric("기간", f"{view['거래일자'].min():%Y-%m} ~ {view['거래일자'].max():%Y-%m}")
                st.divider()

                DIMS = {"구분": "구분", "그룹": "그룹", "거래처": "상호명",
                        "상품": "상품명", "관리코드": "관리코드"}
                dim_label = st.selectbox("집계 기준", list(DIMS), index=0)
                dim = DIMS[dim_label]
                agg = (view.groupby(dim, observed=True)["판매금액"].sum()
                       .sort_values(ascending=False).reset_index())
                agg.columns = [dim_label, "매출"]
                agg["비중(%)"] = (agg["매출"] / total * 100).round(1)
                agg.insert(0, "순위", range(1, len(agg) + 1))

                st.subheader(f"{dim_label}별 매출 — 총 {len(agg):,}개")
                disp = agg.copy()
                disp["매출"] = disp["매출"].map(lambda v: f"{v:,.0f}")
                st.dataframe(disp, use_container_width=True, hide_index=True,
                             height=min(560, 80 + 36 * min(len(disp), 30)))
                st.download_button(
                    "표 CSV 내려받기",
                    agg.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"매출_{dim_label}별.csv",
                    mime="text/csv",
                )
