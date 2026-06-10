"""대시보드 (Phase 4) — 매출 집계 + 데이터 추가 + 거래처 그룹 + 구분 분류.

탭:
 📊 대시보드   : 매출 KPI + 기간(날짜범위)·일/월/연 추이·구분/그룹/거래처/상품/관리코드 집계 + 그룹 내 거래처 체크박스.
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
from core.dashboard.sales_data import make_attr_lookup, make_box_lookup, make_classifier, parse_sales

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
    attr = pd.read_csv(_REF / "product_attributes.csv", dtype=str, encoding="utf-8-sig")
    amaps = make_attr_lookup(attr)
    classify = make_classifier(cls, pm, amaps["식품음료"])  # 식품음료 = 구분 3차 fallback
    df["구분"] = df["관리코드"].map(classify)
    m = amaps["최종분류"]
    df["최종분류"] = df["관리코드"].map(lambda x, mm=m: mm.get(_nfc(x), "미지정")).astype("category")
    # 온라인 상품마진용: 합포수량(결측 NaN) + 박스내품(결측/0→1.0)
    hap = {_nfc(k): v for k, v in zip(attr["관리코드"], attr["합포수량"])}

    def _hapq(c):
        try:
            f = float(str(hap.get(_nfc(c), "")).strip())
            return f if f > 0 else float("nan")
        except (TypeError, ValueError):
            return float("nan")

    boxq = make_box_lookup(pm)
    df["합포수량"] = df["관리코드"].map(_hapq)
    df["박스내품"] = df["관리코드"].map(boxq)
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
tab_dash, tab_add, tab_group, tab_cls, tab_margin = st.tabs(
    ["📊 대시보드", "➕ 데이터 추가", "👥 거래처 그룹", "🏷 구분 분류", "💰 상품마진(온라인)"])

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
def _won(v) -> str:
    v = float(v)
    if abs(v) >= 1e8:
        return f"{v/1e8:,.2f}억"
    if abs(v) >= 1e4:
        return f"{v/1e4:,.0f}만"
    return f"{v:,.0f}"



def _dim_key(view, label):
    """집계 기준 라벨 → (key Series, is_time)."""
    if label == "일별":
        return view["거래일자"].dt.strftime("%Y-%m-%d"), True
    if label == "월별":
        return view["거래일자"].dt.strftime("%Y-%m"), True
    if label == "연별":
        return view["거래일자"].dt.year.astype(str), True
    col = {"구분": "구분", "그룹": "그룹", "거래처": "상호명",
           "상품": "상품명", "관리코드": "관리코드", "세분류": "최종분류"}[label]
    return view[col].astype(str), False


def _pivot_table(view, box, d1, d2, mode, unit):
    """행=d1, 열=d2 교차표. 셀=매출(매출 모드) / 이익(이익 모드). 반환 (표시DF, 부제, 전체행수)."""
    k1, _ = _dim_key(view, d1)
    k2, t2 = _dim_key(view, d2)
    v = view.assign(_r=k1, _c=k2)
    if mode == "매출":
        v = v[~box]
        p = (v.groupby(["_r", "_c"], observed=True)["판매금액"].sum()
             .unstack("_c", fill_value=0.0))
    else:
        v["_pi"] = v["판매이익"].where(~box, 0.0)
        v["_cnt"] = v["수량"].where(box, 0.0)
        gg = v.groupby(["_r", "_c"], observed=True).agg(pi=("_pi", "sum"), cnt=("_cnt", "sum"))
        gg["val"] = gg["pi"] - gg["cnt"] * unit
        p = gg["val"].unstack("_c", fill_value=0.0)
    note = ""
    if t2:
        cols = sorted(p.columns)
        if len(cols) > 60:
            note += f" · 열 {len(cols)}개(가로 스크롤)"
    else:
        cols = list(p.sum(axis=0).sort_values(ascending=False).index)
        if len(cols) > 30:
            note += f" · 열 상위 30개(전체 {len(cols)})"
            cols = cols[:30]
    p = p[cols]
    p["합계"] = p.sum(axis=1)
    grand = p.sum(axis=0)
    p = p.sort_values("합계", ascending=False)
    n = len(p)
    if n > 100:
        p = p.head(100); note += f" · 행 상위 100개(전체 {n})"
    p = p.round().astype("int64")
    out = p.reset_index().rename(columns={"_r": d1})
    gt = {d1: "합계"}
    for c in list(cols) + ["합계"]:
        gt[c] = int(round(grand[c]))
    out = pd.concat([out, pd.DataFrame([gt])], ignore_index=True)
    return out, note, n


def _render_dashboard(pat: str, repo: str) -> None:
    df = load_sales(pat, repo)
    if df.empty:
        st.info("적재된 매출 데이터가 없습니다. [➕ 데이터 추가] 탭에서 파일을 올려주세요.")
        return
    gmap = load_group_map(pat, repo)

    _pref = ["음료", "식품", "선물세트", "미분류"]
    _present = list(df["구분"].dropna().unique())
    gubuns = ([g for g in _pref if g in _present]
              + sorted(g for g in _present if g not in _pref))
    group_opts = sorted(set(gmap.values())) + ["(미지정)"]
    dmin, dmax = df["거래일자"].min().date(), df["거래일자"].max().date()

    metric = st.radio("지표", ["매출", "이익"], horizontal=True, key="dash_metric")
    is_profit = metric == "이익"

    c1, c2, c3 = st.columns(3)
    with c1:
        dr = st.date_input("기간", value=(dmin, dmax),
                           min_value=dmin, max_value=dmax, format="YYYY-MM-DD")
    with c2:
        if is_profit:
            st.caption("이익은 **전체 구분** 기준 — 택배비는 상품 구분에 배분되지 않습니다.")
            sel_gubun = gubuns
        else:
            sel_gubun = st.multiselect("구분", gubuns, default=gubuns)
    with c3:
        sel_group = st.multiselect("그룹", group_opts, default=group_opts)

    if isinstance(dr, (list, tuple)):
        d_start, d_end = (dr[0], dr[-1]) if dr else (dmin, dmax)
    else:
        d_start = d_end = dr
    ts0, ts1 = pd.Timestamp(d_start), pd.Timestamp(d_end) + pd.Timedelta(days=1)

    view = df[(df["거래일자"] >= ts0) & (df["거래일자"] < ts1)].copy()
    if not is_profit:
        view = view[view["구분"].isin(sel_gubun)]
    view["그룹"] = view["상호명"].map(lambda s: gmap.get(_nfc(s), "(미지정)"))
    view = view[view["그룹"].isin(sel_group)]
    if view.empty:
        st.info("선택한 조건에 해당하는 데이터가 없습니다.")
        return

    # ── 그룹 내 거래처 선택 (체크박스, 매출 기준 멤버) ──────────
    _box0 = view["관리코드"].astype(str) == "00-12"
    SMALL = 50
    store_sales = view.loc[~_box0].groupby("상호명", observed=True)["판매금액"].sum()
    members_by_grp: dict = {}
    for store, sales in store_sales.items():
        g = gmap.get(_nfc(store), "(미지정)")
        members_by_grp.setdefault(g, []).append((store, int(sales)))
    for g in members_by_grp:
        members_by_grp[g].sort(key=lambda t: -t[1])
    small_rows, big_groups = [], {}
    for g, members in members_by_grp.items():
        if len(members) <= SMALL:
            small_rows += [{"포함": True, "그룹": g, "거래처": s, "매출": v} for s, v in members]
        else:
            big_groups[g] = members
    excluded: set = set()
    if small_rows or big_groups:
        with st.expander("그룹 내 거래처 선택 (체크 해제 = 제외)", expanded=bool(small_rows)):
            if small_rows:
                if "dash_store_bulk" not in st.session_state:
                    st.session_state["dash_store_bulk"] = True
                if "dash_store_ver" not in st.session_state:
                    st.session_state["dash_store_ver"] = 0
                _bulk = st.session_state["dash_store_bulk"]
                if st.button("전체 해제" if _bulk else "전체 선택", key="dash_store_all"):
                    st.session_state["dash_store_bulk"] = not _bulk
                    st.session_state["dash_store_ver"] += 1
                    st.rerun()
                _picks = pd.DataFrame(small_rows)
                _picks["포함"] = st.session_state["dash_store_bulk"]
                ed = st.data_editor(
                    _picks, key=f"dash_store_pick_{st.session_state['dash_store_ver']}",
                    use_container_width=True, hide_index=True, num_rows="fixed",
                    column_config={
                        "포함": st.column_config.CheckboxColumn(default=True),
                        "그룹": st.column_config.TextColumn(disabled=True),
                        "거래처": st.column_config.TextColumn(disabled=True),
                        "매출": st.column_config.NumberColumn(format="accounting", disabled=True),
                    },
                    height=min(560, 80 + 36 * min(len(small_rows), 30)))
                excluded |= {r["거래처"] for _, r in ed.iterrows() if not r["포함"]}
            for g, members in big_groups.items():
                opts = [s for s, _ in members]
                ex = st.multiselect(f"제외할 거래처 — {g} ({len(opts):,}곳)", opts, key=f"dash_excl_{g}")
                excluded |= set(ex)
    if excluded:
        view = view[~view["상호명"].isin(excluded)]
        if view.empty:
            st.info("거래처를 모두 제외하여 표시할 데이터가 없습니다.")
            return

    box = view["관리코드"].astype(str) == "00-12"  # 택배비 라인(C타입 택배비)

    # ── 매출 모드 (택배비 라인 00-12 제외) ─────────────────────
    if not is_profit:
        vs = view[~box]
        if vs.empty:
            st.info("선택한 조건에 해당하는 매출 데이터가 없습니다.")
            return
        total = vs["판매금액"].sum()
        k1, k2, k3 = st.columns(3)
        k1.metric("총 매출", _won(total))
        k2.metric("거래 건수", f"{len(vs):,}건")
        k3.metric("기간", f"{vs['거래일자'].min():%Y-%m-%d} ~ {vs['거래일자'].max():%Y-%m-%d}")
        st.divider()

        TIME_DIMS = ["일별", "월별", "연별"]
        CAT_DIMS = {"구분": "구분", "그룹": "그룹", "거래처": "상호명",
                    "상품": "상품명", "관리코드": "관리코드", "세분류": "최종분류"}
        ALL_DIMS = TIME_DIMS + list(CAT_DIMS)
        cc = st.columns(2)
        with cc[0]:
            d1 = st.selectbox("집계 기준 (행)", ALL_DIMS, index=1, key="sales_d1")
        with cc[1]:
            d2 = st.selectbox("× 기준 2 (열, 선택)", ["(없음)"] + [d for d in ALL_DIMS if d != d1],
                              key="sales_d2")
        if d2 != "(없음)":
            out, note, n = _pivot_table(view, box, d1, d2, "매출", None)
            st.subheader(f"{d1} × {d2} 매출 — {min(n, 100):,}행{note}")
            cfg = {c: st.column_config.NumberColumn(format="localized")
                   for c in out.columns if c != d1}
            st.dataframe(out, use_container_width=True, hide_index=True, column_config=cfg,
                         height=min(620, 80 + 36 * min(len(out), 28)))
            st.download_button("표 CSV 내려받기", out.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"매출_{d1}_x_{d2}.csv", mime="text/csv", key="dl_pivot_sales")
            return
        dim_label = d1
        if dim_label in TIME_DIMS:
            fmt = {"일별": "%Y-%m-%d", "월별": "%Y-%m"}.get(dim_label)
            key = vs["거래일자"].dt.strftime(fmt) if fmt else vs["거래일자"].dt.year.astype(str)
            agg = (vs.assign(_k=key).groupby("_k", observed=True)["판매금액"].sum()
                   .sort_index().reset_index())
            agg.columns = [dim_label, "매출"]
            agg["비중(%)"] = (agg["매출"] / total * 100).round(1)
            st.subheader(f"{dim_label} 매출 추이 — {len(agg):,}개 구간")
            chart = agg.copy()
            if dim_label == "일별":
                chart.index = pd.to_datetime(chart[dim_label])
            elif dim_label == "월별":
                chart.index = pd.to_datetime(chart[dim_label] + "-01")
            else:
                chart.index = chart[dim_label]
            st.line_chart(chart["매출"], height=260)
            agg["매출"] = agg["매출"].round().astype("int64")
            st.dataframe(agg, use_container_width=True, hide_index=True,
                         height=min(420, 80 + 36 * min(len(agg), 20)),
                         column_config={"매출": st.column_config.NumberColumn(format="localized")})
            fname, dlkey = f"매출_{dim_label}.csv", "dl_time"
        else:
            dim = CAT_DIMS[dim_label]
            agg = (vs.groupby(dim, observed=True)["판매금액"].sum()
                   .sort_values(ascending=False).reset_index())
            agg.columns = [dim_label, "매출"]
            agg["비중(%)"] = (agg["매출"] / total * 100).round(1)
            agg.insert(0, "순위", range(1, len(agg) + 1))
            st.subheader(f"{dim_label}별 매출 — 총 {len(agg):,}개")
            agg["매출"] = agg["매출"].round().astype("int64")
            st.dataframe(agg, use_container_width=True, hide_index=True,
                         height=min(560, 80 + 36 * min(len(agg), 30)),
                         column_config={"매출": st.column_config.NumberColumn(format="localized")})
            fname, dlkey = f"매출_{dim_label}별.csv", "dl_cat"
        st.download_button("표 CSV 내려받기", agg.to_csv(index=False).encode("utf-8-sig"),
                           file_name=fname, mime="text/csv", key=dlkey)
        return

    # ── 이익 모드 (00-12 = 택배비, 전체 구분) ──────────────────
    fee_label = st.radio("택배비 단가 (송장 1건당)",
                         ["3,000원 (ERP 입력값)", "2,500원 (보정값)"],
                         horizontal=True, key="dash_fee")
    unit = 2500 if "2,500" in fee_label else 3000
    suf = " (보정)" if unit == 2500 else ""

    매출 = view.loc[~box, "판매금액"].sum()
    상품이익 = view.loc[~box, "판매이익"].sum()
    송장 = view.loc[box, "수량"].sum()
    택배비 = 송장 * unit
    매입가 = 매출 - 상품이익
    이익 = 상품이익 - 택배비
    률 = (이익 / 매입가 * 100) if 매입가 else 0.0

    r1 = st.columns(3)
    r1[0].metric("매출 (수수료 차감 후)", _won(매출))
    r1[1].metric("매입가", _won(매입가))
    r1[2].metric("송장 건수", f"{송장:,.0f}건")
    r2 = st.columns(3)
    r2[0].metric("택배비" + suf, _won(택배비))
    r2[1].metric("이익" + suf, _won(이익))
    r2[2].metric("이익률 (이익÷매입가)" + suf, f"{률:.2f}%")
    if unit == 2500:
        st.caption("⚠ **보정값** — ERP 입력 택배비는 3,000원이나 2,500원으로 재계산한 값입니다.")
    st.divider()

    DIMS = {"일별": "거래일자", "월별": "거래일자", "연별": "거래일자",
            "거래처": "상호명", "그룹": "그룹"}
    cc = st.columns(2)
    with cc[0]:
        d1 = st.selectbox("집계 기준 (행)", list(DIMS), index=1, key="profit_d1")
    with cc[1]:
        d2 = st.selectbox("× 기준 2 (열, 선택)", ["(없음)"] + [d for d in DIMS if d != d1],
                          key="profit_d2")
    if d2 != "(없음)":
        out, note, n = _pivot_table(view, box, d1, d2, "이익", unit)
        st.subheader(f"{d1} × {d2} 이익{suf} — {min(n, 100):,}행{note}")
        cfg = {c: st.column_config.NumberColumn(format="localized")
               for c in out.columns if c != d1}
        st.dataframe(out, use_container_width=True, hide_index=True, column_config=cfg,
                     height=min(620, 80 + 36 * min(len(out), 28)))
        st.download_button("표 CSV 내려받기", out.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"이익_{d1}_x_{d2}{'_보정' if unit == 2500 else ''}.csv",
                           mime="text/csv", key="dl_pivot_profit")
        return
    dim_label = d1
    if dim_label == "일별":
        key = view["거래일자"].dt.strftime("%Y-%m-%d")
    elif dim_label == "월별":
        key = view["거래일자"].dt.strftime("%Y-%m")
    elif dim_label == "연별":
        key = view["거래일자"].dt.year.astype(str)
    else:
        key = view[DIMS[dim_label]]

    v2 = view.assign(_k=key)
    v2["_매출"] = v2["판매금액"].where(~box, 0.0)
    v2["_상품이익"] = v2["판매이익"].where(~box, 0.0)
    v2["_건수"] = v2["수량"].where(box, 0.0)
    g = (v2.groupby("_k", observed=True)
         .agg(매출=("_매출", "sum"), 상품이익=("_상품이익", "sum"), 송장=("_건수", "sum"))
         .reset_index())
    g["택배비"] = g["송장"] * unit
    g["매입가"] = g["매출"] - g["상품이익"]
    g["이익"] = g["상품이익"] - g["택배비"]
    g["이익률(%)"] = (g["이익"] / g["매입가"].replace(0, pd.NA) * 100).round(2)
    is_time = dim_label in ("일별", "월별", "연별")
    g = (g.sort_values("_k") if is_time else g.sort_values("이익", ascending=False))
    g = g.rename(columns={"_k": dim_label})

    st.subheader(f"{dim_label} 이익{suf} — {len(g):,}개 " + ("구간" if is_time else "항목"))
    if is_time:
        chart = g.copy()
        if dim_label == "일별":
            chart.index = pd.to_datetime(chart[dim_label])
        elif dim_label == "월별":
            chart.index = pd.to_datetime(chart[dim_label] + "-01")
        else:
            chart.index = chart[dim_label]
        st.line_chart(chart["이익"], height=260)

    show = g[[dim_label, "매출", "매입가", "택배비", "이익", "이익률(%)", "송장"]].copy()
    for c in ["매출", "매입가", "택배비", "이익", "송장"]:
        show[c] = show[c].round().astype("int64")
    _num = st.column_config.NumberColumn(format="localized")
    st.dataframe(show, use_container_width=True, hide_index=True,
                 height=min(520, 80 + 36 * min(len(show), 24)),
                 column_config={"매출": _num, "매입가": _num, "택배비": _num,
                                "이익": _num, "송장": _num,
                                "이익률(%)": st.column_config.NumberColumn(format="%.2f")})
    st.download_button("표 CSV 내려받기", show.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"이익_{dim_label}{'_보정' if unit == 2500 else ''}.csv",
                       mime="text/csv", key="dl_profit")


def _render_online_margin(pat: str, repo: str) -> None:
    df = load_sales(pat, repo)
    if df.empty:
        st.info("적재된 매출 데이터가 없습니다. [➕ 데이터 추가] 탭에서 파일을 올려주세요.")
        return
    st.caption("**온라인(택배비 발생) 거래처 한정 · 상품별 추정 마진율.** "
               "택배비를 `실택배비 × 수량 ÷ (합포수량×내품수)`로 상품에 배분 추정하고, "
               "채널 보정계수(실제송장÷추정송장)로 실제 택배비 총액에 맞춥니다. 절대값보단 상품 간 비교용.")

    online = sorted(set(df.loc[df["관리코드"].astype(str) == "00-12", "상호명"].astype(str)))
    if not online:
        st.info("택배비(00-12) 행이 있는 온라인 거래처가 없습니다.")
        return

    dmin, dmax = df["거래일자"].min().date(), df["거래일자"].max().date()
    c1, c2, c3 = st.columns(3)
    with c1:
        dr = st.date_input("기간", value=(dmin, dmax), min_value=dmin, max_value=dmax,
                           format="YYYY-MM-DD", key="om_date")
    with c2:
        sel_store = st.multiselect("온라인 거래처", online, default=online, key="om_store")
    with c3:
        fee_label = st.radio("택배비 단가", ["3,000원", "2,500원"], horizontal=True, key="om_fee")
    unit = 2500 if "2,500" in fee_label else 3000
    cc = st.columns(2)
    with cc[0]:
        corr = st.toggle("채널 보정계수 적용 (권장)", value=True, key="om_corr")
    with cc[1]:
        dim_label = st.selectbox("상품 기준", ["관리코드", "상품명", "세분류"], key="om_dim")
    dim_col = {"관리코드": "관리코드", "상품명": "상품명", "세분류": "최종분류"}[dim_label]

    if isinstance(dr, (list, tuple)):
        d_start, d_end = (dr[0], dr[-1]) if dr else (dmin, dmax)
    else:
        d_start = d_end = dr
    ts0, ts1 = pd.Timestamp(d_start), pd.Timestamp(d_end) + pd.Timedelta(days=1)
    view = df[(df["거래일자"] >= ts0) & (df["거래일자"] < ts1)
              & (df["상호명"].astype(str).isin(sel_store))].copy()
    if view.empty:
        st.info("선택한 조건에 해당하는 데이터가 없습니다.")
        return

    box = view["관리코드"].astype(str) == "00-12"
    prod = view[~box].copy()
    hap = prod["합포수량"].fillna(1.0)
    hap = hap.where(hap > 0, 1.0)
    boxn = prod["박스내품"].where(prod["박스내품"] > 0, 1.0)
    prod["_송장"] = prod["수량"] / (hap * boxn)
    추정송장 = prod["_송장"].sum()
    실제송장 = view.loc[box, "수량"].sum()
    k = (실제송장 / 추정송장) if (corr and 추정송장) else 1.0
    prod["_택배"] = prod["_송장"] * unit * k

    g = (prod.assign(_d=prod[dim_col].astype(str))
         .groupby("_d", observed=True)
         .agg(매출=("판매금액", "sum"), 판매이익=("판매이익", "sum"),
              추정택배=("_택배", "sum"), 수량=("수량", "sum"))
         .reset_index())
    g["매입가"] = g["매출"] - g["판매이익"]
    g["순이익"] = g["판매이익"] - g["추정택배"]
    g["마진율(%)"] = (g["순이익"] / g["매입가"].replace(0, pd.NA) * 100).round(2)
    g = g.sort_values("매출", ascending=False)

    t매출, t매입, t순 = g["매출"].sum(), g["매입가"].sum(), g["순이익"].sum()
    t률 = (t순 / t매입 * 100) if t매입 else 0.0
    r = st.columns(4)
    r[0].metric("매출 (수수료 차감 후)", _won(t매출))
    r[1].metric("매입가", _won(t매입))
    r[2].metric("순이익 (추정)", _won(t순))
    r[3].metric("마진율 (추정)", f"{t률:.2f}%")
    kv = (실제송장 / 추정송장) if 추정송장 else 0.0
    st.caption(f"보정계수 k = 실제송장 {실제송장:,.0f} ÷ 추정송장 {추정송장:,.0f} = **{kv:.3f}** "
               + ("→ 적용됨" if corr else "→ 미적용(낙관 추정)"))

    n = len(g)
    if n > 200:
        g = g.head(200)
        st.caption(f"매출 상위 200개 표시 (전체 {n}개)")
    out = g.rename(columns={"_d": dim_label})[
        [dim_label, "매출", "매입가", "추정택배", "순이익", "마진율(%)", "수량"]].copy()
    for c in ["매출", "매입가", "추정택배", "순이익"]:
        out[c] = out[c].round().astype("int64")
    cfg = {c: st.column_config.NumberColumn(format="localized")
           for c in ["매출", "매입가", "추정택배", "순이익", "수량"]}
    st.dataframe(out, use_container_width=True, hide_index=True, column_config=cfg,
                 height=min(620, 80 + 36 * min(len(out), 16)))
    st.download_button("표 CSV 내려받기", out.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"온라인_상품마진_{dim_label}.csv", mime="text/csv", key="om_dl")


with tab_dash:
    if not pat:
        st.warning("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    else:
        _render_dashboard(pat, repo)

with tab_margin:
    if not pat:
        st.warning("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    else:
        _render_online_margin(pat, repo)
