"""시장가 매칭 — nadl 행사 개당가 ↔ 우리 관리코드.

시장지능(ADR 0025). 로컬 수집기가 적립한 nadl 행사 개당가(work-automation-data:market/nadl)를
우리 product_master에 매칭. 매칭 모델 v2(박스 하드게이트+bigram+브랜드)가 top3 후보를 제시하고,
**사람이 top3에서 선택 또는 '없음'으로 확정**. 확정 매핑은 ps_goid 키로 영속(nadl_map.csv) →
다음 주 수집에도 유지. 실사용하며 모델 정교화.

탭: ① 조회(전체 nadl) ② 매칭(미매칭 큐 + top3 선택) ③ 매칭본(우리 컬럼 결합 + 다운로드).
page-only 오케스트레이션 — 매칭 로직 core = core/intelligence/market_nadl.py.
"""
import io
import sys
import unicodedata
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import pandas as pd
import streamlit as st

from core.intelligence import market_nadl as M

_REF = Path(__file__).parent.parent.parent / "reference"

st.title("🛒 시장가 매칭")
st.caption("nadl 행사 **개당가**를 우리 관리코드에 매칭합니다. 모델이 후보 3개를 제시하면 "
           "골라서 확정하세요(없으면 '없음'). 확정 매핑은 ps_goid로 저장돼 다음 수집에도 유지됩니다.")


def _nfc(s) -> str:
    return unicodedata.normalize("NFC", str(s)).strip()


def _won(v) -> str:
    try:
        return f"₩{int(round(float(v))):,}"
    except (TypeError, ValueError):
        return "—"


def _data_secret() -> tuple:
    repo = "OURPROJECTDAO/work-automation-data"
    try:
        d = st.secrets["data"]
        return d["pat"], d.get("repo", repo)
    except Exception:
        return st.secrets.get("GITHUB_PAT", ""), repo


@st.cache_data(ttl=3600, show_spinner="product_master 로드 중...")
def load_pm() -> pd.DataFrame:
    return pd.read_csv(_REF / "product_master.csv", dtype=str, encoding="utf-8-sig")


@st.cache_data(ttl=3600, show_spinner="매칭 인덱스 준비 중...")
def pm_index(pm: pd.DataFrame) -> dict:
    return M.build_pm_index(pm)


@st.cache_data(ttl=3600, show_spinner="nadl 가격 불러오는 중...")
def load_prices(pat: str, repo: str, dt: str) -> pd.DataFrame:
    return M.load_prices(pat, repo, dt if dt else None)


@st.cache_data(ttl=600, show_spinner="매핑 불러오는 중...")
def load_map(pat: str, repo: str, version: int) -> pd.DataFrame:
    return M.read_map(pat, repo)


def _bump_map():
    st.session_state["map_ver"] = st.session_state.get("map_ver", 0) + 1


pat, repo = _data_secret()
if not pat:
    st.error("데이터 저장소 토큰(st.secrets['data']['pat'])이 설정되지 않았습니다.")
    st.stop()

st.session_state.setdefault("map_ver", 0)

# 날짜 선택
dates = M.list_price_dates(pat, repo)
if not dates:
    st.warning("아직 수집된 nadl 가격 데이터가 없습니다. 로컬 수집기(run_nadl.bat)를 먼저 실행하세요.")
    st.stop()
sel_date = st.selectbox("수집일", dates[::-1], index=0)

prices = load_prices(pat, repo, sel_date)
pm = load_pm()
idx = pm_index(pm)
mp = load_map(pat, repo, st.session_state["map_ver"])

# 매핑 상태 분류
goid_status = {}  # ps_goid -> 'matched' | 'none'
goid_codes = {}   # ps_goid -> [관리코드]
for _, r in mp.iterrows():
    g = _nfc(r["ps_goid"])
    if r["status"] == "matched":
        goid_status[g] = "matched"
        goid_codes.setdefault(g, []).append(_nfc(r["관리코드"]))
    elif r["status"] == "none" and goid_status.get(g) != "matched":
        goid_status[g] = "none"

all_goids = [_nfc(g) for g in prices["ps_goid"]]
n_total = len(prices)
n_matched = sum(1 for g in all_goids if goid_status.get(g) == "matched")
n_none = sum(1 for g in all_goids if goid_status.get(g) == "none")
n_todo = n_total - n_matched - n_none

tab1, tab2, tab3 = st.tabs(["📋 조회", "🔗 매칭", "📦 매칭본"])

# ─────────────────────────────── 탭1: 조회 ───────────────────────────────
with tab1:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("전체", n_total)
    c2.metric("매칭", n_matched)
    c3.metric("검토완료(없음)", n_none)
    c4.metric("미검토", n_todo)

    q = st.text_input("상품명 검색", key="q_view").strip()
    view = prices.copy()
    view["매칭상태"] = [
        {"matched": "✅매칭", "none": "⚪없음"}.get(goid_status.get(_nfc(g)), "🔲미검토")
        for g in view["ps_goid"]
    ]
    if q:
        view = view[view["name"].str.contains(q, case=False, na=False)]
    show = view[["매칭상태", "name", "spec", "box_price", "unit_price", "ps_page", "ps_goid"]].rename(
        columns={"name": "상품명", "spec": "규격", "box_price": "박스가",
                 "unit_price": "개당가", "ps_page": "페이지"})
    st.dataframe(show, use_container_width=True, hide_index=True,
                 column_config={"박스가": st.column_config.NumberColumn(format="%d"),
                                "개당가": st.column_config.NumberColumn(format="%d")})

# ─────────────────────────────── 탭2: 매칭 ───────────────────────────────
with tab2:
    st.markdown(f"**미검토 {n_todo}건** · 매칭 {n_matched} · 없음 {n_none}")
    todo = prices[[goid_status.get(_nfc(g)) is None for g in prices["ps_goid"]]].reset_index(drop=True)
    if todo.empty:
        st.success("미검토 항목이 없습니다. 모두 처리되었습니다.")
    else:
        # 저신뢰 일괄 없음 — 최고 점수 < 0.3(또는 후보 없음). 진짜 매칭 최저가 ~0.38이라 안전.
        AUTO_THR = 0.3
        low_idx = []
        for ti in range(len(todo)):
            r = todo.iloc[ti]
            cs = M.suggest(r["name"], r["spec"], idx, topn=1)
            if (cs[0]["점수"] if cs else 0.0) < AUTO_THR:
                low_idx.append(ti)
        if low_idx:
            if st.button(f"⚪ 확신 낮은 {len(low_idx)}건 일괄 '없음' (최고점수<{AUTO_THR}·후보없음 포함)",
                         key="bulk_none"):
                new = mp.copy()
                today = date.today().isoformat()
                bulk = []
                for ti in low_idx:
                    rr = todo.iloc[ti]; g = _nfc(rr["ps_goid"])
                    new = new[new["ps_goid"].map(_nfc) != g]
                    bulk.append({"ps_goid": g, "nadl_name": rr["name"], "nadl_spec": rr["spec"],
                                 "관리코드": "", "status": "none", "updated": today})
                new = pd.concat([new, pd.DataFrame(bulk, columns=M.MAP_COLS)], ignore_index=True)
                M.write_map(new, pat, repo)
                _bump_map()
                st.success(f"{len(bulk)}건 '없음' 처리 완료")
                st.rerun()
        labels = [f"{r['name']}  ·  {r['spec']}  ·  개당 {_won(r['unit_price'])}"
                  for _, r in todo.iterrows()]
        pick = st.selectbox("매칭할 항목", range(len(todo)),
                            format_func=lambda i: labels[i], key="match_pick")
        nr = todo.iloc[pick]
        goid = _nfc(nr["ps_goid"])

        st.markdown(f"### {nr['name']}")
        st.markdown(f"#### 규격 {nr['spec']} · 박스가 {_won(nr['box_price'])} · "
                    f"개당가 {_won(nr['unit_price'])}")

        cands = M.suggest(nr["name"], nr["spec"], idx, topn=3)
        box = M.nadl_box(nr["name"], nr["spec"])
        if box is None:
            st.info("규격(용량·팩수)을 읽지 못해 자동 후보를 만들 수 없습니다. "
                    "아래에서 관리코드를 직접 입력하거나 '없음'으로 처리하세요.")

        chosen = []
        if cands:
            st.markdown("**후보 (점수 높은 순)** — 맞는 것을 고르세요. 중복등록이면 여러 개 체크.")
            ct = pd.DataFrame(cands)
            ct = ct[["점수", "관리코드", "상품명", "규격", "매입단가", "매출단가"]]
            st.dataframe(ct, use_container_width=True, hide_index=True,
                         column_config={"매입단가": st.column_config.NumberColumn(format="%d"),
                                        "매출단가": st.column_config.NumberColumn(format="%d")})
            for c in cands:
                lbl = f"{c['관리코드']} · {c['상품명']} · 매입{_won(c['매입단가'])}/매출{_won(c['매출단가'])} · 점수 {c['점수']}"
                if st.checkbox(lbl, key=f"cand_{goid}_{c['관리코드']}"):
                    chosen.append(c["관리코드"])
        else:
            st.caption("박스규격 일치 후보 없음 → 직접 입력하거나 '없음'.")

        extra = st.text_input("직접 관리코드 입력(쉼표로 여러 개)", key=f"extra_{goid}").strip()
        if extra:
            pm_codes = set(pm["관리코드"].map(_nfc))
            for code in [_nfc(x) for x in extra.split(",") if x.strip()]:
                if code in pm_codes:
                    chosen.append(code)
                else:
                    st.warning(f"관리코드 '{code}' 가 product_master에 없습니다.")

        chosen = list(dict.fromkeys(chosen))  # 중복 제거·순서 보존
        b1, b2 = st.columns(2)
        if b1.button(f"✅ 매칭 저장 ({len(chosen)}개)" if chosen else "✅ 매칭 저장",
                     type="primary", disabled=not chosen, key="save_match"):
            new = mp[mp["ps_goid"].map(_nfc) != goid].copy()
            rows = [{"ps_goid": goid, "nadl_name": nr["name"], "nadl_spec": nr["spec"],
                     "관리코드": code, "status": "matched",
                     "updated": date.today().isoformat()} for code in chosen]
            new = pd.concat([new, pd.DataFrame(rows, columns=M.MAP_COLS)], ignore_index=True)
            M.write_map(new, pat, repo)
            _bump_map()
            st.success(f"저장 완료: {nr['name']} → {', '.join(chosen)}")
            st.rerun()
        if b2.button("⚪ 없음(검토완료)", key="save_none"):
            new = mp[mp["ps_goid"].map(_nfc) != goid].copy()
            row = pd.DataFrame([{"ps_goid": goid, "nadl_name": nr["name"], "nadl_spec": nr["spec"],
                                 "관리코드": "", "status": "none",
                                 "updated": date.today().isoformat()}], columns=M.MAP_COLS)
            new = pd.concat([new, row], ignore_index=True)
            M.write_map(new, pat, repo)
            _bump_map()
            st.success(f"'없음' 처리: {nr['name']}")
            st.rerun()

# ─────────────────────────────── 탭3: 매칭본 ──────────────────────────────
with tab3:
    matched_tbl = M.build_matched(prices, mp, pm)
    st.markdown(f"전체 {len(matched_tbl)}행 · 매칭 {(matched_tbl['상태'] == '매칭').sum()} · "
                f"검토완료(없음) {(matched_tbl['상태'] == '검토완료(없음)').sum()} · "
                f"미매칭 {(matched_tbl['상태'] == '미매칭').sum()}")
    only_matched = st.checkbox("매칭된 것만 보기", value=False, key="only_m")
    tbl = matched_tbl[matched_tbl["상태"] == "매칭"] if only_matched else matched_tbl
    st.dataframe(tbl, use_container_width=True, hide_index=True,
                 column_config={"박스가": st.column_config.NumberColumn(format="%d"),
                                "개당가": st.column_config.NumberColumn(format="%d"),
                                "우리_매입단가": st.column_config.NumberColumn(format="%d"),
                                "우리_매출단가": st.column_config.NumberColumn(format="%d")})

    csv = matched_tbl.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ 매칭본 CSV", csv,
                       file_name=f"nadl_매칭본_{sel_date}.csv", mime="text/csv")

    with st.expander("🛠 매핑 관리 (삭제·재매칭 · '없음' 되돌리기)"):
        flt = st.radio("표시", ["매칭", "없음", "전체"], horizontal=True, key="map_flt")
        if flt == "전체":
            mm = mp.copy()
        else:
            mm = mp[mp["status"] == ("matched" if flt == "매칭" else "none")].copy()
        if mm.empty:
            st.caption("항목이 없습니다.")
        else:
            mm = mm[["status", "nadl_name", "nadl_spec", "관리코드", "updated", "ps_goid"]].rename(
                columns={"status": "상태", "nadl_name": "nadl_상품명", "nadl_spec": "nadl_규격"})
            st.dataframe(mm, use_container_width=True, hide_index=True)
            st.caption("삭제하면 해당 매핑이 제거됩니다('없음'이면 다시 검토 큐로 돌아옴).")
            del_goid = st.text_input("삭제할 ps_goid", key="del_goid").strip()
            if st.button("선택 매핑 삭제", disabled=not del_goid, key="del_btn"):
                new = mp[mp["ps_goid"].map(_nfc) != _nfc(del_goid)].copy()
                M.write_map(new, pat, repo)
                _bump_map()
                st.success(f"ps_goid {del_goid} 매핑 삭제")
                st.rerun()
