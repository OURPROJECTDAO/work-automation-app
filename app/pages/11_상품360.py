"""상품 360도 카드 — 관리코드 1개의 *모든 것*을 한 화면에.

상품이 주인공. 흩어진 두뇌(채널 실판매·등재현황·목표마진 역산·매입가 추이·재고)를 상품 단위로 묶음.
답하는 질문(사용자 정의 MVP):
  ① 어디 채널에서 / 어느 가격에 / 매출 얼마나 팔리나   → 채널별 실판매(channel_compare)
  ② 어디 채널 어디 채널 올라갔나                        → 8채널 등재 현황(listing 스캔)
  ③ A 채널에서 마진 N% 보려면 얼마로 올려야 하나        → 채널별 목표마진 역산(매입가·수수료)
  ④ 매입가 올랐나 내렸나 / 최근 매입 언제                → 매입가 그래프(실입고 + 수정로그) + 최근 입고

전부 기존 순수함수 오케스트레이션 — core 무변경(page-only). 관리코드 조인이라 가능.
"""
import math
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import pandas as pd
import streamlit as st

from core.dashboard import store
from core.dashboard.sales_data import make_box_lookup
from core.intelligence import channel_compare as cc
from core.intelligence import orders as _orders
from core.intelligence import price_history as ph
from core.intelligence import purchases as _pur
from core.intelligence import ship_alloc
from core.intelligence import stockout
from core.workflows import channel_margin_monitor as cmm

_REF = Path(__file__).parent.parent.parent / "reference"

st.title("🪪 상품 360도 카드")
st.caption("관리코드 하나를 고르면 그 상품의 **모든 것** — 어디 채널에서 얼마에 얼마나 팔리는지, "
           "어디에 올라가 있는지, 마진 목표 가격, 매입가 추이까지 한 화면에 모읍니다.")


def _nfc(s) -> str:
    return unicodedata.normalize("NFC", str(s)).strip()


def _won(v) -> str:
    try:
        return f"₩{int(round(float(v))):,}"
    except (TypeError, ValueError):
        return "—"


def _pct(v) -> str:
    try:
        return f"{float(v) * 100:.1f}%"
    except (TypeError, ValueError):
        return "—"


def _ceil100(x: float) -> int:
    return int(math.ceil(x / 100) * 100)


def _data_secret() -> tuple[str, str]:
    repo = "OURPROJECTDAO/work-automation-data"
    try:
        d = st.secrets["data"]
        return d["pat"], d.get("repo", repo)
    except Exception:
        return st.secrets.get("GITHUB_PAT", ""), repo


# ── 데이터 로더 (가격 A/B 페이지와 동일 패턴) ──────────────────────────────
@st.cache_data(ttl=3600, show_spinner="매출 데이터 불러오는 중...")
def load_sales_min(pat: str, repo: str) -> pd.DataFrame:
    df = store.load_master(pat, repo)
    if df.empty:
        return df
    attr = pd.read_csv(_REF / "product_attributes.csv", dtype=str, encoding="utf-8-sig")
    pm = pd.read_csv(_REF / "product_master.csv", dtype=str, encoding="utf-8-sig")
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
    return df


@st.cache_data(ttl=3600, show_spinner="EasyAdmin 주문 불러오는 중...")
def load_orders(pat: str, repo: str) -> pd.DataFrame:
    return _orders.read_all(pat, repo)


@st.cache_data(ttl=3600, show_spinner="송장 실배분 계산 중...")
def load_ship_rate(pat: str, repo: str) -> dict:
    od = load_orders(pat, repo)
    if od.empty:
        return {"rate": {}, "ch_rate": {}, "reconcile": {}, "stats": {"codes": 0}}
    pm = pd.read_csv(_REF / "product_master.csv", dtype=str, encoding="utf-8-sig")
    try:
        hapo = set(pd.read_csv(_REF / "hapo_175_190.csv", dtype=str,
                               encoding="utf-8-sig")["관리코드"].dropna())
    except Exception:
        hapo = None
    return ship_alloc.compute_ship_rate(od, load_sales_min(pat, repo),
                                        make_box_lookup(pm), hapo_codes=hapo)


@st.cache_data(ttl=3600, show_spinner=False)
def load_ea_agg(pat: str, repo: str) -> pd.DataFrame:
    return cc.build_ea_price_agg(load_orders(pat, repo))


@st.cache_data(ttl=3600, show_spinner=False)
def load_buyin(pat: str, repo: str) -> pd.DataFrame:
    return _pur.read_all(pat, repo)


@st.cache_data(ttl=3600, show_spinner=False)
def load_price_hist(pat: str, repo: str) -> pd.DataFrame:
    return ph.read_history(pat, repo)


@st.cache_data(ttl=3600, show_spinner=False)
def load_pm() -> pd.DataFrame:
    return pd.read_csv(_REF / "product_master.csv", dtype=str, encoding="utf-8-sig")


@st.cache_data(ttl=3600, show_spinner=False)
def load_online(pat: str, repo: str) -> set:
    g = store.read_groups(pat, repo)
    if g.empty:
        return set()
    return {_nfc(r["상호명"]) for _, r in g.iterrows()
            if str(r.get("그룹", "")).strip() == "온라인" and pd.notna(r.get("상호명"))}


@st.cache_data(ttl=3600, show_spinner=False)
def load_refs() -> dict:
    return cmm.load_references(_REF)


@st.cache_data(ttl=3600, show_spinner=False)
def load_listing_recs(key: str) -> list[dict]:
    p = _REF / f"listing_{key}.csv"
    if not p.exists():
        return []
    return cmm.csv_text_to_recs(p.read_text(encoding="utf-8-sig"))


@st.cache_data(ttl=3600, show_spinner="재고 소진 계산 중...")
def load_forecast(pat: str, repo: str) -> pd.DataFrame:
    sales = load_sales_min(pat, repo)
    buyin = load_buyin(pat, repo)
    dep = stockout.depletion_rate(sales, months=3)
    cad = stockout.restock_cadence(buyin, months=12)
    return stockout.forecast(load_pm(), dep, cad)


# ── 매입가 정규화 (박스단위 오입력 보정 — margin_erosion 동형) ──────────────
def _unit_buy(price, box_price, box_qty):
    """단가가 박스단가로 잘못 들어온 경우(단가==박스단가 & 박스내품>1) 낱개로 환산."""
    try:
        p, bp, bq = float(price), float(box_price), float(box_qty)
    except (TypeError, ValueError):
        return None
    if bq > 1 and bp > 0 and abs(p - bp) < 1e-6:
        return p / bq
    return p


pat, repo = _data_secret()
if not pat:
    st.warning("저장소 접근 정보(secrets `[data] pat`)가 설정되지 않았습니다.")
    st.stop()

pm = load_pm()
if pm.empty:
    st.info("상품마스터가 비어 있습니다.")
    st.stop()
refs = load_refs()
df = load_sales_min(pat, repo)

# ── 관리코드 선택 (상품마스터 전체 — 상품이 주인공) ─────────────────────────
opt = pm[["관리코드", "상품명", "규격", "상품코드"]].copy()
opt["관리코드"] = opt["관리코드"].map(_nfc)
opt = opt[opt["관리코드"] != ""].drop_duplicates("관리코드")
# 정렬: 온라인 매출 큰 상품 먼저(자주 찾는 상품 상단)
if not df.empty:
    online0 = load_online(pat, repo)
    sv = df[df["상호명"].astype(str).map(_nfc).isin(online0)] if online0 else df
    rev = sv.groupby(sv["관리코드"].astype(str).map(_nfc))["판매금액"].sum()
    opt["_매출"] = opt["관리코드"].map(rev).fillna(0.0)
else:
    opt["_매출"] = 0.0
opt = opt.sort_values(["_매출", "관리코드"], ascending=[False, True])

q = st.text_input("관리코드 / 상품명 검색", key="p360_q",
                  placeholder="예: 코카콜라 · 31-01-04").strip()
view = opt
if q:
    qn = _nfc(q)
    m = (view["관리코드"].str.contains(qn, case=False, na=False)
         | view["상품명"].astype(str).map(_nfc).str.contains(qn, case=False, na=False))
    view = view[m]
if view.empty:
    st.info("검색 결과가 없습니다.")
    st.stop()
view = view.head(300)
labels = {f"{r['관리코드']} · {r['상품명']}": r["관리코드"] for _, r in view.iterrows()}
pick = st.selectbox(f"관리코드 ({len(view)}개)", list(labels.keys()), key="p360_code")
code = labels[pick]
crow = opt[opt["관리코드"] == code].iloc[0]
sangpum = crow["상품명"]
prodcode = _nfc(crow.get("상품코드", ""))

# ── A. 헤더 (상품 기본 정보) ───────────────────────────────────────────────
typ, base, stock, spec, note = cmm.resolve_code(code, refs)
pmrow = refs["pm_by_mgmt"].get(code, {})
box_qty = cmm._num(pmrow.get("박스내품"), 0) if pmrow else 0
box_buy = cmm._num(pmrow.get("박스매입단가"), 0) if pmrow else 0
unit_buy = cmm._num(pmrow.get("매입단가"), 0) if pmrow else 0

st.subheader(f"📦 {code} · {sangpum}")
h1, h2, h3, h4 = st.columns(4)
h1.metric("코드유형", typ)
h2.metric("규격", spec or "—")
h3.metric("매입가(박스/낱개)",
          f"{_won(box_buy)} / {_won(unit_buy)}" if box_buy or unit_buy else "—")
h4.metric("박스재고 / 박스내품",
          f"{int(stock):,} / {int(box_qty)}" if stock is not None else "—")
if note:
    st.caption(f"⚠️ {note}")

# ── B+C. 채널별 등재 현황 + 지금 가격/마진 + 목표마진 역산 ────────────────────
st.markdown("### 🛒 채널별 — 등재 현황 · 현재 가격 · 목표마진 가격")
tgt = st.number_input("목표 마진율 (%) — 이만큼 보려면 얼마에 올려야 하나",
                      min_value=0.0, max_value=80.0, value=10.0, step=0.5, key="p360_tgt") / 100.0

listing_rows = []
present = {}      # 채널 → 등재여부
for ch, cfg in cmm.CHANNEL_CONFIG.items():
    recs = load_listing_recs(cfg["key"])
    sub = [r for r in recs if _nfc(r["코드"]) == code]
    present[ch] = bool(sub)
    if not sub:
        continue
    rows = cmm.compute(sub, refs, cfg)
    settle, ship = cfg["ship_settle"], cfg["real_ship"]
    for rec, row in zip(sub, rows):
        매입가 = row["매입가"]
        판매가 = row["판매가"]
        정산액 = row["정산액"]
        net = 판매가 - rec.get("즉시할인", 0) - rec.get("포인트", 0)
        # 정산액에서 유효 정산비율(rate=1-수수료) 역산 → 채널 단일/상품별 수수료 모두 정확
        need = None
        if 매입가 is not None and 정산액 is not None and net > 0 and tgt < 1:
            rate = (정산액 - rec.get("배송비", 0) * settle) / net
            if rate > 0:
                tgt_settle = (매입가 + ship) / (1 - tgt)
                net_new = (tgt_settle - rec.get("배송비", 0) * settle) / rate
                need = _ceil100(net_new + rec.get("즉시할인", 0) + rec.get("포인트", 0))
        listing_rows.append({
            "채널": ch,
            "리스팅(상품번호)": row["상품번호"],
            "N": row["N"],
            "현재판매가": 판매가,
            "매입가": 매입가,
            "현재마진율": row["마진율"],
            "기준마진율": row["기준마진율"],
            "기준여유": row["탐지"],
            f"필요판매가(마진 {tgt * 100:.1f}%)": need,
            "인상폭": (need - 판매가) if need is not None else None,
            "제한": row["제한"],
        })

# 등재 현황 스트립 (어디 채널 어디 채널 올라갔나)
strip = "  ".join(f"{'🟢' if present.get(ch) else '⚪'} {ch}" for ch in cmm.CHANNEL_CONFIG)
n_up = sum(1 for v in present.values() if v)
st.caption(f"**등재 {n_up}/{len(present)}채널** —  {strip}")

if listing_rows:
    ldf = pd.DataFrame(listing_rows)
    st.dataframe(
        ldf, width="stretch", hide_index=True,
        column_config={
            "채널": st.column_config.TextColumn(width="small"),
            "리스팅(상품번호)": st.column_config.TextColumn(width="small"),
            "N": st.column_config.NumberColumn("N(합포)", format="%.2g"),
            "현재판매가": st.column_config.NumberColumn(format="localized"),
            "매입가": st.column_config.NumberColumn(format="localized"),
            "현재마진율": st.column_config.NumberColumn(format="percent"),
            "기준마진율": st.column_config.NumberColumn(format="percent"),
            "기준여유": st.column_config.NumberColumn(format="percent",
                                                 help="현재마진율 − 기준마진율 (음수=미달 🔴)"),
            f"필요판매가(마진 {tgt * 100:.1f}%)": st.column_config.NumberColumn(
                format="localized", help="목표 마진율을 채우는 판매가(2700·올림 표준)"),
            "인상폭": st.column_config.NumberColumn(format="localized",
                                                help="필요판매가 − 현재판매가 (양수=올려야)"),
        },
        height=min(440, 80 + 36 * min(len(ldf), 10)))
    st.caption("ⓘ 같은 관리코드가 한 채널에 여러 리스팅(합포 N 상이)이면 각각 행으로 펼칩니다. "
               "**필요판매가**=현재 수수료·배송비·매입가 그대로 두고 목표 마진율만 채우는 판매가. "
               "가격 변경은 **채널마진모니터**에서 (이 카드는 조회 전용).")
else:
    st.info("8개 채널 저장 리스팅 어디에도 이 관리코드가 없습니다. (미등재 또는 합성코드)")

# ── D. 채널별 실판매 (어디서 얼마에 매출 얼마나) ─────────────────────────────
st.markdown("### 💰 채널별 실판매 — 매출 · 판매량 · 실현마진 (최근 3개월)")
if df.empty:
    st.info("적재된 매출 데이터가 없습니다.")
else:
    gmap_online = load_online(pat, repo)
    dmax = df["거래일자"].max()
    ts0 = pd.Timestamp((pd.Timestamp(dmax) - pd.DateOffset(months=3)).date())
    ts1 = pd.Timestamp(dmax) + pd.Timedelta(days=1)
    vsel = df[(df["거래일자"] >= ts0) & (df["거래일자"] < ts1)]
    if gmap_online:
        vsel = vsel[vsel["상호명"].astype(str).map(_nfc).isin(gmap_online)]
    vsel = vsel.copy()
    if vsel.empty:
        st.info("최근 3개월 온라인 매출이 없습니다.")
    else:
        ship = load_ship_rate(pat, repo)
        prod = cc.compute_online_margin(vsel, ship, 3000.0, use_actual=True)
        ea = load_ea_agg(pat, repo)
        months = cc.months_in_range(ts0.date(), pd.Timestamp(dmax).date())
        ep = cc.ea_price_lookup(ea, code, months)
        bd = cc.channel_breakdown(prod, code, ep)
        if bd.empty:
            st.info("최근 3개월 이 상품의 온라인 판매 기록이 없습니다.")
        else:
            tot = bd[["매출", "판매량"]].sum()
            k1, k2, k3 = st.columns(3)
            k1.metric("총 매출(온라인)", _won(tot["매출"]))
            k2.metric("총 판매량(낱개)", f"{int(tot['판매량']):,}")
            k3.metric("판매 채널 수", f"{len(bd)}")
            show = bd.drop(columns=["_상호명"])
            st.dataframe(
                show, width="stretch", hide_index=True,
                column_config={
                    "마진율(%)": st.column_config.NumberColumn("실현마진율(%)", format="%.2f",
                                                            help="순이익 ÷ 매입가 (택배 실배분)"),
                    "낱개이익": st.column_config.NumberColumn("낱개이익(원)", format="%.1f"),
                    "매출": st.column_config.NumberColumn(format="localized"),
                    "판매량": st.column_config.NumberColumn("판매량(낱개)", format="localized"),
                    "정산단가": st.column_config.NumberColumn("정산단가(net)", format="localized",
                                                         help="매출÷판매량 — 실수령 낱개단가"),
                    "노출가": st.column_config.NumberColumn("노출가(EA)", format="localized",
                                                        help="EasyAdmin 소비자가(판매단위)"),
                    "택배": st.column_config.TextColumn(width="small"),
                },
                height=min(400, 80 + 36 * min(len(show), 9)))
            cL, cR = st.columns(2)
            with cL:
                st.caption("매출 (원)")
                st.bar_chart(bd.set_index("채널")["매출"], height=220)
            with cR:
                st.caption("실현마진율 (%)")
                st.bar_chart(bd.set_index("채널")["마진율(%)"], height=220)

# ── E. 매입가 추이 + 최근 매입 (올랐나 내렸나) ──────────────────────────────
st.markdown("### 📈 매입가 추이 — 올랐나 내렸나 · 최근 매입은 언제")
buyin = load_buyin(pat, repo)
bsub = buyin[buyin["관리코드"].astype(str).map(_nfc) == code].copy() if not buyin.empty else pd.DataFrame()
if not bsub.empty:
    bsub["_d"] = pd.to_datetime(bsub["기준일"], errors="coerce")
    _h = pd.to_numeric(bsub["합계액"].astype(str).str.replace(",", "", regex=False), errors="coerce")
    _q = pd.to_numeric(bsub["수량"].astype(str).str.replace(",", "", regex=False), errors="coerce")
    bsub = bsub[(_h > 0) & (_q > 0)].copy()
if bsub is None or bsub.empty:
    st.info("이 관리코드의 실입고(매입현황) 기록이 없습니다.")
else:
    bsub["_단가"] = [
        _unit_buy(p, bp, bq)
        for p, bp, bq in zip(bsub["단가"], bsub["박스단가"], bsub["박스내품"])
    ]
    bsub = bsub.dropna(subset=["_단가", "_d"]).sort_values("_d")
    last_d = bsub["_d"].max()
    last_p = bsub.loc[bsub["_d"] == last_d, "_단가"].iloc[-1]
    # 6개월 전 대비
    ref_d = last_d - pd.DateOffset(months=6)
    prior = bsub[bsub["_d"] <= ref_d]
    base_p = prior["_단가"].iloc[-1] if not prior.empty else bsub["_단가"].iloc[0]
    delta = last_p - base_p
    m1, m2, m3 = st.columns(3)
    m1.metric("최근 매입가(낱개)", _won(last_p))
    m2.metric("최근 매입일", last_d.date().isoformat())
    m3.metric("vs 6개월 전", _won(last_p),
              delta=f"{'+' if delta >= 0 else ''}{int(round(delta)):,}원")
    # 월별 평균 실입고가 라인
    ser = bsub.set_index("_d")["_단가"].resample("MS").mean().dropna()
    if len(ser) >= 1:
        ser.index = ser.index.strftime("%Y-%m")
        st.line_chart(ser.rename("실입고 매입가(낱개)"), height=240)

# 수정로그(장부 매입가 변경 원장) — 최근 변경 이벤트
if prodcode and not (hist := load_price_hist(pat, repo)).empty:
    hsub = hist[(hist["수정항목"] == "매입단가")
                & (hist["상품코드"].astype(str).map(_nfc).isin({prodcode, ph._code6(prodcode)}))]
    if not hsub.empty:
        hsub = hsub.sort_values("수정일자", ascending=False).head(8).copy()
        hsub["수정일자"] = pd.to_datetime(hsub["수정일자"]).dt.date.astype(str)
        with st.expander(f"🧾 장부 매입가 변경 이력 (수정로그 · 최근 {len(hsub)}건)"):
            st.dataframe(hsub[["수정일자", "수정전", "수정후"]], hide_index=True,
                         width="stretch",
                         column_config={
                             "수정전": st.column_config.NumberColumn(format="localized"),
                             "수정후": st.column_config.NumberColumn(format="localized"),
                         })

# ── F. 재고 · 소진 (이 상품 모든 것) ────────────────────────────────────────
st.markdown("### 🔮 재고 · 소진 예측")
fc = load_forecast(pat, repo)
frow = fc[fc["관리코드"].astype(str).map(_nfc) == code] if not fc.empty else pd.DataFrame()
if frow.empty:
    st.info("재고/소진 데이터가 없습니다.")
else:
    r = frow.iloc[0]
    f1, f2, f3, f4 = st.columns(4)
    f1.metric("구간", str(r["구간"]))
    ttl = r["소진예측일"]
    f2.metric("소진 예측일", f"{ttl:.0f}일" if pd.notna(ttl) else "—",
              help=f"예상 소진일자 {r['예상소진일자']}" if pd.notna(r["예상소진일자"]) else "")
    f3.metric("리드타임", f"{r['리드타임']:.0f}일 ({r['리드출처']})")
    f4.metric("재고금액", _won(r["재고금액"]))
    st.caption(f"현재고 {int(r['현재고(낱개)']):,}낱개 · 박스 {int(r['박스재고']):,} · "
               f"일소진 {r['일소진(낱개)']:.1f}낱개 · 월소진액(매입) {_won(r['월소진액(매입)'])} · "
               f"최근입고 {r['최근입고일']} (입고 {int(r['입고횟수'])}회)")

st.divider()
st.caption("ⓘ 상품 360도 카드 v1 — 조회 전용. 가격 변경은 [채널마진모니터], 채널 비교 심화는 [가격 A/B], "
           "전체 재고 지능은 [재고지능] 페이지에서.")
