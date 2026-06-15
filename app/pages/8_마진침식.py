"""🩸 마진 침식 (두뇌① 강화, ADR 0020) — 3렌즈로 마진 위험 감지.

탭 A 이미 침식   : listing 매입↑(가격이력) ∩ 채널마진 미달 — 반응·잠재(기존 v1).
탭 B 곧 침식(2C) : 실입고가 > master 매입가 ∩ master 미수정 — 예방·잠재(매입현황 vs master).
탭 C 실판매 이상 : 매출자료 실현마진 = 역마진 OR 채널 baseline−2%p 미달 — 진실·실현(합성코드 포함).
공통             : velocity(EA 주문) 월손실액 = 매입Δ(원/낱개) × 월낱개판매량 — 정렬축.

★ 마진 기준 = master 매입가(ERP가 상품관리에서 직접 고친 값). 매출자료 판매이익도 같은 기준(사용자 확정).
"""
import io
import csv
import json
import sys
import urllib.error
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import pandas as pd
import streamlit as st

from core.workflows import channel_margin_monitor as cmm
from core.intelligence import margin_erosion as me
from core.intelligence import price_history
from core.intelligence import orders as od_mod
from core.intelligence import purchases as buy_mod

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


def _nfc(v):
    import unicodedata
    return unicodedata.normalize("NFC", str(v)).strip() if pd.notna(v) else ""


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


@st.cache_data(ttl=600, show_spinner=False)
def _baseline_dict():
    """{관리코드(NFC): {채널: 기준마진}} — 실판매 이상(탭C)용 직접 파싱."""
    code, text = _gh_raw("reference/baseline_margin.csv")
    if code != 200:
        return {}
    lines = text.decode("utf-8-sig").splitlines()
    rd = csv.DictReader(lines)
    chans = [c for c in (rd.fieldnames or []) if c != "관리코드"]
    out = {}
    for row in csv.DictReader(lines):
        k = _nfc(row["관리코드"])
        out[k] = {c: float(row[c]) for c in chans if row[c] not in ("", "None")}
    return out


@st.cache_data(ttl=600, show_spinner=False)
def _master_lookup():
    """관리코드(NFC) → (박스내품, master매입단가, 상품명)."""
    code, text = _gh_raw("reference/product_master.csv")
    if code != 200:
        return {}, {}, {}
    df = pd.read_csv(io.BytesIO(text), dtype=str)
    df["k"] = df["관리코드"].map(_nfc)
    box = {k: float(v) for k, v in zip(df["k"], pd.to_numeric(df["박스내품"], errors="coerce").fillna(1.0))}
    price = {k: float(v) for k, v in zip(df["k"], pd.to_numeric(df["매입단가"], errors="coerce")) if pd.notna(v)}
    name = {k: n for k, n in zip(df["k"], df["상품명"].fillna(""))}
    return box, price, name


@st.cache_data(ttl=600, show_spinner="가격이력 불러오는 중...")
def _price_changes() -> pd.DataFrame:
    pat, repo = _data_secret()
    return price_history.read_history(pat, repo) if pat else pd.DataFrame()


@st.cache_data(ttl=600, show_spinner="주문(velocity) 불러오는 중...")
def _orders_recent(months: int) -> pd.DataFrame:
    pat, repo = _data_secret()
    if not pat:
        return pd.DataFrame()
    mons = od_mod.list_months(pat, repo)[-months:]
    parts = [od_mod.read_partition(pat, repo, m) for m in mons]
    parts = [p for p in parts if p is not None and len(p)]
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


@st.cache_data(ttl=600, show_spinner="매입현황 불러오는 중...")
def _buyin_recent(months: int) -> pd.DataFrame:
    pat, repo = _data_secret()
    if not pat:
        return pd.DataFrame()
    mons = buy_mod.list_months(pat, repo)[-months:]
    parts = [buy_mod.read_partition(pat, repo, m) for m in mons]
    parts = [p for p in parts if p is not None and len(p)]
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


@st.cache_data(ttl=600, show_spinner="매출자료 불러오는 중...")
def _sales_recent(months: int) -> pd.DataFrame:
    pat, repo = _data_secret()
    if not pat:
        return pd.DataFrame()
    import re
    url = f"https://api.github.com/repos/{repo}/contents/master?ref=main"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {pat}", "Accept": "application/vnd.github+json"})
    try:
        items = json.load(urllib.request.urlopen(req))
    except urllib.error.HTTPError:
        return pd.DataFrame()
    mons = sorted(m.group(1) for it in items for m in [re.match(r"sales_(\d{4}-\d{2})\.parquet$", it.get("name", ""))] if m)[-months:]
    parts = []
    for ym in mons:
        c, t = _gh_raw(f"master/sales_{ym}.parquet")
        if c == 200:
            parts.append(pd.read_parquet(io.BytesIO(t)))
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def _to_xlsx(df: pd.DataFrame, title: str) -> bytes:
    from openpyxl import Workbook
    def _cell(v):
        if pd.isna(v):
            return None
        if hasattr(v, "isoformat"):
            return v.isoformat()
        if hasattr(v, "item"):
            return v.item()
        return v
    wb = Workbook(); ws = wb.active; ws.title = title[:31]
    ws.append(list(df.columns))
    for _, r in df.iterrows():
        ws.append([_cell(v) for v in r])
    if "관리코드" in df.columns:
        col = df.columns.get_loc("관리코드") + 1
        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            for c in row:
                c.number_format = "@"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ─────────────────────────────────────────────────────────
st.title("🩸 마진 침식")
st.caption("마진 위험을 세 각도로 봅니다 — **이미 침식**(가격표 미달) · **곧 침식**(실입고가↑·master 미반영) · "
           "**실판매 이상**(실제 팔린 마진). 기준 = 상품관리(master) 매입가.")

box_lookup, master_price, name_lookup = _master_lookup()

with st.expander("⚙️ 설정", expanded=False):
    cc = st.columns(4)
    months_buy = cc[0].number_input("매입 인상 기간(개월)", 1, 12, 3, 1,
                                    help="탭A·B: 최근 몇 개월 매입가 인상/실입고를 볼지")
    vel_months = cc[1].number_input("판매량(velocity) 기간(개월)", 1, 12, 3, 1,
                                    help="월 손실액 계산용 평균 판매량 기간")
    min_pct = cc[2].number_input("2C 최소 상승%", 1, 50, 3, 1,
                                 help="탭B: 실입고가가 master보다 이만큼↑일 때 경보") / 100.0
    buffer = cc[3].number_input("실판매 안전마진 여유(%p)", 0, 10, 2, 1,
                                help="탭C: 채널 기준마진보다 이만큼↓이면 이상") / 100.0

vel = me.channel_velocity(_orders_recent(int(vel_months)), box_lookup, months=int(vel_months)) \
    if box_lookup else {"by_code_ch": {}, "by_code": {}}
by_ch, by_code = vel["by_code_ch"], vel["by_code"]

tabA, tabB, tabC = st.tabs(["🩸 이미 침식", "⏳ 곧 침식 (2C)", "📉 실판매 이상"])

# ── 탭 A: 이미 침식 (기존 v1 + velocity 손실액) ──
with tabA:
    pc = _price_changes()
    if pc.empty:
        st.warning("가격이력(price_changes)이 없습니다. secrets `[data] pat` 또는 수정로그 적재를 확인해 주세요.")
    else:
        _chan_keys = {ch: cfg["key"] for ch, cfg in cmm.CHANNEL_CONFIG.items()}
        chans = st.multiselect("채널", list(_chan_keys.keys()), default=list(_chan_keys.keys()), key="a_ch")
        raises = me.recent_buy_raises(pc, months=int(months_buy), now=pd.Timestamp.now())
        bo = _baseline_override()
        all_rows = []
        for ch in chans:
            recs, meta = _listing(_chan_keys[ch])
            if recs is None:
                continue
            rows, _ = cmm.compute_listing(recs, ch, str(_REF), baseline_override=bo)
            all_rows += me.erosion_rows(rows, ch, raises)
        df = me.to_frame(all_rows)
        if df.empty:
            st.success("✅ 최근 매입가 인상으로 기준마진 아래로 떨어진(미수정) 상품이 없습니다.")
        else:
            df["월손실액"] = [((r["현재매입"] - r["과거매입"]) * by_ch.get((r["채널"], _nfc(r["관리코드"])), 0.0))
                           for _, r in df.iterrows()]
            df = df.sort_values("월손실액", ascending=False).reset_index(drop=True)
            c = st.columns(3)
            c[0].metric("침식 경보", f"{len(df)} 건")
            c[1].metric("관리코드", f"{df['관리코드'].nunique()} 개")
            c[2].metric("월 손실액 합", f"{df['월손실액'].sum():,.0f} 원")
            _disp = df.copy()
            for col in ("매입Δ%", "마진율", "기준마진율", "미달폭"):
                _disp[col] = (_disp[col].astype(float) * 100).round(1)
            st.dataframe(_disp, hide_index=True, use_container_width=True, height=440,
                         column_config={
                             "월손실액": st.column_config.NumberColumn("월손실액(원)", format="%d",
                                          help="매입상승분 × 월 판매낱개 — 아직 못 올려 흡수 중인 금액"),
                             "매입Δ%": st.column_config.NumberColumn("매입↑%", format="%.0f"),
                             "마진율": st.column_config.NumberColumn("현재마진%", format="%.1f"),
                             "기준마진율": st.column_config.NumberColumn("기준마진%", format="%.1f"),
                             "미달폭": st.column_config.NumberColumn("미달폭%p", format="%.1f"),
                             "마지막인상일": st.column_config.DatetimeColumn("마지막인상", format="YYYY-MM-DD"),
                         })
            st.download_button("📥 XLSX", _to_xlsx(df, "이미침식"), "마진침식_이미.xlsx", key="a_dl")

# ── 탭 B: 곧 침식 (2C) ──
with tabB:
    st.caption("실입고가는 올랐는데 상품관리(master) 매입가는 아직 옛 값 → 곧 마진이 떨어질 상품. "
               "(이미 master를 올린 건 탭A에서 잡힙니다.)")
    buyin = _buyin_recent(int(max(months_buy, 6)))
    pc = _price_changes()
    if buyin.empty or not master_price:
        st.warning("매입현황(buyin) 또는 product_master를 불러올 수 없습니다.")
    else:
        raises = me.recent_buy_raises(pc, months=int(months_buy), now=pd.Timestamp.now()) if not pc.empty else {}
        res = me.pending_buyin_raises(buyin, master_price, raises, months=6,
                                      min_pct=float(min_pct), max_pct=0.6, now=pd.Timestamp.now())
        alerts, suspect = res["alerts"], res["suspect"]
        rows = []
        for code, i in alerts.items():
            v = by_code.get(code, 0.0)
            rows.append({"관리코드": code, "상품명": name_lookup.get(code, ""),
                         "master매입": i["master매입"], "실입고가": i["실입고가"], "입고Δ%": i["입고Δ%"],
                         "월낱개판매": v, "월손실액": (i["실입고가"] - i["master매입"]) * v,
                         "입고일": i["입고일"]})
        bdf = pd.DataFrame(rows)
        if bdf.empty:
            st.success("✅ 실입고가가 master를 유의하게 추월한(미반영) 상품이 없습니다.")
        else:
            bdf = bdf.sort_values("월손실액", ascending=False).reset_index(drop=True)
            c = st.columns(3)
            c[0].metric("곧 침식 경보", f"{len(bdf)} 건")
            c[1].metric("월 손실액 합", f"{bdf['월손실액'].sum():,.0f} 원")
            c[2].metric("검토 필요", f"{len(suspect)} 건", help="Δ%>60% — 관리코드 충돌/단위/극과거 master 의심")
            _b = bdf.copy(); _b["입고Δ%"] = (_b["입고Δ%"] * 100).round(0)
            st.dataframe(_b, hide_index=True, use_container_width=True, height=440,
                         column_config={
                             "master매입": st.column_config.NumberColumn(format="%d"),
                             "실입고가": st.column_config.NumberColumn(format="%d"),
                             "입고Δ%": st.column_config.NumberColumn("실입고↑%", format="%.0f"),
                             "월낱개판매": st.column_config.NumberColumn(format="%d"),
                             "월손실액": st.column_config.NumberColumn("예상 월손실(원)", format="%d",
                                          help="(실입고−master) × 월 판매낱개 — master 반영 시 줄어들 마진"),
                             "입고일": st.column_config.DatetimeColumn(format="YYYY-MM-DD"),
                         })
            st.download_button("📥 XLSX", _to_xlsx(bdf, "곧침식2C"), "마진침식_곧2C.xlsx", key="b_dl")
            if suspect:
                with st.expander(f"⚠️ 검토 필요 {len(suspect)}건 (비현실적 배수 — 단위/코드 의심)"):
                    sdf = pd.DataFrame([{"관리코드": k, "상품명": name_lookup.get(k, ""),
                                         "master매입": v["master매입"], "실입고가": v["실입고가"],
                                         "입고Δ%": round(v["입고Δ%"] * 100)} for k, v in suspect.items()])
                    st.dataframe(sdf, hide_index=True, use_container_width=True)

# ── 탭 C: 실판매 이상 ──
with tabC:
    st.caption("실제 팔린 거래의 마진(매출자료·정산 진실) — 역마진 또는 채널 안전마진보다 크게 미달. 합포·소분 포함.")
    sales = _sales_recent(int(months_buy))
    bdict = _baseline_dict()
    if sales.empty:
        st.warning("매출자료(sales)를 불러올 수 없습니다.")
    else:
        anom = me.sales_margin_anomalies(sales, bdict, months=int(months_buy),
                                         buffer=float(buffer), now=pd.Timestamp.now())
        cdf = pd.DataFrame(anom)
        if cdf.empty:
            st.success("✅ 최근 실판매에서 역마진·기준 미달 상품이 없습니다.")
        else:
            cdf = cdf.sort_values(["역마진", "판매금액"], ascending=[False, False]).reset_index(drop=True)
            c = st.columns(4)
            c[0].metric("이상 건", f"{len(cdf)} 건")
            c[1].metric("역마진", f"{int(cdf['역마진'].sum())} 건")
            c[2].metric("기준 미달", f"{int((~cdf['역마진'] & cdf['기준마진율'].notna()).sum())} 건")
            c[3].metric("관련 매출", f"{cdf['판매금액'].sum():,.0f} 원")
            _c = cdf.copy()
            for col in ("마진율", "기준마진율", "미달폭"):
                _c[col] = (_c[col].astype(float) * 100).round(1)
            st.dataframe(_c, hide_index=True, use_container_width=True, height=440,
                         column_config={
                             "판매금액": st.column_config.NumberColumn(format="%d"),
                             "판매이익": st.column_config.NumberColumn(format="%d"),
                             "수량": st.column_config.NumberColumn("낱개", format="%d"),
                             "마진율": st.column_config.NumberColumn("실현마진%", format="%.1f"),
                             "기준마진율": st.column_config.NumberColumn("기준%", format="%.1f"),
                             "미달폭": st.column_config.NumberColumn("미달%p", format="%.1f"),
                             "역마진": st.column_config.CheckboxColumn("역마진"),
                         })
            st.download_button("📥 XLSX", _to_xlsx(cdf, "실판매이상"), "마진침식_실판매.xlsx", key="c_dl")
