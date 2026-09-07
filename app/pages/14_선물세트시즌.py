"""🎁 선물세트 시즌 — 추석/설 시즌 관제판(진도·재고·채널·가격).

구 VBA 워크북 `스마트스토어_마진율_계산기.xlsm` **선물세트 시트** 이관.
정본 = workflows/giftset-season-ops.md · 목표 체계 = ADR 0030(오픈마켓 15억·인하 재원 2,200만).

구 시트에서 바로잡은 것:
- 34행 하드코딩 → 자동 산출(선물세트 유니버스 ∩ (박스재고>0 ∪ 시즌매출>0)).
- 사장 채널(멸치·11번가) 제거, 자사몰 편입. SUMIFS 참조행 어긋남·#REF! 소멸.
- W~Z 금액/수량 단위 혼용 → 세트 단위 재구축(소진예측일·이월 잔여세트).
- 이익·마진·YoY·D-day 진도 밴드 신설.
"""
import io
import sys
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import pandas as pd
import streamlit as st

from core import ui
from core.dashboard import store
from core.intelligence import giftset_season as gsn
from core.workflows import channel_margin_monitor as cmm

_APP_API = "https://api.github.com/repos/OURPROJECTDAO/work-automation-app/contents"
_REF = Path(__file__).parent.parent.parent / "reference"
_KST = None
_CFG_PATH = "reference/giftset_season.json"


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
def _ref_csv(name: str) -> pd.DataFrame:
    code, text = _gh_raw(f"reference/{name}")
    if code != 200 or not text:
        return pd.DataFrame()
    return pd.read_csv(io.BytesIO(text), dtype=str)


@st.cache_data(ttl=600, show_spinner=False)
def _division() -> dict:
    code, text = _gh_raw("reference/internet_division_stores.csv")
    return gsn.load_division(text if code == 200 else None)


@st.cache_data(ttl=600, show_spinner=False)
def _config() -> dict:
    code, text = _gh_raw(_CFG_PATH)
    return gsn.load_config(text if code == 200 else None)


@st.cache_data(ttl=60, show_spinner=False)
def _part_index() -> dict:
    """master/ 파티션의 {YYYY-MM: blob sha}. ★sha를 캐시 키에 넣어야 '월 목록은 그대로인데
    내용만 갱신된' 경우(다른 페이지에서 그 달을 재적재)를 잡는다. 이게 없으면 TTL 만료 전까지
    옛 파케이를 그대로 내준다(2026-09-07 실사고)."""
    pat, repo = _data_secret()
    if not pat:
        return {}
    try:
        req = urllib.request.Request(
            f"https://api.github.com/repos/{repo}/contents/master?ref=main",
            headers={"Authorization": f"Bearer {pat}",
                     "Accept": "application/vnd.github+json"},
        )
        import json as _json
        items = _json.load(urllib.request.urlopen(req))
    except Exception:  # noqa: BLE001
        return {}
    out = {}
    for it in items:
        n = it.get("name", "")
        if n.startswith("sales_") and n.endswith(".parquet"):
            out[n[6:13]] = it.get("sha", "")
    return out


@st.cache_data(ttl=600, show_spinner="매출자료 불러오는 중...")
def _sales(months: tuple, _index: tuple) -> pd.DataFrame:
    """data repo master/sales_YYYY-MM.parquet 지정 월 로드.
    `_index` = (월, sha) 튜플 — 캐시 무효화 전용 키(값 자체는 안 씀)."""
    pat, repo = _data_secret()
    if not pat:
        return pd.DataFrame()
    avail = {m for m, _ in _index}
    parts = []
    for ym in months:
        if ym not in avail:
            continue
        p = store.read_partition(pat, repo, ym)
        if p is not None and len(p):
            parts.append(p)
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def _load_sales(a, b) -> pd.DataFrame:
    months = _months_between(a, b)
    idx = _part_index()
    return _sales(months, tuple(sorted((m, idx.get(m, "")) for m in months)))


def _months_between(a, b) -> tuple:
    return tuple(d.strftime("%Y-%m")
                 for d in pd.date_range(pd.Timestamp(a).replace(day=1),
                                        pd.Timestamp(b), freq="MS"))


def _nfc(v):
    return gsn._nfc(v)


def _won(v, unit="원"):
    if pd.isna(v):
        return "-"
    if abs(v) >= 1e8:
        return f"{v/1e8:,.2f}억"
    return f"{v:,.0f}{unit}"


def _to_xlsx(df: pd.DataFrame, title: str) -> bytes:
    from openpyxl import Workbook

    def _cell(v):
        if isinstance(v, float) and pd.isna(v):
            return None
        if hasattr(v, "isoformat"):
            return v.isoformat()
        if hasattr(v, "item"):
            return v.item()
        return v

    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
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


# ── 가격변경 글루 (0b_데일리대시보드 패턴 재사용) ────────────────────────────
@st.cache_data(ttl=600, show_spinner=False)
def _baseline_text():
    code, text = _gh_raw("reference/baseline_margin.csv")
    return text.decode("utf-8-sig") if code == 200 and text else None


def _cmm_key(channel):
    if channel in cmm.CHANNEL_CONFIG:
        return channel
    low = str(channel).lower()
    for k in cmm.CHANNEL_CONFIG:
        if k.lower() == low:
            return k
    return channel


def _supports_price_change(cfg) -> bool:
    if cfg.get("price_form"):
        return True
    cols = cfg.get("cols") or {}
    return ("즉시할인" in cols) and not cfg.get("consolidate")


@st.cache_data(ttl=600, show_spinner="채널 권장가 불러오는 중...")
def _cmm_listing(channel: str):
    ck = _cmm_key(channel)
    cfg = cmm.CHANNEL_CONFIG.get(ck)
    if not cfg:
        return None
    code, text = _gh_raw(f"reference/listing_{cfg['key']}.csv")
    if code != 200 or not text:
        return None
    recs = cmm.csv_text_to_recs(text.decode("utf-8"))
    bl = _baseline_text()
    override = cmm.parse_baseline_dict(bl) if bl else None
    rows, _ = cmm.compute_listing(recs, ck, str(_REF), baseline_override=override)
    return recs, rows


def _gen_price_form(channel, cfg, pf, recs, rows, pids):
    err = "권장가 산출 가능 항목이 없습니다(미매칭/기준 미설정)."
    miss = ("{} 원본 다운로드가 없습니다. 채널마진모니터에서 "
            "'상품관리 갱신 → 전체 교체'를 1회 실행하세요.")
    try:
        mode = (pf or {}).get("mode")
        prev = []
        if mode == "append":
            items, prev, _sk = cmm.build_append_items(pf, rows, recs, pids)
            if not items:
                return {"channel": channel, "error": err}
            out = cmm.build_price_form_append((_REF / pf["template"]).read_bytes(), items, pf)
        elif mode == "filter":
            rc, raw = _gh_raw(f"reference/listing_{cfg['key']}.xlsx")
            if rc != 200 or not raw:
                return {"channel": channel, "error": miss.format(channel)}
            out, prev, _sk, _ms = cmm.build_filter_price_xlsx(raw, rows, pids, cfg)
            if not prev:
                return {"channel": channel, "error": err}
        elif mode == "simple":
            out, prev, _sk = cmm.build_simple_price_xlsx(rows, pids, cfg, channel)
            if not prev:
                return {"channel": channel, "error": err}
        elif mode in ("multi_filter", "csv_filter"):
            ext = "csv" if mode == "csv_filter" else "xlsx"
            path = (f"reference/listing_{cfg['key']}_raw.csv" if ext == "csv"
                    else f"reference/listing_{cfg['key']}.xlsx")
            rc, raw = _gh_raw(path)
            if rc != 200 or not raw:
                return {"channel": channel, "error": miss.format(channel)}
            build = (cmm.build_multi_filter_xlsx if mode == "multi_filter"
                     else cmm.build_csv_filter_xlsx)
            out, prev, _sk, _ms = build(raw, rows, pids, cfg)
            if not prev:
                return {"channel": channel, "error": err}
        else:
            new_prices, _sk = cmm.compute_new_prices(rows, recs, set(pids))
            if not new_prices:
                return {"channel": channel, "error": err}
            rc, raw = _gh_raw(f"reference/listing_{cfg['key']}.xlsx")
            if rc != 200 or not raw:
                return {"channel": channel, "error": miss.format(channel)}
            out, _kept, _ms = cmm.build_bulk_price_xlsx(raw, new_prices, cfg)
            rb = {r["상품번호"]: r for r in recs}
            ro = {r["상품번호"]: r for r in rows}
            prev = [{"상품명": ro[p]["상품명"], "현재판매가": int(rb[p]["판매가"]),
                     "새판매가": v[0], "권장가": ro[p]["권장가"]} for p, v in new_prices.items()]
        return {"channel": channel, "bytes": out, "preview": prev,
                "name": f"선물세트_{channel}_가격변경_{datetime.now():%Y%m%d}.xlsx"}
    except Exception as e:  # noqa: BLE001
        return {"channel": channel, "error": f"생성 오류: {e}"}


def _append_division(names: list, grp: str):
    """화이트리스트 CSV에 상호명을 추가하고 커밋. 기존 행은 건드리지 않는다."""
    import base64
    import csv
    import json as _json
    path = "reference/internet_division_stores.csv"
    api = f"{_APP_API}/{path}"
    hdr = {"Authorization": f"Bearer {_pat()}", "Accept": "application/vnd.github+json",
           "Content-Type": "application/json"}
    try:
        meta = _json.load(urllib.request.urlopen(urllib.request.Request(api, headers=hdr)))
        code, body = _gh_raw(path)
        if code != 200:
            return False, "기존 CSV를 읽지 못했습니다."
        text = body.decode("utf-8-sig")
        rows = list(csv.DictReader(io.StringIO(text)))
        have = {gsn._nfc(r.get("상호명")) for r in rows}
        cols = list(rows[0].keys()) if rows else ["상호명", "구분", "첫확인", "출처", "비고"]
        added = 0
        today = datetime.now().strftime("%Y-%m-%d")
        for n in names:
            if gsn._nfc(n) in have:
                continue
            rows.append({**{c: "" for c in cols}, "상호명": n, "구분": grp,
                         "첫확인": today, "출처": "앱 등록(미분류 편입)"})
            added += 1
        if not added:
            return False, "이미 전부 등록되어 있습니다."
        buf = io.StringIO()
        wcsv = csv.DictWriter(buf, fieldnames=cols)
        wcsv.writeheader()
        wcsv.writerows(rows)
        payload = _json.dumps({
            "message": f"ref(giftset): 인터넷사업부 화이트리스트 +{added}곳 ({grp})",
            "content": base64.b64encode(buf.getvalue().encode()).decode(),
            "sha": meta["sha"],
        }).encode()
        urllib.request.urlopen(urllib.request.Request(api, data=payload,
                                                      method="PUT", headers=hdr))
        return True, f"{added}곳을 `{grp}` 로 등록했습니다."
    except Exception as e:  # noqa: BLE001
        return False, f"등록 실패: {e}"


# ═════════════════════════════════════════════════════════════════════════════
ui.page_header("선물세트 시즌", icon="🎁")
cfg = _config()

with st.sidebar:
    st.subheader("⚙️ 시즌 설정")
    st.caption(f"**{cfg['시즌명']}** · D-day {cfg['dday']}")
    base_day = st.date_input("기준일", value=pd.Timestamp.now().date())
    st.caption("다른 페이지(데이터현황)에서 적재했는데 반영이 안 되면 아래를 누르세요.")
    if st.button("🔄 데이터 새로고침"):
        st.cache_data.clear()
        st.rerun()
    st.caption(f"측정창 D-{cfg.get('측정창_일수',55)} · 운영창 D-{cfg.get('운영창_일수',31)}")
    st.caption(f"목표 {float(cfg['매출목표'])/1e8:.1f}억(오픈마켓) · "
               f"계획마진 {float(cfg['계획마진'])*100:.1f}% · "
               f"하한 {float(cfg['마진하한'])*100:.1f}% · "
               f"절대하한 {float(cfg['절대하한'])*100:.0f}%")
    st.caption(f"설정 파일: `{_CFG_PATH}` (없으면 기본값)")

win = gsn.season_window(cfg, base_day)
D = win["D"]

if not _data_secret()[0]:
    st.warning("매출 데이터 접근용 PAT(st.secrets [data] 또는 GITHUB_PAT)가 없습니다.")
    st.stop()

# 비시즌 안내 (D-day 15일 경과 후 ~ 시즌창 개시 전)
_open_d = int(cfg.get("운영창_일수", 31))
if D > _open_d or D < -15:
    st.info(
        f"**지금은 시즌 밖입니다.** ({cfg['시즌명']} D-day `{cfg['dday']}` 기준 D{D:+d})\n\n"
        f"시즌 트랙은 **운영창 D-{_open_d} 개시 ~ D+15 정산**까지만 엽니다"
        f"(매출 누적 범위인 **측정창 D-{cfg.get('측정창_일수', 55)}** 와 다릅니다). "
        f"다음 명절을 시작하려면 `{_CFG_PATH}` 의 `시즌명`·`dday`·`작년_dday`·`매출목표`를 "
        "갈아끼우세요. 나머지 로직(유니버스 자동 산출·밴드·판정)은 그대로 재사용됩니다."
    )
    st.caption("정본: workflows/giftset-season-ops.md · 목표 체계 ADR 0030")
    st.stop()

pm = _ref_csv("product_master.csv")
attrs = _ref_csv("product_attributes.csv")
lineup = _ref_csv("giftset_lineup_2026.csv")
if pm.empty:
    st.error("product_master를 불러오지 못했습니다.")
    st.stop()

cur = _load_sales(win["start"], win["end"])
prev = _load_sales(win["ly_start"], win["ly_end"])
if cur.empty:
    st.warning("시즌창에 해당하는 매출 파티션이 없습니다. **[데이터 적재]** 탭에서 "
               "영업이익현황을 올리세요.")

division = _division()
res = gsn.build_board(pm, attrs, lineup, cur, prev, cfg, now=base_day, division=division)
b, k = res["board"], res["kpi"]
if b.empty:
    st.info("표시할 선물세트가 없습니다(재고 0 · 시즌매출 0).")
    st.stop()

# ── 상단 KPI ────────────────────────────────────────────────────────────────
band = k["밴드"]
c = st.columns(5)
c[0].metric(f"D-{D}", f"판매 잔여 {k['잔여일']}일", help=f"D-{cfg.get('마감_D',5)} 배송 마감 기준")
c[1].metric("오픈마켓 누적", _won(k["오픈마켓매출"]),
            f"{k['밴드판정']} 진도 {k['진도']*100:.2f}%")
c[2].metric("마진", f"{k['마진']*100:.2f}%",
            f"{k['적립']:+,.0f}원 (계획 {float(cfg['계획마진'])*100:.1f}% 대비)")
c[3].metric("인하 재원 잔액", _won(k["재원잔액"]),
            f"총 재원 {_won(k['재원'])}", delta_color="off")
def _pct(v):
    return "-" if v is None or pd.isna(v) else f"{v*100:.1f}%"


c[4].metric("YoY (7일 블록)", _pct(k["YoY블록"]),
            f"직전 블록 {_pct(k['YoY블록_직전'])}", delta_color="off")

if band:
    msg = (f"D-{D} 진도 밴드 **{band[0]*100:.0f}~{band[1]*100:.0f}%** · "
           f"현재 **{k['진도']*100:.2f}%** ({_won(k['오픈마켓매출'])} / 목표 "
           f"{float(cfg['매출목표'])/1e8:.1f}억)")
    (st.error if k["밴드판정"] == "🔴" else
     st.success if k["밴드판정"] == "🟢" else st.info)(f"{k['밴드판정']} {msg}")
st.caption(
    f"**진도 정본은 7일 블록 YoY** — 같은 D 구간을 주말 포함 통째로 잘라 비교하므로 "
    f"요일 효과가 상쇄됩니다. 참고: 누적 YoY 달력 {_pct(k['YoY달력'])} "
    f"(`{k['작년달력'].date()}`) / 영업일 정렬 {_pct(k['YoY영업일'])} "
    f"(`{k['작년영업일'].date()}`, 잔여 영업일 {k['잔여영업일']}일 기준). "
    "요일 배치가 어긋나는 구간에서 누적 YoY만 보면 오독합니다. "
    "목표·밴드는 **오픈마켓 한정**이며 리테일·자사몰·나들·오프라인은 참고 집계입니다."
)

_last = pd.to_datetime(cur["거래일자"], errors="coerce").max() if len(cur) else None
if _last is not None and pd.notna(_last):
    _gap = (pd.Timestamp(base_day) - _last.normalize()).days
    _txt = f"📥 적재된 매출 최신 거래일 **{_last:%Y-%m-%d}** (기준일 대비 {_gap}일 전)"
    (st.caption if _gap <= 1 else st.warning)(
        _txt if _gap <= 1 else _txt + " — 최근 매출이 빠져 있습니다. "
        "**[데이터 적재]** 탭에서 올리거나, 다른 페이지에서 적재했다면 사이드바 "
        "**🔄 데이터 새로고침**을 누르세요.")

dc = st.columns(4)
dc[0].metric("인터넷사업부 계", _won(k["사업부매출"]),
             f"마진 {k['사업부마진']*100:.2f}%", delta_color="off")
dc[1].metric("사업부 진도", "-" if pd.isna(k["사업부진도"]) else f"{k['사업부진도']*100:.1f}%",
             f"기준선 {_won(k['사업부기준선'])} ({k['사업부기준선출처']})", delta_color="off")
dc[2].metric("사업부 YoY", _pct(k["사업부YoY"]), "작년 동기(달력)", delta_color="off")
dc[3].metric("오픈마켓 비중",
             f"{k['오픈마켓매출']/k['사업부매출']*100:.1f}%" if k["사업부매출"] else "-",
             "목표 스코프가 차지하는 몫", delta_color="off")
st.caption(
    "**인터넷사업부 = 주요 채널(오픈마켓·자사몰) + 나들 + 온라인 B2B.** 오프라인 영업부는 "
    "제외합니다 — `master/sales_*.parquet` 은 **전사 데이터**라 "
    "`reference/internet_division_stores.csv` 화이트리스트로 걸러냅니다. "
    "⚠️ **작년 온라인 B2B는 가를 수 없어 사업부 YoY·진도는 과소 집계**입니다"
    "(작년 시즌창 354곳 중 화이트리스트 히트 18곳뿐). 작년 인터넷사업부 export가 있어야 정확해집니다."
)

_unk = k.get("미분류")
if _unk:
    with st.expander(f"⚠️ 화이트리스트 미등재 {_unk['곳']}곳 / {_won(_unk['매출'])} — "
                     "사업부 안팎 미판정", expanded=False):
        st.caption("온라인 B2B는 1회성이라 매 시즌 새 상호명이 생깁니다. 오프라인으로 "
                   "단정하지 않고 여기 모아둡니다. **현재 어느 집계에도 포함되지 않습니다.** "
                   "인터넷사업부 export를 [데이터 적재] 탭에 올리면 자동으로 편입됩니다.")
        st.dataframe(_unk["목록"].rename("매출").reset_index(), hide_index=True, width="stretch")
        _names = list(_unk["목록"].index)
        _pick = st.multiselect("화이트리스트에 편입할 상호명", _names, default=_names,
                               key="gs_unk_pick")
        _grp = st.radio("구분", ["온라인B2B", "사업부밖", "주요-오픈마켓", "주요-자사몰",
                                "나들", "제외"], horizontal=True, key="gs_unk_grp")
        st.caption("`사업부밖` = 오프라인 확정(관제판 집계에서 빠집니다). "
                   "측정창이 인터넷사업부 export 커버 기간 안이라면, export에 없는 상호명은 "
                   "오프라인으로 확정해도 됩니다.")
        if _pick and st.button(f"📝 {len(_pick)}곳을 `{_grp}` 로 등록",
                               type="primary", key="gs_unk_go"):
            _ok, _msg = _append_division(_pick, _grp)
            (st.success if _ok else st.error)(_msg)
            if _ok:
                st.cache_data.clear()
                st.rerun()

tabs = st.tabs(["📊 진도판", "📋 관제판", "📦 재고 경보", "🛠️ 가격변경", "⬆️ 데이터 적재"])

# ── 진도판 ──────────────────────────────────────────────────────────────────
with tabs[0]:
    ui.section_head("채널군", icon="🧭")
    g = res["group"].copy()
    g["매출"] = g["매출"].map(lambda v: f"{v:,.0f}")
    g["이익"] = g["이익"].map(lambda v: f"{v:,.0f}")
    g["마진"] = g["마진"].map(lambda v: f"{v*100:.2f}%")
    st.dataframe(g, hide_index=True, width="stretch")

    ui.section_head("오픈마켓 채널별", icon="🛒")
    sc = res["scoped"]
    op = sc[sc["채널군"] == gsn.GRP_OPEN]
    if len(op):
        ch = op.groupby("채널").agg(매출=("판매금액", "sum"), 이익=("판매이익", "sum"),
                                  세트=("수량", "sum"))
        ch["마진"] = ch["이익"] / ch["매출"]
        ch = ch.sort_values("매출", ascending=False).reset_index()
        st.dataframe(ch.style.format({"매출": "{:,.0f}", "이익": "{:,.0f}",
                                      "세트": "{:,.0f}", "마진": "{:.2%}"}),
                     hide_index=True, width="stretch")

    ui.section_head("일별 추이 (주요 채널)", icon="📈")
    on = sc[sc["채널군"].isin(gsn.MAIN_GROUPS)]
    if len(on):
        dd = on.groupby(on["거래일자"].dt.date)["판매금액"].sum().tail(21)
        st.bar_chart(dd, height=220)
        st.caption("주말은 발주를 돌리지 않아 0으로 잡히고 월요일이 흡수합니다. "
                   "일 단위가 아니라 **7일 블록**으로 보세요.")

# ── 관제판 ──────────────────────────────────────────────────────────────────
with tabs[1]:
    f = st.columns([2, 2, 3, 2])
    only_line = f[0].checkbox("시즌 라인업만", value=True,
                              help="끄면 라인업 밖 선물세트(구 코드·시즌 비운영)도 표시")
    vmode = f[1].radio("채널 값", ["매출", "이익"], horizontal=True)
    verdicts = f[2].multiselect("판정", [gsn.V_CUT, gsn.V_BUY, gsn.V_HOLD,
                                        gsn.V_UP, gsn.V_NOSALE])
    q = f[3].text_input("검색", placeholder="코드·상품명")

    v = b.copy()
    if only_line:
        v = v[v["시즌라인업"]]
    if verdicts:
        v = v[v["판정"].isin(verdicts)]
    if q.strip():
        s = q.strip()
        v = v[v["관리코드"].str.contains(s, case=False, na=False)
              | v["상품명"].str.contains(s, case=False, na=False)]

    if vmode == "이익":
        pv = res.get("이익피벗")
        if pv is not None and len(pv):
            for cc in gsn.channel_columns():
                v[cc] = v["관리코드"].map(pv[cc] if cc in pv.columns else pd.Series(dtype=float)).fillna(0.0)

    cols = (["관리코드", "상품명", "박스재고", "세트재고", "재고금액",
             "시즌매출", "시즌이익", "마진율", "판매세트",
             "주요채널계", "온라인B2B계", "사업부계", "일평균세트7", "소진예측일", "잔여시즌일", "마감후잔여세트",
             "작년동기_달력", "YoY", "작년시즌", "작년진도", "판정", "결손"]
            + gsn.channel_columns())
    view = v[cols].sort_values("시즌매출", ascending=False)
    st.dataframe(
        view.style.format({
            "박스재고": "{:,.0f}", "세트재고": "{:,.0f}", "재고금액": "{:,.0f}",
            "시즌매출": "{:,.0f}", "시즌이익": "{:,.0f}", "마진율": "{:.2%}",
            "판매세트": "{:,.0f}", "주요채널계": "{:,.0f}",
            "온라인B2B계": "{:,.0f}", "사업부계": "{:,.0f}",
            "일평균세트7": "{:,.1f}", "소진예측일": "{:,.1f}",
            "마감후잔여세트": "{:,.0f}", "작년동기_달력": "{:,.0f}", "YoY": "{:.1%}",
            "작년시즌": "{:,.0f}", "작년진도": "{:.1%}",
            **{c: "{:,.0f}" for c in gsn.channel_columns()},
        }),
        hide_index=True, width="stretch", height=560,
    )
    st.caption(f"{len(view)}행 · 채널 컬럼 = **{vmode}**. "
               "`소진예측일` = 세트재고 ÷ 최근 7일 일평균 판매세트. "
               "`마감후잔여세트` = 이 속도로 D-5까지 팔고 남는 세트(이월 물량).")
    st.download_button("⬇️ 관제판 XLSX", _to_xlsx(view, "선물세트시즌"),
                       f"선물세트_관제판_{pd.Timestamp(base_day):%Y%m%d}.xlsx")

    bad = b[b["결손"] != ""]
    if len(bad):
        st.warning(f"⚠️ 기준데이터 결손 {len(bad)}건 — 유니버스 두 소스"
                   "(`product_attributes` / `giftset_lineup_2026.csv`) 중 한쪽에만 있는 코드입니다.")
        st.dataframe(bad[["관리코드", "상품명", "결손", "박스재고", "재고금액", "시즌매출"]],
                     hide_index=True, width="stretch")
    neg = b[b["재고음수"]]
    if len(neg):
        st.caption(f"ℹ️ 박스재고 음수 {len(neg)}건 — 시즌 중 매입 전표 지연으로 흔합니다. "
                   "출고 차단 사유가 아니며 소진예측에서는 0으로 클램프했습니다.")

# ── 재고 경보 ───────────────────────────────────────────────────────────────
with tabs[2]:
    carry = float(cfg.get("이월비용률", 0.033))
    st.caption(
        f"**인하 vs 이월 손익분기** — 이월 비용률 {carry*100:.1f}%. "
        f"D-{cfg.get('마감_D',5)}까지 못 빼는 물량은 이월 비용이 붙으므로, "
        "인하폭이 손익분기(약 3%p) 안이면 **지금 인하해서 파는 쪽**이 낫습니다."
    )
    left = b[(b["마감후잔여세트"] > 0) & (b["세트재고"] > 0)].copy()
    left["이월비용"] = left["이월재고금액"] * carry
    left = left.sort_values("이월재고금액", ascending=False)
    ui.section_head(f"이월 위험 — {len(left)}종 / 이월 예상 {_won(left['이월재고금액'].sum())}",
                    icon="⚪")
    if len(left):
        st.dataframe(
            left[["관리코드", "상품명", "세트재고", "일평균세트7", "소진예측일",
                  "잔여시즌일", "마감후잔여세트", "이월재고금액", "이월비용",
                  "마진율", "판정"]].style.format({
                      "세트재고": "{:,.0f}", "일평균세트7": "{:,.1f}", "소진예측일": "{:,.1f}",
                      "마감후잔여세트": "{:,.0f}", "이월재고금액": "{:,.0f}",
                      "이월비용": "{:,.0f}", "마진율": "{:.2%}"}),
            hide_index=True, width="stretch")

    short = b[(b["세트재고"] > 0) & (b["시즌매출"] > 0)
              & (b["소진예측일"] < b["잔여시즌일"])].copy()
    short = short.sort_values("소진예측일")
    ui.section_head(f"품절 임박 — {len(short)}종 (잔여 시즌일 내 소진)", icon="🔴")
    if len(short):
        st.dataframe(
            short[["관리코드", "상품명", "세트재고", "일평균세트7", "소진예측일",
                   "잔여시즌일", "시즌매출", "마진율", "판정"]].style.format({
                       "세트재고": "{:,.0f}", "일평균세트7": "{:,.1f}",
                       "소진예측일": "{:,.1f}", "시즌매출": "{:,.0f}", "마진율": "{:.2%}"}),
            hide_index=True, width="stretch")
        st.caption("잔여 시즌일 안에 재고가 바닥납니다. 추가 매입 가능 여부(CJ·동원 시즌 생산 마감) "
                   "확인 대상입니다.")

    gone = b[(b["세트재고"] <= 0) & (b["시즌매출"] > 0)]
    if len(gone):
        ui.section_head(f"재고 소진 — {len(gone)}종 (팔리는데 재고 0 이하)", icon="⛔")
        st.dataframe(gone[["관리코드", "상품명", "박스재고", "세트재고", "시즌매출",
                           "일평균세트7", "재고음수"]].style.format({
                               "박스재고": "{:,.0f}", "세트재고": "{:,.0f}",
                               "시즌매출": "{:,.0f}", "일평균세트7": "{:,.1f}"}),
                     hide_index=True, width="stretch")
        st.caption("`재고음수 = True` 는 매입 전표 지연일 가능성이 높습니다(시즌 중 흔함). "
                   "전표가 아니라 실물이 없는 것이면 즉시 판매중지 대상입니다.")

    ui.section_head("주요채널 무매출 · 저마진", icon="⚠️")
    watch = b[(b["주요채널계"] <= 0) | ((b["시즌매출"] > 0)
                                    & (b["마진율"] < float(cfg["절대하한"])))]
    if len(watch):
        st.dataframe(
            watch[["관리코드", "상품명", "박스재고", "재고금액", "주요채널계", "온라인B2B계", "시즌매출",
                   "마진율", "작년시즌", "판정"]].sort_values(
                       "재고금액", ascending=False).style.format({
                           "박스재고": "{:,.0f}", "재고금액": "{:,.0f}",
                           "주요채널계": "{:,.0f}", "온라인B2B계": "{:,.0f}",
                           "시즌매출": "{:,.0f}",
                           "마진율": "{:.2%}", "작년시즌": "{:,.0f}"}),
            hide_index=True, width="stretch")
        st.caption("`주요채널계 = 0` 은 오픈마켓·자사몰엔 안 나가고 나들·온라인B2B로만 "
                   "빠지고 있다는 뜻입니다(등재·노출 점검 대상). "
                   f"마진 {float(cfg['절대하한'])*100:.0f}% 미만은 절대하한 관통입니다. "
                   "나들은 floor anchor라 낮은 게 정상이지만 **0%·역마진은 정상이 아닙니다**.")
    else:
        st.success("무매출·절대하한 미달 없음.")

# ── 가격변경 ────────────────────────────────────────────────────────────────
with tabs[3]:
    st.caption("행을 고르고 채널을 선택하면 **채널마진모니터 빌더**로 그 채널 가격변경 시트를 "
               "만듭니다. 권장가는 기준마진율 기반이며 3% 절대하한이 적용됩니다.")
    opts = (b.sort_values("시즌매출", ascending=False)
             .apply(lambda r: f"{r['관리코드']} · {r['상품명'][:34]}", axis=1).tolist())
    picked = st.multiselect("관리코드 선택", opts, key="gs_pick")
    codes = {p.split(" · ")[0] for p in picked}
    chans = [c for c in cmm.CHANNEL_CONFIG
             if _supports_price_change(cmm.CHANNEL_CONFIG[c])]
    channel = st.selectbox("채널", chans, key="gs_ch")

    if not codes:
        st.info("관리코드를 1개 이상 고르세요.")
    else:
        data = _cmm_listing(channel)
        if not data:
            st.warning(f"**{channel}** 저장 listing이 없습니다. "
                       "채널마진모니터에서 '상품관리 갱신'을 1회 실행하세요.")
        else:
            recs, rows = data
            pids = [r["상품번호"] for r in rows if _nfc(r.get("관리코드")) in codes]
            if not pids:
                st.warning("선택 상품이 해당 채널 listing에 없습니다(미등재이거나 listing 갱신 필요).")
            elif st.button(f"🛠️ {channel} 가격변경 시트 생성 — "
                           f"관리코드 {len(codes)}개 → listing {len(pids)}건",
                           type="primary", key="gs_pc_gen"):
                st.session_state["gs_pcform"] = _gen_price_form(
                    channel, cmm.CHANNEL_CONFIG[_cmm_key(channel)],
                    cmm.CHANNEL_CONFIG[_cmm_key(channel)].get("price_form"),
                    recs, rows, pids)

    form = st.session_state.get("gs_pcform")
    if form and form.get("channel") == channel:
        if form.get("error"):
            st.warning(form["error"])
        else:
            st.download_button(f"⬇️ {channel} 가격변경 시트 (.xlsx)", form["bytes"],
                               form["name"], type="primary", key="gs_pc_dl")
            if form.get("preview"):
                st.dataframe(pd.DataFrame(form["preview"]), hide_index=True, width="stretch")

# ── 데이터 적재 ─────────────────────────────────────────────────────────────
with tabs[4]:
    st.caption("천년경영 **영업이익현황**을 올리면 `master/sales_YYYY-MM.parquet` 에 "
               "바로 적재됩니다(같은 날짜 구간은 교체 — 멱등). 적재 후 관제판이 최신이 됩니다.")
    pat, repo = _data_secret()
    try:
        mons = store.list_partition_months(pat, repo)
        need = [m for m in _months_between(win["start"], win["end"]) if m not in mons]
        if need:
            st.warning("시즌창 중 **미적재 월**: " + ", ".join(need))
        else:
            st.success("시즌창 전 구간 적재됨: " + ", ".join(_months_between(win["start"], win["end"])))
        st.caption("적재된 월: " + ", ".join(mons[-8:]))
    except Exception as e:  # noqa: BLE001
        st.info(f"적재 현황을 읽지 못했습니다: {e}")

    up = st.file_uploader("영업이익현황 (.xlsx)", type=["xlsx"], key="gs_up")
    if up is not None:
        st.caption(f"업로드: **{up.name}**")
        if st.button("📥 파케이에 적재", type="primary", key="gs_ing"):
            with st.spinner("적재 중..."):
                try:
                    info = store.ingest(pat, repo, io.BytesIO(up.getvalue()))
                    st.cache_data.clear()
                    st.success(f"적재 완료 — {info}. 관제판을 다시 계산합니다.")
                    st.rerun()
                except Exception as e:  # noqa: BLE001
                    st.error(f"적재 실패: {e}")
    st.caption("적재는 **올린 파일의 날짜 구간만 교체**합니다(`date_range_replace`) — "
               "부분 기간 업로드도 기존 데이터를 지우지 않습니다. 다만 아직 파티션이 "
               "없는 달은 올린 구간만 생기므로, 새 달 첫 적재는 그 달 전체를 뽑으세요.")
