"""채널 가격·마진 모니터 (channel-margin-monitor).

저장된 상품관리(listing) 스냅샷을 자동 로드해 상품별 마진율·기준마진 대비 탐지·권장가(또는
제한 텍스트)·재고를 계산. 표에서 상품을 선택해 CSV 또는 '가격 일괄변경 양식'으로 내보낸다.
공식·근거 = KB workflows/channel-margin-monitor.md.
"""
import base64
import json
import sys
import urllib.request
from datetime import datetime, timedelta, timezone
from io import StringIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # repo root

import pandas as pd
import streamlit as st

from core.workflows import channel_margin_monitor as cmm
from core.intelligence import listing_history
from core.dashboard import store

_REF = Path(__file__).parent.parent.parent / "reference"
_APP_REPO = "OURPROJECTDAO/work-automation-app"
_KST = timezone(timedelta(hours=9))
_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.title("💹 채널 가격·마진 모니터")
st.caption(
    "저장된 상품관리 기준으로 마진율·기준마진 대비 이탈·권장가를 계산합니다. "
    "매입가/재고는 상품관리(product_master), 기준마진율은 baseline_margin 기준. "
    "표에서 선택한 상품을 CSV 또는 가격 일괄변경 양식으로 내보낼 수 있습니다."
)


def _pat() -> str:
    return st.secrets.get("GITHUB_PAT", "")


def _gh(path, method="GET", data=None, raw=False):
    url = f"https://api.github.com/repos/{_APP_REPO}/contents/{path}"
    req = urllib.request.Request(url, method=method)
    if _pat():
        req.add_header("Authorization", f"Bearer {_pat()}")
    req.add_header("Accept", "application/vnd.github.raw" if raw else "application/vnd.github+json")
    if data is not None:
        req.data = json.dumps(data).encode()
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req) as r:
            return r.status, (r.read().decode() if raw else json.load(r))
    except urllib.error.HTTPError as e:
        return e.code, (e.read().decode() if raw else json.loads(e.read().decode() or "{}"))


def _gh_bytes(path):
    """바이너리(원본 .xlsx) raw 다운로드 → (status, bytes)."""
    url = f"https://api.github.com/repos/{_APP_REPO}/contents/{path}"
    req = urllib.request.Request(url)
    if _pat():
        req.add_header("Authorization", f"Bearer {_pat()}")
    req.add_header("Accept", "application/vnd.github.raw")
    try:
        with urllib.request.urlopen(req) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, b""


def _listing_path(key): return f"reference/listing_{key}.csv"
def _meta_path(key): return f"reference/listing_{key}.meta.json"
def _raw_path(key): return f"reference/listing_{key}.xlsx"


@st.cache_data(ttl=600, show_spinner="저장된 상품관리 불러오는 중...")
def _load_listing(key: str):
    code, text = _gh(_listing_path(key), raw=True)
    if code != 200:
        return None, {}
    recs = cmm.csv_text_to_recs(text)
    mcode, mtext = _gh(_meta_path(key), raw=True)
    meta = json.loads(mtext) if mcode == 200 else {}
    return recs, meta


def _commit_listing(key: str, recs: list) -> dict:
    meta = {"updated_at": datetime.now(_KST).isoformat(timespec="seconds"), "rows": len(recs)}
    for path, body in [(_listing_path(key), cmm.recs_to_csv(recs)),
                       (_meta_path(key), json.dumps(meta, ensure_ascii=False))]:
        code, m = _gh(path)
        payload = {"message": f"data(listing): {key} 갱신 ({len(recs)}건)",
                   "content": base64.b64encode(body.encode("utf-8")).decode()}
        if code == 200:
            payload["sha"] = m["sha"]
        _gh(path, "PUT", payload)
    return meta


def _commit_raw(key: str, xlsx_bytes: bytes):
    """원본 일괄변경 양식(.xlsx) 저장 — 가격변경 양식 출력의 원천(전체 컬럼 보존)."""
    path = _raw_path(key)
    code, m = _gh(path)
    payload = {"message": f"data(listing-raw): {key} 원본양식 갱신",
               "content": base64.b64encode(xlsx_bytes).decode()}
    if code == 200:
        payload["sha"] = m["sha"]
    _gh(path, "PUT", payload)


def _data_secret():
    """(pat, repo) — private data repo(이력 엔진 1d). secrets [data] 우선, 없으면 GITHUB_PAT 폴백."""
    repo = "OURPROJECTDAO/work-automation-data"
    try:
        d = st.secrets["data"]
        return d["pat"], d.get("repo", repo)
    except Exception:
        return st.secrets.get("GITHUB_PAT", ""), repo


def _accumulate_listing(key: str, recs: list) -> None:
    """listing 커밋 직후 그 채널 가격을 날짜본으로 data repo에 적립(이력 1d). 비차단·best-effort.

    forward 축적(과거 소급 불가). 두뇌③ 채널 가격 A/B 가격변경 전후 토대. PII 없음(가격만).
    """
    pat, repo = _data_secret()
    if not pat:
        return  # data repo 미설정 — 조용히 건너뜀(이력 비활성)
    try:
        snap_date = datetime.now(_KST).date()
        snap = listing_history.snapshot_from_recs(recs, key, snap_date)
        res = listing_history.ingest_listing_snapshot(snap, pat, repo)
        st.toast(f"📚 listing 가격 스냅샷 적립 {snap_date} · {res['added']}행", icon="📚")
    except Exception as e:
        st.toast(f"⚠️ listing 스냅샷 적립 실패(저장은 완료): {e}", icon="⚠️")


_BASELINE_PATH = "reference/baseline_margin.csv"


@st.cache_data(ttl=600, show_spinner="기준마진율 불러오는 중...")
def _load_baseline_text() -> str:
    """baseline_margin.csv 를 GitHub 라이브로 읽음(로컬 배포본 대신) → 편집 즉시 반영."""
    code, text = _gh(_BASELINE_PATH, raw=True)
    return text if code == 200 else ""


def _commit_baseline(new_text: str):
    code, m = _gh(_BASELINE_PATH)
    payload = {"message": "data(baseline): 기준마진율 갱신(대시보드 편집)",
               "content": base64.b64encode(new_text.encode("utf-8")).decode()}
    if code == 200:
        payload["sha"] = m["sha"]
    _gh(_BASELINE_PATH, "PUT", payload)


_FLOOR_PATH = "reference/margin_floor.csv"


@st.cache_data(ttl=600, show_spinner=False)
def _load_floor_text() -> str:
    """margin_floor.csv 를 GitHub 라이브로 읽음 → 제한 등록/해제 즉시 반영."""
    code, text = _gh(_FLOOR_PATH, raw=True)
    return text if code == 200 else ""


def _commit_floor(new_text: str):
    code, m = _gh(_FLOOR_PATH)
    payload = {"message": "data(floor): 제한 상품 갱신(대시보드 편집)",
               "content": base64.b64encode(new_text.encode("utf-8")).decode()}
    if code == 200:
        payload["sha"] = m["sha"]
    _gh(_FLOOR_PATH, "PUT", payload)


@st.cache_data(ttl=600, show_spinner=False)
def _load_refs():
    """canonical_code용 reference(sobun·pm_by_prod 등). 가벼운 csv 로드."""
    return cmm.load_references(str(_REF))


# 채널마진모니터 채널키 → 매출자료(천년경영 정산) 상호명(들). 쿠팡=윙배송+로켓창고.
_CH_TO_SANGHO = {
    "스마트스토어": ["오픈마켓- 스마트스토어"],
    "esm": ["오픈마켓(ESM/옥션.지마켓.G9)"],
    "식봄": ["오픈마켓- (주) 마켓보로"],
    "캐시노트": ["오픈마켓 (주) 한국신용데이터"],
    "알리": ["오픈마켓- 알리"],
    "쿠팡": ["오픈마켓 쿠팡 (윙배송)", "쿠팡(로켓창고)"],
    "배민상회": ["오픈마켓- (주) 우아한형제들"],
    "올웨이즈": ["오픈마켓- 올웨이즈"],
}


@st.cache_data(ttl=600, show_spinner="전월 매출 불러오는 중...")
def _load_prev_sales():
    """적재 최신월 매출 파티션 → (ym, {코드: 전체매출}, {상호명: {코드: 매출}}). 박스코드 기준.

    매출자료는 박스코드라 정규화 불요 — listing 행만 canonical_code로 맞춰 lookup.
    """
    pat, repo = _data_secret()
    if not pat:
        return None, {}, {}
    try:
        months = store.list_partition_months(pat, repo)
        if not months:
            return None, {}, {}
        ym = months[-1]
        sdf = store.read_partition(pat, repo, ym)
    except Exception:
        return None, {}, {}
    if sdf is None or sdf.empty:
        return ym, {}, {}
    sdf = sdf.copy()
    sdf["_code"] = sdf["관리코드"].map(cmm._nfc)
    sdf["_sg"] = sdf["상호명"].astype(str).map(cmm._nfc)
    _managed = {cmm._nfc(s) for v in _CH_TO_SANGHO.values() for s in v}  # 내 관리 8채널(나들·B2B 제외)
    msdf = sdf[sdf["_sg"].isin(_managed)]
    total = {k: float(v) for k, v in msdf.groupby("_code")["판매금액"].sum().items()}
    by_ch: dict = {}
    for (sg, cd), v in sdf.groupby(["_sg", "_code"])["판매금액"].sum().items():
        by_ch.setdefault(sg, {})[cd] = float(v)
    return ym, total, by_ch


def _col_config(cfg: dict, prev_ym=None) -> dict:
    """표 컬럼 헤더에 채널별 수식 설명(help) 부여 — 사람이 한 번 더 검증.

    수수료·실택배비·배송비 출처·N 출처가 채널마다 달라 cfg에서 동적 생성.
    """
    per_item = cfg.get("commission_source") == "download"   # 상품별 수수료(배민상회)
    comm = cfg.get("commission")
    settle = cfg["ship_settle"]; ship = int(cfg["real_ship"])
    rate_txt = "(1−상품별수수료)" if per_item else f"{(1 - comm):.2f}"
    comm_txt = "상품별(다운로드 BU/100+추가)" if per_item else f"{comm*100:.0f}%"
    if cfg.get("ship_fee_const") is not None:
        ship_src = f"상수 {int(cfg['ship_fee_const']):,}원(다운로드에 배송비 숫자 없음)"
    elif cfg.get("ship_fee_policy"):
        sp = cfg["ship_fee_policy"]
        paid = ", ".join(f"{k}→{int(v):,}원" for k, v in sp["map"].items())
        ship_src = f"배송정책코드 조건부({paid}, 그 외→{int(sp.get('default', 0)):,}원)"
    else:
        ship_src = "다운로드 기본배송비"
    if cfg.get("n_source") == "ref":
        n_help = "판매단위 배수(합포량). hapo_multiplier에서 상품번호로 조회 · 빈값→1 · 분수 가능. 매입가에 ×N"
    else:
        n_help = "판매단위 배수(판매자바코드) · 빈값→1 · 분수 가능. 매입가에 ×N"
    NC, TC = st.column_config.NumberColumn, st.column_config.TextColumn
    return {
        "상품번호": TC("상품번호", help="채널 상품 고유키(선택·양식 키)"),
        "관리코드": TC("관리코드", help="판매자상품코드 — 박스(관리코드)/PC낱개/소분(변환코드)/합포(-CB-)"),
        "상품명": TC("상품명", help="리스팅 상품명"),
        "규격": TC("규격", help="코드해석 규격(product_master / 소분)"),
        "코드유형": TC("코드유형", help="박스 · PC낱개 · 소분 · 합포 · 빈코드"),
        "N": NC("N", format="%.4g", help=n_help),
        "재고": NC("재고", format="localized",
                  help="product_master 재고. 박스=박스재고 · PC낱개=그 상품코드의 박스재고 · 소분=원코드 박스재고 · 합포=Σ구성코드 박스재고"),
        "매입가": NC("매입가", format="localized",
                   help="기준매입가 × N. 박스=박스매입단가 · PC낱개=낱개매입단가 · 소분=박스매입단가÷내품나누기 · 합포=Σ박스매입단가+700"),
        "판매가": NC("판매가", format="localized",
                   help="리스팅 판매가. 정산액엔 즉시할인·포인트 차감 후(net) 반영"),
        "배송비": NC("배송비", format="localized",
                   help=f"정산 반영 배송비 — {ship_src}. 정산액에 ×{settle} 가산"),
        "정산액": NC("정산액", format="localized",
                   help=f"= (판매가 − 즉시할인 − 포인트) × {rate_txt} + 배송비 × {settle}   (수수료 {comm_txt})"),
        "마진율": NC("마진율", format="percent",
                   help=f"= (정산액 − 매입가 − {ship:,}) ÷ 정산액   (실택배비 {ship:,}원)"),
        "기준마진율": NC("기준마진율", format="percent",
                     help=f"baseline_margin '{cfg['baseline_col']}' 컬럼의 확정마진율"),
        "탐지": NC("탐지(현-기준)", format="percent",
                  help="= 마진율 − 기준마진율. -1%p 미만이면 '마진미달'"),
        "권장가/제한": TC("권장가 / 제한",
                      help=(f"기준마진 달성 판매가(net 기준, 100원 올림): "
                            f"⌈((매입가+{ship:,})÷(1−기준마진율) − 배송비×{settle})÷{rate_txt}⌉. "
                            "제한상품은 제한 텍스트.")),
        "전월매출": NC("전월매출(이채널)", format="localized",
                    help=(f"{prev_ym or '최근 적재월'} 이 채널 판매금액(천년경영 정산). "
                          "박스/낱개/소분 통일(원박스 기준)·쿠팡=윙배송+로켓. "
                          "같은 상품의 박스·낱개 행은 같은 값 → 세로 합산 금지")),
        "전월매출(전체)": NC("전월매출(전체)", format="localized",
                        help=(f"{prev_ym or '최근 적재월'} 내 관리 8채널 합(나들·B2B 제외·통일 기준). "
                              "이 상품이 내 채널 전체로 얼마나 도는지")),
        "비고": TC("비고", help="미매칭·미등록 사유(정상 매칭이면 빈칸)"),
    }


channel = st.selectbox("채널", list(cmm.CHANNEL_CONFIG.keys()), index=0)
cfg = cmm.CHANNEL_CONFIG[channel]
key = cfg["key"]
_comm_txt = "상품별(다운로드)" if cfg.get("commission_source") == "download" else f"{cfg['commission']*100:.0f}%"
st.caption(
    f"수수료 {_comm_txt} · 배송비 정산 ×{cfg['ship_settle']} · "
    f"실택배비 {cfg['real_ship']:,}원 · 기준마진 '{cfg['baseline_col']}'"
    + (" · 마진제한 적용" if cfg.get("apply_floor") else "")
)

# ── 상품관리 갱신 ─────────────────────────────────────────────────────────────
committed = None
flash = None
with st.expander("📥 상품관리 갱신 (새 다운로드 업로드)"):
    multi = cfg.get("multi_file", False)   # ESM: 다운로드 500상품 한도 → 여러 배치 한번에
    up = st.file_uploader(
        f"{channel} 상품관리 다운로드 (.xlsx 전체 업로드)"
        + ("  — 여러 배치 파일을 한 번에 올리면 자동 병합" if multi else ""),
        type=["xlsx"], key=f"up_{key}", accept_multiple_files=multi)
    files = (up or []) if multi else ([up] if up is not None else [])
    if files:
        up_bytes = files[0].getvalue()      # raw 저장용(대표). multi(ESM)는 raw 미사용(모니터 전용)
        new_recs = []
        for f in files:                      # 여러 파일 파싱·이어붙이기(파일내 지마켓필터·dedup은 parse가 처리)
            new_recs += cmm.parse_download(f.getvalue(), cfg)
        dk = cfg.get("dedup_key")            # 다중파일 교차 중복제거(keep first)
        if dk:
            _seen, _uniq = set(), []
            for r in new_recs:
                k = r.get(dk)
                if k in _seen:
                    continue
                _seen.add(k); _uniq.append(r)
            new_recs = _uniq
        st.write(f"업로드 파싱: **{len(new_recs):,}건**"
                 + (f" ({len(files)}개 파일 병합)" if multi and len(files) > 1 else ""))
        if not _pat():
            st.error("저장용 PAT(st.secrets GITHUB_PAT)가 없어 커밋할 수 없습니다.")
        else:
            # 네이티브 raw 필수 채널(filter형=쿠팡): '신규만 추가'(append_rows_to_raw=openpyxl)는
            #   원본을 inlineStr로 변질시켜 업로더가 거부 → 비활성화. '전체 교체'(업로드 바이트 verbatim)만.
            native_raw = cfg.get("price_form", {}).get("mode") == "filter"
            b1, b2 = st.columns(2)
            if b1.button("전체 교체 저장", type="primary", use_container_width=True,
                         help="최신 전체 다운로드로 덮어쓰기 (신규+가격변동 반영)"
                              + ("" if multi else " + 원본양식 저장")):
                meta = _commit_listing(key, new_recs)
                if not multi:                # multi(ESM)=모니터전용 → raw 불요
                    _commit_raw(key, up_bytes)          # 원본 양식(전체 컬럼) 저장
                _load_listing.clear()
                committed = new_recs
                flash = f"전체 교체 완료 — {meta['rows']:,}건 ({meta['updated_at']})"
            _merge_help = ("이 채널은 네이티브 포맷 보존을 위해 '전체 교체'만 사용합니다 "
                           "(신규만 추가=openpyxl 저장이 원본을 inlineStr로 변질 → 업로드 거부)"
                           if native_raw else
                           "기존 유지 + 새 상품번호만 병합 (기존 상품의 가격변동은 미반영 → 갱신은 '전체 교체')")
            if b2.button("신규만 추가", use_container_width=True,
                         disabled=native_raw, help=_merge_help):
                cur, _ = _load_listing(key)
                merged, added = cmm.merge_listing(cur or [], new_recs)
                meta = _commit_listing(key, merged)
                if not multi:                # multi(ESM)=raw 미사용 → append 생략
                    added_pids = {r["상품번호"] for r in new_recs} - {r["상품번호"] for r in (cur or [])}
                    rcode, rawb = _gh_bytes(_raw_path(key))
                    newraw = (up_bytes if cfg.get("consolidate")          # 알리 다중시트: openpyxl raw 병합 부적합 → 최신 업로드 유지(raw 미사용)
                              else cmm.append_rows_to_raw(rawb, up_bytes, added_pids, cfg)
                              if (rcode == 200 and rawb) else up_bytes)
                    _commit_raw(key, newraw)
                _load_listing.clear()
                committed = merged
                flash = f"신규 {added:,}건 추가 — 총 {meta['rows']:,}건"
            if native_raw:
                st.caption("ℹ️ 이 채널은 '전체 교체'만 사용 — 원본 네이티브 포맷 보존(쿠팡 업로드 호환). "
                           "신규만 추가는 비활성화됨.")

if committed is not None:
    recs = committed
    meta = {"updated_at": datetime.now(_KST).isoformat(timespec="seconds"), "rows": len(committed)}
    st.success(flash)
    _accumulate_listing(key, committed)   # 이력 1d: 갱신 시점 가격 날짜본 적립(비차단)
else:
    recs, meta = _load_listing(key)

if not recs:
    st.info("저장된 상품관리가 없습니다. 위 '📥 상품관리 갱신'에서 다운로드를 올려 저장해 주세요.")
    st.stop()

st.caption(f"📦 저장된 상품관리 기준 · 최종 갱신 **{meta.get('updated_at', '?')}** · {len(recs):,}건")

# 구버전 listing 가드: 다운로드에서만 캡처되는 필드(price_form용 OFR/SKU, 상품별 수수료 등)가
# 저장본에 전부 비어 있으면(해당 필드 도입 전 스냅샷) 다운스트림이 빈/오류값 → 전체 교체 안내.
_pf = cfg.get("price_form")
_extra = cfg.get("extra_cols", {})
_stale = []
if _pf and _extra:
    _need = [k for k in _extra if k in set(_pf.get("source", {}).values())]
    _stale += [k for k in _need if all(not (r.get(k) or "") for r in recs)]
if cfg.get("commission_source") == "download":
    _cf = cfg["commission_field"]
    if all(not str(r.get(_cf) or "").strip() for r in recs):
        _stale.append(_cf + "(상품별 수수료)")
if _stale:
    st.warning(f"⚠️ 저장된 상품관리가 구버전이라 **{', '.join(_stale)}**(가) 비어 있습니다. "
               "위 '📥 상품관리 갱신 → **전체 교체**'를 1회 실행하면 채워집니다(신규만 추가로는 기존 행이 안 채워짐).")

_baseline_text = _load_baseline_text()
_baseline_override = cmm.parse_baseline_dict(_baseline_text) if _baseline_text else None
_floor_text = _load_floor_text()
_floor_override = cmm.parse_floor_dict(_floor_text) if _floor_text else None
rows, stats = cmm.compute_listing(recs, channel, str(_REF),
                                  baseline_override=_baseline_override, floor_override=_floor_override)

# ── KPI ──────────────────────────────────────────────────────────────────────
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("총 상품", f"{stats['총건수']:,}")
c2.metric("평균 마진율", f"{stats['평균마진율']*100:.2f}%" if stats["평균마진율"] is not None else "—")
c3.metric("마진 미달", f"{stats['마진미달']:,}", help="기준마진율보다 1%p 이상 낮음")
c4.metric("제한 상품", f"{stats['제한상품']:,}")
c5.metric("기준 미설정", f"{stats['미설정']:,}")
c6.metric("미매칭", f"{stats['미매칭']:,}")

# ── 필터 ──────────────────────────────────────────────────────────────────────
df = pd.DataFrame(rows)


def _rec_disp(r):
    if r["제한"]:
        return str(r["제한"])
    if r["매입가"] is None:
        return str(r["비고"]) or "미매칭"
    if r["기준마진율"] is None:
        return "기준 미설정"
    return f"{int(r['권장가']):,}" if pd.notna(r["권장가"]) else ""


df["권장가/제한"] = df.apply(_rec_disp, axis=1)

# ── 전월매출 (박스/낱개/소분 통일) — 적재 최신월·이 채널/전체 ──
_refs_canon = _load_refs()
_prev_ym, _sales_total, _sales_by_ch = _load_prev_sales()
_sanghos = [cmm._nfc(s) for s in _CH_TO_SANGHO.get(channel, [])]
_ch_sales: dict = {}
for _sg in _sanghos:
    for _cd, _v in _sales_by_ch.get(_sg, {}).items():
        _ch_sales[_cd] = _ch_sales.get(_cd, 0.0) + _v
_canon = df["관리코드"].map(lambda c: cmm.canonical_code(c, _refs_canon))
df["전월매출"] = _canon.map(lambda c: _ch_sales.get(c))
df["전월매출(전체)"] = _canon.map(lambda c: _sales_total.get(c))

search = st.text_input("🔍 검색", placeholder="상품번호 · 관리코드 · 상품명 (부분일치)",
                       label_visibility="collapsed")
q = cmm._nfc(search).lower() if search else ""

types = sorted(df["코드유형"].unique().tolist())
fc1, fc2 = st.columns([2, 3])
with fc1:
    pick = st.multiselect("코드유형", types, default=types)
with fc2:
    f1, f2, f3, f4 = st.columns(4)
    only_under = f1.checkbox("마진미달만", help="기준마진율보다 1%p 이상 낮은 상품")
    min_stock = f2.number_input("재고 N개 이상", min_value=0, value=0, step=1,
                                help="입력한 개수 이상의 재고만 표시 (0=전체)")
    only_floor = f3.checkbox("제한상품만")
    only_miss = f4.checkbox("미매칭만")

view = df[df["코드유형"].isin(pick)].copy()
if q:
    hay = (view["상품번호"].astype(str) + " ||| " + view["관리코드"].astype(str)
           + " ||| " + view["상품명"].astype(str)).str.lower()
    view = view[hay.str.contains(q, regex=False, na=False)]
if only_under:
    view = view[view["탐지"].notna() & (view["탐지"] < cmm.MARGIN_UNDER_THRESHOLD)]
if min_stock > 0:
    view = view[view["재고"].fillna(-1) >= min_stock]
if only_floor:
    view = view[view["제한"].astype(str) != ""]
if only_miss:
    view = view[view["매입가"].isna()]

DISPLAY = ["상품번호", "관리코드", "상품명", "규격", "코드유형", "N", "재고",
           "매입가", "판매가", "배송비", "정산액", "마진율", "기준마진율", "탐지", "권장가/제한",
           "전월매출", "전월매출(전체)", "비고"]

# ── 선택 — st.dataframe 다중행 선택(헤더 체크박스=전체선택 + 개별, 현재 필터/검색 기준) ──
st.session_state.setdefault("cmm_tblver", 0)  # 저장 시 +1 → 표 선택 강제 초기화(두뇌④ mo_tblver 패턴)
view_reset = view.reset_index(drop=True)
# 키에 행 수 포함 → 데이터/필터로 행 수가 바뀌면 위젯 리셋(어긋난 선택 복원 방지)
filter_sig = hash((tuple(sorted(pick)), only_under, min_stock, only_floor, only_miss, q, len(view_reset)))

event = st.dataframe(
    view_reset[DISPLAY],
    use_container_width=True,
    hide_index=True,
    on_select="rerun",
    selection_mode="multi-row",
    key=f"cmm_df_{key}_{filter_sig}_{st.session_state['cmm_tblver']}",
    column_config=_col_config(cfg, _prev_ym),
)
_raw_sel = event.selection.rows if event and getattr(event, "selection", None) else []
# 데이터 변동/rerun으로 범위를 벗어난 선택 인덱스 방어(IndexError: positional indexers out-of-bounds)
sel_rows = [i for i in _raw_sel if 0 <= i < len(view_reset)]
sel_pids = set(view_reset.iloc[sel_rows]["상품번호"].tolist()) if sel_rows else set()

st.caption(f"표시 {len(view):,} / 전체 {len(df):,}건 · ✅ 선택 **{len(sel_pids):,}건** "
           "(표 왼쪽 체크박스로 개별, 헤더 체크박스로 현재 화면 전체 선택)")

# ── 내보내기: CSV / 가격 일괄변경 양식 ───────────────────────────────────────
dc1, dc2 = st.columns(2)

csv_src = df[df["상품번호"].isin(sel_pids)] if sel_pids else view
csv_buf = StringIO()
csv_src[DISPLAY].to_csv(csv_buf, index=False)
dc1.download_button(
    f"📄 CSV 다운로드 ({'선택 '+str(len(sel_pids)) if sel_pids else '현재필터 '+str(len(view))}건)",
    data=csv_buf.getvalue().encode("utf-8-sig"),
    file_name=f"{channel}_마진모니터.csv",
    mime="text/csv",
    use_container_width=True,
)

if dc2.button(f"🛠️ 가격 일괄변경 양식 생성 (선택 {len(sel_pids)}건)",
              type="primary", use_container_width=True, disabled=(len(sel_pids) == 0)):
    pf = cfg.get("price_form")
    row_by = {r["상품번호"]: r for r in rows}
    rec_by = {r["상품번호"]: r for r in recs}
    if pf and pf.get("mode") == "append":
        # 식봄·캐시노트형: 채널 '일괄수정' 템플릿에 선택 행만 기입. 새 판매단가 = 권장가.
        items, prev, skipped = cmm.build_append_items(pf, rows, recs, sel_pids)
        if not items:
            st.session_state[f"form_{key}"] = {"error": "선택 상품 중 권장가 산출 가능 항목이 없습니다(미매칭/기준 미설정)."}
        else:
            out = cmm.build_price_form_append((_REF / pf["template"]).read_bytes(), items, pf)
            st.session_state[f"form_{key}"] = {
                "bytes": out, "kept": len(items), "skipped": skipped, "missing": [],
                "preview": prev, "append": True,
                "name": f"{channel}_가격변경_{datetime.now(_KST):%Y%m%d}.xlsx",
            }
    elif pf and pf.get("mode") == "filter":
        # 쿠팡형: 원본 다운로드의 '변경요청' 컬럼(P/Q)에 권장가·가짜정가 기입, 선택만 남김(R/S 미변경).
        rcode, raw = _gh_bytes(_raw_path(key))
        if rcode != 200 or not raw:
            st.session_state[f"form_{key}"] = {"error": "원본 양식(.xlsx)이 저장돼 있지 않습니다. '상품관리 갱신 → 전체 교체'를 1회 실행해 주세요."}
        else:
            out, prev, skipped, missing = cmm.build_filter_price_xlsx(raw, rows, sel_pids, cfg)
            if not prev:
                st.session_state[f"form_{key}"] = {"error": "선택 상품 중 권장가 산출 가능 항목이 없습니다(미매칭/기준 미설정)."}
            else:
                st.session_state[f"form_{key}"] = {
                    "bytes": out, "kept": len(prev), "skipped": skipped, "missing": missing,
                    "preview": prev, "append": False, "filter": True,
                    "name": f"{channel}_가격변경_{datetime.now(_KST):%Y%m%d}.xlsx",
                }
    else:
        new_prices, skipped = cmm.compute_new_prices(rows, recs, sel_pids)
        if not new_prices:
            st.session_state[f"form_{key}"] = {"error": "선택 상품 중 권장가 산출 가능 항목이 없습니다(미매칭/기준 미설정)."}
        else:
            rcode, raw = _gh_bytes(_raw_path(key))
            if rcode != 200 or not raw:
                st.session_state[f"form_{key}"] = {"error": "원본 양식(.xlsx)이 저장돼 있지 않습니다. '상품관리 갱신 → 전체 교체'를 1회 실행해 주세요."}
            else:
                out, kept, missing = cmm.build_bulk_price_xlsx(raw, new_prices, cfg)
                prev = []
                for pid, (np_, nd_) in new_prices.items():
                    rc, ro = rec_by[pid], row_by[pid]
                    cur_net = rc["판매가"] - rc["즉시할인"] - rc["포인트"]
                    new_net = np_ - nd_ - rc["포인트"]
                    prev.append({
                        "상품명": ro["상품명"], "현재판매가": int(rc["판매가"]), "현재할인": int(rc["즉시할인"]),
                        "새판매가": np_, "새할인": nd_, "권장가(net)": ro["권장가"],
                        "방향": "인상" if new_net > cur_net else ("인하" if new_net < cur_net else "유지"),
                    })
                st.session_state[f"form_{key}"] = {
                    "bytes": out, "kept": kept, "skipped": skipped, "missing": missing,
                    "preview": prev, "append": False,
                    "name": f"{channel}_가격일괄변경_{datetime.now(_KST):%Y%m%d}.xlsx",
                }

form = st.session_state.get(f"form_{key}")
if form:
    if form.get("error"):
        st.warning(form["error"])
    else:
        msg = f"양식 생성 완료 — **{form['kept']}건**"
        if form["skipped"]:
            msg += f" · 제외(권장가 없음) {len(form['skipped'])}건"
        if form["missing"]:
            msg += f" · 원본에 없어 누락 {len(form['missing'])}건(상품관리 전체 교체 필요)"
        st.success(msg)
        st.download_button(
            "⬇️ 가격 일괄변경 양식 다운로드 (.xlsx)",
            data=form["bytes"], file_name=form["name"], mime=_XLSX_MIME,
            use_container_width=True,
        )
        with st.expander(f"변경 미리보기 ({form['kept']}건)", expanded=True):
            st.dataframe(pd.DataFrame(form["preview"]), use_container_width=True, hide_index=True,
                         column_config={
                             "현재판매가": st.column_config.NumberColumn(format="localized"),
                             "현재할인": st.column_config.NumberColumn(format="localized"),
                             "새판매가": st.column_config.NumberColumn(format="localized"),
                             "새판매단가": st.column_config.NumberColumn(format="localized"),
                             "새할인": st.column_config.NumberColumn(format="localized"),
                             "정가": st.column_config.NumberColumn(format="localized"),
                             "권장가(net)": st.column_config.NumberColumn(format="localized"),
                         })
        if form.get("append"):
            st.caption(f"★ {channel} 일괄수정 양식입니다. 선택 상품만 기입 — 판매단가=기준마진 달성 권장가, "
                       "정가/할인전단가는 판매단가 이상으로 보존, 고정값(변경타입·진열·수량 등)은 양식 규칙대로 채웠습니다. "
                       f"{channel}에 그대로 업로드하세요.")
        elif form.get("filter"):
            st.caption(f"★ {channel} 가격 일괄변경 양식입니다. 선택 상품만 남기고 '변경 요청' 칸에 "
                       "판매가=기준마진 달성 권장가, 할인율기준가=무늬용 가짜정가(권장가+20~30%)를 기입했습니다. "
                       "판매상태·재고는 건드리지 않았으니 그대로 업로드하세요.")
        else:
            st.caption("★ 가격은 net(판매가−즉시할인−포인트) 기준으로 기준마진 달성가에 맞춥니다. "
                       "할인 우선: 인상 시 즉시할인을 먼저 줄이고 모자라면 판매가를 올립니다(인하는 반대).")


# ── 기준마진율 설정: 현재 마진율 → 기준마진율 (전 채널 공통 baseline_margin 편집) ──
# 선택 → 바로 아래 인라인 편집 → 저장 → 표 선택 초기화 + 즉시 반영. (두뇌④와 동일 패턴·session_state 중간단계 없음)
st.divider()
st.markdown("#### 🎯 기준마진율 설정 (현재 마진율 → 기준)")
_bcol = cfg["baseline_col"]
st.caption(f"위 표에서 상품을 선택하면 **현재 마진율**이 기본값으로 채워집니다. **새 기준(%)** 칸을 직접 고친 뒤 "
           f"저장하면 이 채널({channel}=`{_bcol}`) 컬럼만 갱신(다른 채널 보존)되고 표에 즉시 반영·선택은 풀립니다. "
           "미달 상품을 현재 마진으로 두면 더 이상 미달로 안 잡힙니다.")

if not sel_pids:
    st.info("⬆️ 위 표에서 상품을 선택하세요. (행 왼쪽 체크박스=개별 · 헤더 체크박스=현재 화면 전체)")
elif not _pat():
    st.warning("저장용 PAT(st.secrets GITHUB_PAT)가 없어 커밋할 수 없습니다.")
else:
    _prop, _conf = cmm.propose_baseline(rows, sel_pids, offset=0.0)
    if not _prop and not _conf:
        st.warning("선택 상품 중 마진율 산출 가능 항목이 없습니다(미매칭/정산불가).")
    else:
        _items = []   # 관리코드 1행: 기본값=현재(비충돌)/최저(충돌, 보수적) · 충돌은 후보 표시
        for _code, _v in _prop.items():
            _items.append({"관리코드": _code, "현재마진": [_v], "기본값": _v, "충돌": False})
        for _code, _cands in _conf.items():
            _vals = sorted({c["값"] for c in _cands})
            _items.append({"관리코드": _code, "현재마진": _vals, "기본값": min(_vals), "충돌": True})
        _items.sort(key=lambda x: x["관리코드"])
        _curbase = cmm.parse_baseline_dict(_load_baseline_text())
        if any(it["충돌"] for it in _items):
            st.warning("⚠️ '현재 마진율'이 여러 개(`/`)인 관리코드는 같은 코드에 마진이 다른 상품이 섞인 경우입니다. "
                       "**새 기준(%)** 칸에 원하는 값을 직접 정해 주세요(기본은 가장 낮은 값).")
        _disp = []
        for it in _items:
            _code = it["관리코드"]
            _old = (_curbase.get(_code, {}) or {}).get(_bcol, "") or ""
            _disp.append({
                "관리코드": _code,
                f"기존 {_bcol}": (f"{float(_old)*100:.1f}%" if _old else "—"),
                "현재 마진율": " / ".join(f"{v*100:.1f}%" for v in it["현재마진"]) + (" ⚠️" if it["충돌"] else ""),
                "새 기준(%)": round(it["기본값"] * 100, 1),
            })
        _edited = st.data_editor(
            pd.DataFrame(_disp),
            column_config={
                "관리코드": st.column_config.TextColumn(disabled=True),
                f"기존 {_bcol}": st.column_config.TextColumn(disabled=True),
                "현재 마진율": st.column_config.TextColumn("현재 마진율(기본값 출처)", disabled=True),
                "새 기준(%)": st.column_config.NumberColumn(
                    "새 기준(%) ✏️", min_value=0.0, max_value=100.0, step=0.1, format="%.1f",
                    help="직접 수정 가능. 0.1%p 단위. 비우면 그 행은 저장 안 함."),
            },
            hide_index=True, use_container_width=True,
            key=f"bl_editor_{key}_{st.session_state['cmm_tblver']}",
        )
        st.caption(f"적용 대상 **{len(_disp)} 관리코드** · 이 채널({_bcol}) 컬럼만 수정 · 새 기준 = 입력값(0.1%p 반올림 저장)")
        if st.button("💾 저장 (커밋)", type="primary", key=f"bl_save_{key}"):
            _updates = {}
            for _, _r in _edited.iterrows():
                _val = _r["새 기준(%)"]
                if _val is None or (isinstance(_val, float) and pd.isna(_val)):
                    continue                      # 빈 칸 = 그 행 제외
                _updates[str(_r["관리코드"])] = round(float(_val) / 100, 3)
            if not _updates:
                st.warning("저장할 행이 없습니다(새 기준이 모두 비어 있음).")
            else:
                _newtext, _upd, _added = cmm.update_baseline_csv(_load_baseline_text(), _bcol, _updates)
                _commit_baseline(_newtext)
                _load_baseline_text.clear()
                st.session_state["cmm_tblver"] += 1       # 표 선택 초기화(체크박스 풀림)
                st.session_state.pop(f"bl_{key}", None)    # 구버전 잔여 단계 정리
                st.success(f"기준마진율 저장 완료 — 수정 {_upd} · 신규 {_added} 관리코드 ({_bcol}). 표에 즉시 반영됩니다.")
                st.rerun()


# ── 제한 상품 등록 / 해제 (margin_floor — 권장가 ↔ 제한) ──
st.divider()
st.markdown("#### 🔒 제한 상품 등록 / 해제")
st.caption("선택한 상품을 **제한**으로 등록하면 권장가 대신 제한 문구가 표시되고, 기준마진율 최적화(두뇌④)에서도 제외됩니다. "
           "제한내용을 입력하면 등록, **비우고 저장하면 해제**됩니다. margin_floor.csv(전 채널 공통)에 반영되어 표에 즉시 적용됩니다.")
if not sel_pids:
    st.info("⬆️ 위 표에서 상품을 선택하세요.")
elif not _pat():
    st.warning("저장용 PAT(st.secrets GITHUB_PAT)가 없어 커밋할 수 없습니다.")
else:
    _floormap = cmm.parse_floor_dict(_load_floor_text())
    _by_code = {}
    for _r in (r for r in rows if r["상품번호"] in sel_pids):
        _c = str(_r["관리코드"])
        if _c and _c not in _by_code:
            _by_code[_c] = str(_r["상품명"] or "")
    _fdisp = []
    for _c, _nm in sorted(_by_code.items()):
        _cur = _floormap.get(_c, {})
        _fdisp.append({
            "관리코드": _c, "상품명": _nm,
            "현재": "🔒 제한" if _cur else "—",
            "제한내용": ((_cur.get("제한내용") or _cur.get("비고") or "") if _cur else ""),
        })
    _fed = st.data_editor(
        pd.DataFrame(_fdisp),
        column_config={
            "관리코드": st.column_config.TextColumn(disabled=True),
            "상품명": st.column_config.TextColumn(disabled=True),
            "현재": st.column_config.TextColumn(disabled=True),
            "제한내용": st.column_config.TextColumn(
                "제한내용 ✏️ (입력=등록 · 비움=해제)",
                help="권장가 칸에 표시될 문구. 예: 마진율 민감 상품 / 배송비 미포함 17000원. "
                     "비우고 저장하면 제한 해제(빈 채로 두면 등록 안 함)."),
        },
        hide_index=True, use_container_width=True,
        key=f"fl_editor_{key}_{st.session_state['cmm_tblver']}",
    )
    st.caption(f"선택 **{len(_fdisp)} 관리코드** · 제한은 전 채널 공통(margin_floor) · 입력한 행만 등록")
    if st.button("🔒 제한 저장 (등록 / 수정 / 해제)", type="primary", key=f"fl_save_{key}"):
        _ups, _rms = {}, set()
        for _, _r in _fed.iterrows():
            _c = str(_r["관리코드"])
            _note = str(_r["제한내용"] or "").strip()
            if _note:
                _ups[_c] = {"상품명": str(_r["상품명"] or ""), "비고": "마진율 민감 상품", "제한내용": _note}
            elif _c in _floormap:
                _rms.add(_c)
        if not _ups and not _rms:
            st.warning("변경할 내용이 없습니다(제한내용을 입력하거나, 해제하려면 기존 제한을 비우세요).")
        else:
            _newtext, _na, _nr = cmm.update_floor_csv(_load_floor_text(), _ups, _rms)
            _commit_floor(_newtext)
            _load_floor_text.clear()
            st.session_state["cmm_tblver"] += 1
            st.success(f"제한 저장 완료 — 등록/수정 {_na} · 해제 {_nr} 관리코드. 표에 즉시 반영됩니다.")
            st.rerun()
