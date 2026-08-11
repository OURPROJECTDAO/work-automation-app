"""기준데이터관리 — 리테일앤인사이트(토마토).

관리 대상 (reference/, 고정):
  - retail_barcode_map.csv   바코드 ↔ 관리코드 변환표 (바코드, 관리코드, N, 상품명)

★ 왜 필요한가: 리테일앤인사이트 상품 다운로드는 '상품코드' 자리에 **바코드(EAN-13)만** 오고
  우리 관리코드가 없다. 채널마진모니터가 이 표로 관리코드를 해소해 매입가·재고·기준마진율을 붙인다.
  **조회 시점 변환**이라 여기서 저장하면 상품관리 재다운로드 없이 즉시 반영된다.

컬럼:
  - 바코드   : 리테일 다운로드의 '상품코드'(키). 합성코드(monsterset1 등)도 그대로 넣으면 됨.
  - 관리코드 : 우리 ERP 관리코드(박스). PC낱개·소분 코드도 4-tier로 해석된다.
  - N        : 리테일 등재 1건 = 우리 박스 몇 개인가. **비우면** hapo_multiplier(바코드) → 없으면 1.
               리테일만 묶음 단위가 다른 경우(예: 355ml 24개×2박스)에 여기서 채널 전용으로 덮어쓴다.
  - 상품명   : 사람이 알아보기 위한 메모(계산에 미사용).

편집: 검색 후 인라인 편집(필터된 행만 편집·나머지 보존) + 전체 CSV 교체(expander).
"""
import base64
import json
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))

import pandas as pd
import streamlit as st

from core.base import sanitize_ref_df

_REPO = "OURPROJECTDAO/work-automation-app"
_REF = Path(__file__).resolve().parent.parent.parent.parent / "reference"
_FILE = "retail_barcode_map.csv"

st.title("🧾 기준데이터관리 — 리테일앤인사이트")
st.caption("리테일앤인사이트(토마토) 상품 다운로드는 **바코드만** 옵니다. 여기서 바코드↔관리코드를 "
           "이어주면 채널마진모니터가 매입가·재고·기준마진율을 붙여 마진을 계산합니다. "
           "저장 즉시 반영됩니다(상품관리 재다운로드 불요).")


def _commit_csv(path: str, df: pd.DataFrame, message: str):
    token = st.secrets.get("GITHUB_PAT", "")
    if not token:
        st.error("GITHUB_PAT 시크릿이 설정되지 않았습니다.")
        return False
    df = sanitize_ref_df(df)   # 붙여넣기 개행·유령 빈행 제거 (2026-08-04)
    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    b64 = base64.b64encode(csv_bytes).decode()
    base = f"https://api.github.com/repos/{_REPO}/contents/{path}"
    hdr = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
    sha = None
    try:
        with urllib.request.urlopen(urllib.request.Request(base, headers=hdr)) as r:
            sha = json.loads(r.read())["sha"]
    except Exception:
        pass
    body = {"message": message, "content": b64}
    if sha:
        body["sha"] = sha
    req = urllib.request.Request(base, data=json.dumps(body).encode(), method="PUT",
                                 headers={**hdr, "Content-Type": "application/json"})
    urllib.request.urlopen(req)
    return True


def _edit_with_search(df, name, key_col=None, height_cap=620, save_label="💾 저장(커밋)"):
    """검색 필터 + 인라인 편집 + 키 무관 merge-back (기준데이터관리 공통 패턴).

    필터가 걸려 있으면 필터된 행만 편집 대상이고 나머지는 보존된다.
    data_editor와 저장 버튼은 st.form으로 묶는다(폼 밖 버튼이면 마지막 셀 입력이
    커밋되기 전에 클릭이 처리되어 '한 번에 저장 안 됨' 경합 — pitfalls).
    """
    q = st.text_input("🔍 검색 (모든 열 대상)", key=f"search_{name}",
                      placeholder="바코드·관리코드·상품명 일부 · 비우면 전체 편집")
    sdf = df.fillna("").astype(str)
    if q:
        mask = sdf.apply(lambda c: c.str.contains(q, case=False, na=False, regex=False)).any(axis=1)
        view = df[mask]
        st.caption(f"🔎 {int(mask.sum())}건 일치 / 전체 {len(df)}건 · "
                   f"**필터된 행만 편집**되고 나머지는 보존됩니다.")
    else:
        mask = pd.Series(True, index=df.index)
        view = df
    h = min(38 * (len(view) + 1) + 3, height_cap)
    with st.form(key=f"form_{name}", clear_on_submit=False):
        edited = st.data_editor(view, width="stretch", num_rows="dynamic",
                                key=f"editor_{name}", height=h)
        submitted = st.form_submit_button(save_label, type="primary")
    result = pd.concat([df[~mask], edited], ignore_index=True)
    if key_col and key_col in result.columns:
        result = result[result[key_col].astype(str).str.strip() != ""].reset_index(drop=True)
    else:
        result = result.dropna(how="all").reset_index(drop=True)
    return result, submitted


local = _REF / _FILE
if not local.exists():
    st.warning(f"{_FILE} 없음")
    st.stop()

df = pd.read_csv(local, encoding="utf-8-sig", dtype=str).fillna("")
for c in ["바코드", "관리코드", "N", "상품명"]:
    if c not in df.columns:
        df[c] = ""
df = df[["바코드", "관리코드", "N", "상품명"]]

c1, c2, c3 = st.columns(3)
c1.metric("등록 매핑", f"{len(df):,}")
c2.metric("N 지정", f"{(df['N'].astype(str).str.strip() != '').sum():,}",
          help="비어 있으면 hapo_multiplier(바코드) → 없으면 1")
c3.metric("관리코드 빈칸", f"{(df['관리코드'].astype(str).str.strip() == '').sum():,}")

st.info("**N** = 리테일 등재 1건이 우리 박스 몇 개인가. 예) 리테일 '코카콜라 355ml 업소용'은 "
        "1건이 2박스(38,600원 ≈ 박스매입 15,931×2) → N=2. 비워두면 hapo_multiplier(바코드) 값을 쓰고, "
        "그것도 없으면 1로 계산합니다. **N이 틀리면 마진율이 통째로 틀립니다.**", icon="ℹ️")

save_df, submitted = _edit_with_search(df, _FILE, "바코드",
                                       save_label="💾 바코드 변환표 저장(커밋)")
if submitted:
    if _commit_csv(f"reference/{_FILE}", save_df, "ref: 리테일앤인사이트 바코드 변환표 갱신"):
        st.success(f"✅ {len(save_df)}건 저장 — 채널마진모니터에 즉시 반영됩니다"
                   "(캐시 때문에 안 보이면 채널 페이지에서 새로고침).")

with st.expander("📁 전체 CSV 파일로 교체 (선택)"):
    st.caption("리테일 쪽에서 받은 변환표 전체를 갈아끼울 때만 사용하세요. 기존 N 지정도 함께 사라집니다.")
    up = st.file_uploader(f"새 {_FILE} 업로드 (.csv, 전체 교체)", type=["csv"], key=f"up_{_FILE}")
    if up:
        try:
            new = pd.read_csv(up, encoding="utf-8-sig", dtype=str).fillna("")
        except UnicodeDecodeError:
            up.seek(0)
            new = pd.read_csv(up, dtype=str).fillna("")
        miss = [c for c in ["바코드", "관리코드"] if c not in new.columns]
        if miss:
            st.error(f"필수 컬럼 누락: {miss} · 현재 컬럼: {list(new.columns)}")
        else:
            st.success(f"미리보기 {len(new)}행")
            st.dataframe(new.head(20), width="stretch")
            if st.button("💾 전체 교체 커밋", key="saveup_retail"):
                if _commit_csv(f"reference/{_FILE}", new, "ref: 리테일앤인사이트 바코드 변환표 전체 교체"):
                    st.success("✅ 전체 교체 완료.")
