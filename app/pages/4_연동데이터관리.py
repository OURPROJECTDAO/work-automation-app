"""연동데이터관리 — 매일 갱신이 필요한 데이터(상품관리 시트) 업로드·관리."""
import base64
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import pandas as pd
import streamlit as st

from core.workflows.logistics_order import get_product_master_updated

_GITHUB_API = "https://api.github.com/repos/OURPROJECTDAO/work-automation-app/contents"
_REF_PATH   = "reference"


def _get_pat() -> str:
    return st.secrets.get("GITHUB_PAT", "")


def _github_get_sha(token: str, path: str) -> str | None:
    import urllib.request, json
    req = urllib.request.Request(
        f"{_GITHUB_API}/{path}",
        headers={"Authorization": f"Bearer {token}",
                 "Accept": "application/vnd.github+json"},
    )
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read())["sha"]
    except Exception:
        return None


def _github_put(token: str, path: str, content_b64: str, sha: str | None, msg: str):
    import urllib.request, json
    body = {"message": msg, "content": content_b64}
    if sha:
        body["sha"] = sha
    data = json.dumps(body).encode()
    req = urllib.request.Request(
        f"{_GITHUB_API}/{path}",
        data=data,
        method="PUT",
        headers={"Authorization": f"Bearer {token}",
                 "Accept": "application/vnd.github+json",
                 "Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def _save_to_github(token: str, filename: str, content_bytes: bytes, commit_msg: str):
    sha = _github_get_sha(token, f"{_REF_PATH}/{filename}")
    b64 = base64.b64encode(content_bytes).decode()
    _github_put(token, f"{_REF_PATH}/{filename}", b64, sha, commit_msg)


# ─────────────────────────────────────────────────────────
st.title("🔗 연동데이터관리")
st.caption("매일 갱신이 필요한 데이터를 여기서 업로드합니다.")

st.subheader("상품관리 시트")

updated = get_product_master_updated()
if updated:
    st.info(f"📅 마지막 업데이트: **{updated}**")
else:
    st.warning("⚠️ 상품관리 데이터가 아직 없습니다. 파일을 업로드해주세요.")

uploaded = st.file_uploader(
    "상품관리 .xlsx 업로드",
    type=["xlsx"],
    help="Exp______상품관리_.xlsx — 오늘의 재고 현황이 담긴 파일",
)

if uploaded:
    try:
        df = pd.read_excel(uploaded, header=0, dtype=str)
        st.success(f"✅ {len(df)}개 상품 인식됨")

        col_preview, col_info = st.columns([3, 1])
        with col_preview:
            st.dataframe(df.head(8), use_container_width=True, height=200)
        with col_info:
            if df.shape[1] >= 15:
                st.metric("관리코드 컬럼", df.columns[4])
                st.metric("박스재고 컬럼", df.columns[14])
            st.metric("전체 상품 수", len(df))

        if st.button("📤 업로드 & 저장", type="primary", use_container_width=True):
            token = _get_pat()
            if not token:
                st.error("GITHUB_PAT 시크릿이 설정되지 않았습니다.")
            else:
                with st.spinner("GitHub에 저장 중..."):
                    # CSV 변환 후 저장
                    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                    _save_to_github(token, "product_master.csv", csv_bytes,
                                    "data: 상품관리 갱신")

                    # 타임스탬프 저장
                    now_str = datetime.now().strftime("%Y년 %m월 %d일 %H시 %M분")
                    ts_bytes = now_str.encode("utf-8")
                    _save_to_github(token, "product_master_updated.txt", ts_bytes,
                                    "data: 상품관리 업데이트 시각 기록")

                st.success(f"✅ 저장 완료! ({now_str})")
                st.rerun()
    except Exception as e:
        st.error(f"파일 읽기 오류: {e}")
        st.exception(e)
