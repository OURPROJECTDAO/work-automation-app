"""기준데이터관리 — 발주서 출력 업무 참조 데이터."""
import base64
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))

import pandas as pd
import streamlit as st

_GITHUB_API = "https://api.github.com/repos/OURPROJECTDAO/work-automation-app/contents"
_REF_PATH   = "reference"

_REF_CONFIG = {
    "멸치쇼핑 분류표": {
        "filename": "logistics_classification.csv",
        "key_col":  "관리코드",
        "dtypes":   {"관리코드": str, "구분": str},
        "help":     "관리코드 → 음료 / 식품 / 선물세트 분류",
        "large":    False,
    },
    "낱개처리목록": {
        "filename": "unit_list.csv",
        "key_col":  "관리코드",
        "dtypes":   {"관리코드": str, "원코드": str},
        "help":     "낱개 판매 코드 → 원코드 + 박스내품 배수",
        "large":    False,
    },
    "규격파일": {
        "filename": "spec_master.csv",
        "key_col":  "관리코드",
        "dtypes":   {"관리코드": str},
        "help":     "관리코드 → 규격 (4000+ 행, 읽기 전용 검색)",
        "large":    True,
    },
}


def _get_pat() -> str:
    return st.secrets.get("GITHUB_PAT", "")


def _github_get(token: str, path: str) -> bytes:
    import urllib.request
    req = urllib.request.Request(
        f"{_GITHUB_API}/{path}",
        headers={"Authorization": f"Bearer {token}",
                 "Accept": "application/vnd.github.raw"},
    )
    with urllib.request.urlopen(req) as r:
        return r.read()


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
        f"{_GITHUB_API}/{path}", data=data, method="PUT",
        headers={"Authorization": f"Bearer {token}",
                 "Accept": "application/vnd.github+json",
                 "Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def _load_csv(token: str, filename: str, dtypes: dict) -> pd.DataFrame:
    import io
    raw = _github_get(token, f"{_REF_PATH}/{filename}")
    return pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig", dtype=dtypes)


def _save_csv(token: str, filename: str, df: pd.DataFrame, commit_msg: str):
    sha = _github_get_sha(token, f"{_REF_PATH}/{filename}")
    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    b64 = base64.b64encode(csv_bytes).decode()
    _github_put(token, f"{_REF_PATH}/{filename}", b64, sha, commit_msg)


# ─────────────────────────────────────────────────────────
st.title("📋 기준데이터관리 — 발주서출력업무")

token = _get_pat()
if not token:
    st.error("GITHUB_PAT 시크릿이 없습니다.")
    st.stop()

tabs = st.tabs(list(_REF_CONFIG.keys()))

for tab, (tab_name, cfg) in zip(tabs, _REF_CONFIG.items()):
    with tab:
        st.caption(cfg["help"])
        try:
            df = _load_csv(token, cfg["filename"], cfg["dtypes"])
        except Exception as e:
            st.error(f"로딩 실패: {e}")
            continue

        if cfg["large"]:
            # 규격파일 : 검색 전용
            query = st.text_input("🔍 관리코드·상품명 검색", key=f"search_{tab_name}")
            if query:
                mask = df.astype(str).apply(
                    lambda col: col.str.contains(query, na=False)
                ).any(axis=1)
                st.dataframe(df[mask], use_container_width=True, height=350)
                st.caption(f"{mask.sum()}건 검색됨 / 전체 {len(df)}건")
            else:
                st.dataframe(df.head(30), use_container_width=True, height=350)
                st.caption(f"전체 {len(df)}건 (상위 30건 미리보기)")

            up = st.file_uploader("전체 파일 교체 (CSV)", type=["csv"],
                                   key=f"upload_{tab_name}")
            if up:
                new_df = pd.read_csv(up, encoding="utf-8-sig",
                                     dtype=cfg["dtypes"])
                st.info(f"{len(new_df)}건 인식. 저장하면 전체 교체됩니다.")
                if st.button("💾 저장", key=f"save_upload_{tab_name}",
                             type="primary"):
                    _save_csv(token, cfg["filename"], new_df,
                              f"ref: {tab_name} 전체 교체")
                    st.success("✅ 저장 완료")
                    st.rerun()

        else:
            # 일반 : 인라인 편집
            edited = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic",
                key=f"editor_{tab_name}",
                height=400,
            )
            if st.button(f"💾 {tab_name} 저장", key=f"save_{tab_name}",
                         type="primary", use_container_width=True):
                clean = edited.dropna(subset=[cfg["key_col"]])
                _save_csv(token, cfg["filename"], clean,
                          f"ref: {tab_name} 갱신")
                st.success(f"✅ {len(clean)}건 저장 완료")
                st.rerun()
