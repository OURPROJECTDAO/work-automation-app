"""분류 기준 데이터 — 도서산간·필터링·미배송 키워드 관리."""
import sys, base64, json, urllib.request, urllib.parse
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))  # repo root

import streamlit as st
import pandas as pd

REPO_ROOT = Path(__file__).parent.parent.parent.parent
REF_DIR   = REPO_ROOT / "reference"
REPO      = "OURPROJECTDAO/work-automation-app"

def _get_pat() -> str:
    return st.secrets.get("GITHUB_PAT", "")

def _gh_get_sha(path: str, pat: str) -> str | None:
    url = f"https://api.github.com/repos/{REPO}/contents/{urllib.parse.quote(path)}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {pat}"})
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read())["sha"]
    except Exception:
        return None

def _gh_put_csv(path: str, df: pd.DataFrame, msg: str, pat: str) -> bool:
    url = f"https://api.github.com/repos/{REPO}/contents/{urllib.parse.quote(path)}"
    sha = _gh_get_sha(path, pat)
    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    payload: dict = {"message": msg, "content": base64.b64encode(csv_bytes).decode()}
    if sha: payload["sha"] = sha
    req = urllib.request.Request(url, data=json.dumps(payload).encode(),
        headers={"Authorization": f"Bearer {pat}", "Content-Type": "application/json"},
        method="PUT")
    try:
        with urllib.request.urlopen(req): return True
    except urllib.error.HTTPError as e:
        st.error(f"GitHub 저장 실패: {e.read().decode()[:200]}")
        return False

@st.cache_data(ttl=60)
def load_ref(filename: str) -> pd.DataFrame:
    path = REF_DIR / filename
    if path.exists():
        return pd.read_csv(path, encoding="utf-8-sig", dtype=str).fillna("")
    return pd.DataFrame()

def render_tabs(config: dict, readonly: bool, pat: str) -> None:
    tabs = st.tabs(list(config.keys()))
    for tab, (name, cfg) in zip(tabs, config.items()):
        with tab:
            df = load_ref(cfg["file"])
            st.caption(cfg["desc"])
            st.write(f"현재 **{len(df)}건**")
            if cfg["large"]:
                st.dataframe(df.head(20), use_container_width=True, height=300)
                st.caption("※ 처음 20행만 미리보기. 전체 교체는 아래에서 파일 업로드.")
                if not readonly:
                    new_file = st.file_uploader("새 파일로 교체 (Excel .xlsx 또는 CSV .csv)",
                        type=["xlsx","csv"], key=f"upload_{name}")
                    if new_file:
                        new_df = (pd.read_csv(new_file, dtype=str) if new_file.name.endswith(".csv")
                                  else pd.read_excel(new_file, dtype=str)).fillna("")
                        st.write(f"업로드된 데이터: **{len(new_df)}건**")
                        st.dataframe(new_df.head(10), use_container_width=True)
                        if st.button(f"💾 {name} 저장", key=f"save_{name}"):
                            with st.spinner("GitHub에 저장 중..."):
                                ok = _gh_put_csv(f"reference/{cfg['file']}", new_df,
                                    f"data: {name} 업데이트 ({len(new_df)}건)", pat)
                            if ok:
                                st.success("✅ 저장 완료! 앱 재배포 후 반영됩니다.")
                                st.cache_data.clear()
            else:
                if readonly:
                    st.dataframe(df, use_container_width=True)
                else:
                    edited = st.data_editor(df, use_container_width=True,
                        num_rows="dynamic", key=f"editor_{name}")
                    if st.button(f"💾 {name} 저장", key=f"save_{name}"):
                        key_col = cfg.get("key_col")
                        save_df = (edited[edited[key_col].str.strip() != ""].reset_index(drop=True)
                                   if key_col and key_col in edited.columns
                                   else edited.dropna(how="all").reset_index(drop=True))
                        with st.spinner("GitHub에 저장 중..."):
                            ok = _gh_put_csv(f"reference/{cfg['file']}", save_df,
                                f"data: {name} 업데이트 ({len(save_df)}건)", pat)
                        if ok:
                            st.success(f"✅ 저장 완료! ({len(save_df)}건)")
                            st.cache_data.clear()
            csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(label=f"⬇️ {name} CSV 다운로드", data=csv_bytes,
                file_name=cfg["file"], mime="text/csv", key=f"dl_{name}")


CLASSIFY_CONFIG = {
    "도서산간리스트": {
        "file": "dosan_list.csv",
        "desc": "도서산간 지역 키워드 (10,000건+). 주소에 포함되면 도서산간으로 분류.",
        "large": True,
    },
    "도서산간아님": {
        "file": "dosan_except_list.csv",
        "desc": "도서산간에서 제외할 예외 키워드 (소수). 주소에 포함되면 도서산간 제외.",
        "large": False,
    },
    "필터링리스트": {
        "file": "filter_list.csv",
        "desc": "필터링 대상 상품 코드 목록. 상품명에 포함되면 필터링확인 시트로 분류.",
        "large": False,
    },
    "미배송지리스트": {
        "file": "undelivered_list.csv",
        "desc": "미배송 지역 주소 키워드. 주소에 포함되면 미배송지역확인 시트로 분류.",
        "large": False,
    },
}

st.title("📋 분류 기준 데이터")
st.caption("오픈마켓합포도서산간확인 워크플로우에서 사용합니다.")

pat = _get_pat()
if not pat:
    st.warning("⚠️ GITHUB_PAT secret이 설정되지 않아 읽기 전용 모드입니다.")
    readonly = True
else:
    readonly = False

render_tabs(CLASSIFY_CONFIG, readonly, pat)
