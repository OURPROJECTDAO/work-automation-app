"""{doc}"""
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

def _edit_with_search(df, name, key_col=None, height_cap=560):
    """검색 필터 + 인라인 편집 + 키 무관 merge-back.
    필터가 걸려 있으면 필터된 행만 편집 대상이 되고, 필터 밖 행은 그대로 보존된다.
    (부분집합으로 전체 CSV를 덮어써 나머지가 날아가는 사고 방지.)"""
    import pandas as _pd
    q = st.text_input("🔍 검색 (모든 열 대상)", key=f"search_{name}",
                      placeholder="코드·이름 일부 입력 · 비우면 전체 편집")
    sdf = df.fillna("").astype(str)
    if q:
        mask = sdf.apply(lambda c: c.str.contains(q, case=False, na=False, regex=False)).any(axis=1)
        view = df[mask]
        st.caption(f"🔎 {int(mask.sum())}건 일치 / 전체 {len(df)}건 · "
                   f"**필터된 행만 편집**되고 나머지는 보존됩니다.")
    else:
        mask = _pd.Series(True, index=df.index)
        view = df
    h = min(38 * (len(view) + 1) + 3, height_cap)
    edited = st.data_editor(view, width="stretch", num_rows="dynamic",
                            key=f"editor_{name}", height=h)
    result = _pd.concat([df[~mask], edited], ignore_index=True)
    if key_col and key_col in result.columns:
        result = result[result[key_col].astype(str).str.strip() != ""].reset_index(drop=True)
    else:
        result = result.dropna(how="all").reset_index(drop=True)
    return result

def render_tabs(config: dict, readonly: bool, pat: str) -> None:
    tabs = st.tabs(list(config.keys()))
    for tab, (name, cfg) in zip(tabs, config.items()):
        with tab:
            df = load_ref(cfg["file"])
            st.caption(cfg["desc"])
            st.write(f"현재 **{len(df)}건**")
            if cfg["large"]:
                ql = st.text_input("🔍 검색 (모든 열 대상)", key=f"searchL_{name}",
                                   placeholder="비우면 처음 20행 미리보기")
                if ql:
                    sdf = df.fillna("").astype(str)
                    m = sdf.apply(lambda c: c.str.contains(ql, case=False, na=False,
                                                           regex=False)).any(axis=1)
                    st.dataframe(df[m], width="stretch", height=300)
                    st.caption(f"🔎 {int(m.sum())}건 / 전체 {len(df)}건 (대용량 — 읽기전용 검색)")
                else:
                    st.dataframe(df.head(20), width="stretch", height=300)
                    st.caption("※ 처음 20행만 미리보기. 대용량이라 인라인 편집 대신 아래에서 파일 교체.")
                if not readonly:
                    new_file = st.file_uploader("새 파일로 교체 (Excel .xlsx 또는 CSV .csv)",
                        type=["xlsx","csv"], key=f"upload_{name}")
                    if new_file:
                        new_df = (pd.read_csv(new_file, dtype=str) if new_file.name.endswith(".csv")
                                  else pd.read_excel(new_file, dtype=str)).fillna("")
                        st.write(f"업로드된 데이터: **{len(new_df)}건**")
                        st.dataframe(new_df.head(10), width="stretch")
                        if st.button(f"💾 {name} 저장", key=f"save_{name}"):
                            with st.spinner("GitHub에 저장 중..."):
                                ok = _gh_put_csv(f"reference/{cfg['file']}", new_df,
                                    f"data: {name} 업데이트 ({len(new_df)}건)", pat)
                            if ok:
                                st.success("✅ 저장 완료! 앱 재배포 후 반영됩니다.")
                                st.cache_data.clear()
            else:
                if readonly:
                    st.dataframe(df, width="stretch")
                else:
                    save_df = _edit_with_search(df, name, cfg.get("key_col"))
                    if st.button(f"💾 {name} 저장", key=f"save_{name}"):
                        with st.spinner("GitHub에 저장 중..."):
                            ok = _gh_put_csv(f"reference/{cfg['file']}", save_df,
                                f"data: {name} 업데이트 ({len(save_df)}건)", pat)
                        if ok:
                            st.success(f"✅ 저장 완료! ({len(save_df)}건)")
                            st.cache_data.clear()
            csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(label=f"⬇️ {name} CSV 다운로드", data=csv_bytes,
                file_name=cfg["file"], mime="text/csv", key=f"dl_{name}")


PRICE_CONFIG = {
    "SKU단가표": {
        "file": "sku_list.csv",
        "desc": "온누리양식 발주서 합계 계산 기준 단가표. **공급가(VAT 포함)** 와 **배송비** 가 합계:판매가 산출에 직접 사용됩니다.",
        "large": False,
        "key_col": "관리코드",
    },
}

st.title("💰 단가 기준 데이터")
st.caption("온누리양식_발주서 워크플로우에서 사용합니다.")

pat = _get_pat()
if not pat:
    st.warning("⚠️ GITHUB_PAT secret이 설정되지 않아 읽기 전용 모드입니다.")
    readonly = True
else:
    readonly = False

render_tabs(PRICE_CONFIG, readonly, pat)
