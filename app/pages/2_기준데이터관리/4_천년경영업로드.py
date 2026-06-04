"""기준데이터관리 — 천년경영업로드.

관리 대상 (reference/, 고정):
  - bm_commission.csv  배민상회 수수료율 (관리코드, 옵션명, 수수료율)
  - sub_list.csv       소분목록 (관리코드, 낱개개수, 원코드, 구분)
멸치쇼핑 분류표(구분 판정)는 **발주서출력업무**와 공유 → 그 페이지에서 관리.
"""
import base64
import json
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))

import pandas as pd
import streamlit as st

_REPO = "OURPROJECTDAO/work-automation-app"
_REF = Path(__file__).resolve().parent.parent.parent.parent / "reference"

st.title("🏪 기준데이터관리 — 천년경영업로드")
st.caption("배민상회 수수료율·소분목록을 관리합니다. "
           "구분(식품/음료) 분류표는 **발주서출력업무**와 공유하므로 그 페이지에서 관리하세요.")


def _commit_csv(path: str, df: pd.DataFrame, message: str):
    token = st.secrets.get("GITHUB_PAT", "")
    if not token:
        st.error("GITHUB_PAT 시크릿이 설정되지 않았습니다.")
        return False
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


def _section(title: str, filename: str, expected_cols: list[str], commit_label: str):
    st.subheader(title)
    local = _REF / filename
    if local.exists():
        cur = pd.read_csv(local, encoding="utf-8-sig", dtype=str)
        st.caption(f"현재 {len(cur)}행")
        st.dataframe(cur.head(50), use_container_width=True, height=240)
    else:
        st.warning(f"{filename} 없음")

    up = st.file_uploader(f"새 {filename} 업로드 (.csv, 전체 교체)",
                          type=["csv"], key=f"up_{filename}")
    if up:
        try:
            new = pd.read_csv(up, encoding="utf-8-sig", dtype=str)
        except UnicodeDecodeError:
            up.seek(0)
            new = pd.read_csv(up, dtype=str)
        miss = [c for c in expected_cols if c not in new.columns]
        if miss:
            st.error(f"필수 컬럼 누락: {miss} · 현재 컬럼: {list(new.columns)}")
        else:
            st.success(f"미리보기 {len(new)}행")
            st.dataframe(new.head(20), use_container_width=True)
            if st.button(f"💾 {title} 저장(커밋)", key=f"save_{filename}", type="primary"):
                if _commit_csv(f"reference/{filename}", new, commit_label):
                    st.success("✅ 저장 완료 — 1~2분 후 재배포 반영.")


_section("배민상회 수수료율", "bm_commission.csv",
         ["관리코드", "수수료율"], "ref: 배민상회 수수료율 갱신")
st.divider()
_section("소분목록", "sub_list.csv",
         ["관리코드", "낱개개수", "원코드"], "ref: 소분목록 갱신")
