"""core/ui.py — 전역 UI 테마/헬퍼 (Phase A).

설계: 색·폰트·라운드는 .streamlit/config.toml(전역 자동 상속).
이 모듈은 그 위에 (1) 전 페이지 공통 CSS 폴리시 (2) 재사용 컴포넌트 헬퍼를 얹는다.

사용:
- streamlit_app.py 진입점에서 `inject_css()` 1회 → 엔트리가 매 페이지 로드마다
  먼저 실행되므로 전 페이지에 적용.
- 페이지에서: `from core.ui import page_header, kpi_row, status_pill, delta_html`

토큰은 config.toml 테마와 동일 값으로 유지(시각 일관).
"""
import streamlit as st

# ── 디자인 토큰 (config.toml 테마 미러) ──────────────────────────
ACCENT = "#3B5BDB"   # primaryColor
INK    = "#181B22"   # textColor
MUTED  = "#6B7280"
LINE   = "#E6E8EC"   # borderColor
UP     = "#E03131"   # 인상/상승 = 빨강 (한국식)
DOWN   = "#1971C2"   # 인하/하락 = 파랑 (한국식)
G, Y, R = "#2F9E44", "#F08C00", "#E03131"  # 🟢🟡🔴

_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600&display=swap');
/* ===== 본문 간격·타이포 ===== */
.block-container { padding-top: 2.2rem; padding-bottom: 4rem; max-width: none; }
h1 { letter-spacing: -0.02em; font-weight: 700; }
h2, h3 { letter-spacing: -0.01em; }

/* ===== 사이드바 ===== */
[data-testid="stSidebar"] { border-right: 1px solid #E6E8EC; }
.ui-brand { display:flex; align-items:center; gap:9px; font-weight:700; font-size:16px;
  letter-spacing:-0.02em; padding:4px 4px 12px; }
.ui-brand .logo { width:26px; height:26px; border-radius:7px; background:#3B5BDB;
  display:grid; place-items:center; color:#fff; font-size:15px; line-height:1; }

/* ===== st.metric → 카드 (기존 페이지 자동 적용) ===== */
[data-testid="stMetric"], [data-testid="metric-container"] {
  background:#FFFFFF; border:1px solid #E6E8EC; border-radius:12px; padding:14px 16px;
}
[data-testid="stMetricValue"], [data-testid="stMetricValue"] div {
  font-weight:700; letter-spacing:-0.02em; font-variant-numeric:tabular-nums;
}
[data-testid="stMetricLabel"] { color:#6B7280; }

/* ===== 버튼 ===== */
.stButton > button, .stDownloadButton > button { border-radius:8px; font-weight:600; }

/* ===== 탭 ===== */
[data-testid="stTabs"] button { font-weight:500; }

/* ===== 커스텀 헤더 ===== */
.ui-head { margin-bottom:18px; }
.ui-head h1 { display:flex; align-items:center; gap:10px; margin:0; }
.ui-head p { color:#6B7280; margin:5px 0 0; font-size:14px; }

/* ===== KPI 카드 (커스텀) ===== */
.ui-kpis { display:flex; gap:14px; flex-wrap:wrap; margin-bottom:8px; }
.ui-kpi { flex:1; min-width:150px; background:#FFFFFF; border:1px solid #E6E8EC;
  border-radius:12px; padding:14px 16px; }
.ui-kpi .lab { font-size:12.5px; color:#6B7280; font-weight:500;
  display:flex; align-items:center; gap:6px; }
.ui-kpi .val { font-size:26px; font-weight:700; margin-top:6px; letter-spacing:-0.02em;
  font-variant-numeric:tabular-nums; }
.ui-kpi .dlt { font-size:12px; font-weight:600; margin-top:3px; color:#6B7280; }

/* ===== 핀필 / 델타 (한국식 ▲빨강 ▼파랑) ===== */
.ui-dot { width:8px; height:8px; border-radius:50%; display:inline-block; }
.ui-pill { display:inline-flex; align-items:center; gap:6px; font-size:12.5px;
  font-weight:600; padding:3px 10px; border-radius:20px; }
.ui-pill.g { background:#EBFBEE; color:#2F9E44; }
.ui-pill.y { background:#FFF4E6; color:#F08C00; }
.ui-pill.r { background:#FFF0F0; color:#E03131; }
.ui-up { color:#E03131; font-weight:700; font-variant-numeric:tabular-nums; }
.ui-down { color:#1971C2; font-weight:700; font-variant-numeric:tabular-nums; }

/* ===== 섹션 헤더 (인디고 액센트 바) ===== */
.ui-eyebrow{font-family:"JetBrains Mono",ui-monospace,monospace;font-size:11.5px;font-weight:600;
  letter-spacing:.12em;text-transform:uppercase;color:#3B5BDB}
.ui-sec{display:flex;align-items:center;gap:11px;margin:10px 0 16px}
.ui-sec .ui-bar{width:4px;height:21px;border-radius:3px;background:#3B5BDB;flex:none}
.ui-sec h2{font-size:19px;font-weight:700;letter-spacing:-.01em;margin:0;line-height:1.2}
.ui-sec .ui-tag{font-family:"JetBrains Mono",ui-monospace,monospace;font-size:11px;color:#9AA1AC;
  margin-left:auto;font-weight:500;font-variant-numeric:tabular-nums}
</style>
"""


def inject_css() -> None:
    """전역 CSS 주입 — streamlit_app.py 진입점에서 1회 호출."""
    st.markdown(_CSS, unsafe_allow_html=True)


def page_header(title: str, sub: str | None = None, icon: str = "") -> None:
    """페이지 상단 일관 헤더 (제목 + 부제)."""
    ic = f"{icon} " if icon else ""
    sub_html = f"<p>{sub}</p>" if sub else ""
    st.markdown(
        f'<div class="ui-head"><h1>{ic}{title}</h1>{sub_html}</div>',
        unsafe_allow_html=True,
    )


def status_pill(text: str, level: str = "g") -> str:
    """상태 핀필 HTML 반환. level: 'g'|'y'|'r' (🟢🟡🔴)."""
    lv = level if level in ("g", "y", "r") else "g"
    return (
        f'<span class="ui-pill {lv}">'
        f'<span class="ui-dot" style="background:currentColor"></span>{text}</span>'
    )


def delta_html(text: str, direction: str = "up") -> str:
    """변화량 HTML (한국식: up=빨강 ▲, down=파랑 ▼)."""
    cls = "ui-up" if direction == "up" else "ui-down"
    return f'<span class="{cls}">{text}</span>'


def kpi_row(items: list[dict]) -> None:
    """KPI 카드 한 줄. items: [{label, value, delta?, delta_dir?, color?, dot?}].

    delta_dir: 'up'|'down'|None (한국식 색). dot: 색상 hex (라벨 앞 점).
    """
    cards = []
    for it in items:
        lab = it.get("label", "")
        val = it.get("value", "")
        dot = it.get("dot")
        dot_html = f'<span class="ui-dot" style="background:{dot}"></span>' if dot else ""
        color = it.get("color")
        vstyle = f' style="color:{color}"' if color else ""
        dlt = it.get("delta")
        dlt_html = ""
        if dlt:
            ddir = it.get("delta_dir")
            dcls = {"up": "ui-up", "down": "ui-down"}.get(ddir, "")
            dlt_html = f'<div class="dlt {dcls}">{dlt}</div>'
        cards.append(
            f'<div class="ui-kpi"><div class="lab">{dot_html}{lab}</div>'
            f'<div class="val"{vstyle}>{val}</div>{dlt_html}</div>'
        )
    st.markdown('<div class="ui-kpis">' + "".join(cards) + "</div>",
                unsafe_allow_html=True)


def section_head(title: str, icon: str = "", tag: str | None = None) -> None:
    """섹션 헤더 — 인디고 액센트 바 + 제목 (+선택 mono 태그)."""
    ic = f"{icon} " if icon else ""
    tag_html = f'<span class="ui-tag">{tag}</span>' if tag else ""
    st.markdown(
        f'<div class="ui-sec"><span class="ui-bar"></span>'
        f'<h2>{ic}{title}</h2>{tag_html}</div>',
        unsafe_allow_html=True,
    )
