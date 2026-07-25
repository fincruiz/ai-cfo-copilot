"""AI CFO Copilot design system.

This module contains reusable UI styling and small visual components.
It is intentionally lightweight so it can be added safely without changing
finance calculations or upload/reporting logic.
"""

from __future__ import annotations

import streamlit as st


BRAND = {
    "bg": "#0B0F17",
    "panel": "#111827",
    "panel_2": "#162033",
    "card": "#121A2A",
    "card_soft": "#172033",
    "border": "#2B3A55",
    "muted": "#9CA3AF",
    "text": "#F9FAFB",
    "subtle_text": "#CBD5E1",
    "primary": "#2563EB",
    "primary_2": "#0EA5E9",
    "success": "#22C55E",
    "warning": "#F59E0B",
    "danger": "#EF4444",
    "purple": "#8B5CF6",
}


def apply_design_system() -> None:
    """Apply global Streamlit CSS for a premium dark SaaS look."""
    st.markdown(
        f"""
        <style>
        :root {{
            --aicfo-bg: {BRAND['bg']};
            --aicfo-panel: {BRAND['panel']};
            --aicfo-card: {BRAND['card']};
            --aicfo-border: {BRAND['border']};
            --aicfo-text: {BRAND['text']};
            --aicfo-muted: {BRAND['muted']};
            --aicfo-primary: {BRAND['primary']};
            --aicfo-success: {BRAND['success']};
            --aicfo-warning: {BRAND['warning']};
            --aicfo-danger: {BRAND['danger']};
        }}

        html, body, [class*="css"] {{
            font-family: Arial, Helvetica, sans-serif !important;
        }}

        .stApp {{
            background: radial-gradient(circle at top left, #0F1B2D 0, #0B0F17 32%, #080B10 100%) !important;
            color: var(--aicfo-text) !important;
        }}

        h1, h2, h3, h4, h5, h6, p, span, label, div {{
            font-family: Arial, Helvetica, sans-serif !important;
        }}

        h1 {{
            color: var(--aicfo-text) !important;
            font-size: 40px !important;
            font-weight: 800 !important;
            letter-spacing: -0.04em !important;
        }}

        h2 {{
            color: var(--aicfo-text) !important;
            font-size: 28px !important;
            font-weight: 760 !important;
            letter-spacing: -0.03em !important;
        }}

        h3 {{
            color: var(--aicfo-text) !important;
            font-size: 21px !important;
            font-weight: 730 !important;
        }}

        p, label, .stMarkdown, .stTextInput label, .stSelectbox label, .stDateInput label {{
            color: var(--aicfo-subtle-text, #CBD5E1) !important;
        }}

        [data-testid="stHeader"] {{
            background: rgba(11, 15, 23, 0.82) !important;
            backdrop-filter: blur(16px);
        }}

        [data-testid="stToolbar"] {{
            right: 1rem !important;
        }}

        .block-container {{
            padding-top: 2.2rem !important;
            padding-bottom: 4rem !important;
            max-width: 1500px !important;
        }}

        div[data-testid="stMetric"] {{
            background: linear-gradient(145deg, rgba(18,26,42,0.96), rgba(22,32,51,0.88));
            border: 1px solid rgba(148,163,184,0.18);
            border-radius: 18px;
            padding: 18px 18px;
            box-shadow: 0 18px 45px rgba(0,0,0,0.18);
        }}

        div[data-testid="stMetric"] label {{
            color: #A7B0C0 !important;
            font-weight: 700 !important;
        }}

        div[data-testid="stMetric"] [data-testid="stMetricValue"] {{
            color: #FFFFFF !important;
            font-weight: 800 !important;
        }}

        div[data-testid="stDataFrame"] {{
            border-radius: 16px !important;
            overflow: hidden !important;
            border: 1px solid rgba(148,163,184,0.18) !important;
        }}

        .stButton > button {{
            background: linear-gradient(135deg, #1E293B, #111827) !important;
            border: 1px solid #334155 !important;
            color: #FFFFFF !important;
            border-radius: 14px !important;
            padding: 0.72rem 1.05rem !important;
            font-weight: 750 !important;
            box-shadow: 0 10px 25px rgba(0,0,0,0.16) !important;
            transition: all 0.18s ease-in-out !important;
        }}

        .stButton > button:hover {{
            transform: translateY(-1px);
            border-color: #60A5FA !important;
            box-shadow: 0 14px 34px rgba(37,99,235,0.22) !important;
        }}

        .stDownloadButton > button {{
            background: linear-gradient(135deg, #2563EB, #0EA5E9) !important;
            border: 0 !important;
            color: #FFFFFF !important;
            border-radius: 14px !important;
            font-weight: 750 !important;
        }}

        .stTextInput input,
        .stNumberInput input,
        .stDateInput input,
        .stSelectbox div[data-baseweb="select"],
        .stTextArea textarea {{
            background: #0F172A !important;
            color: #FFFFFF !important;
            border-radius: 12px !important;
            border-color: #334155 !important;
        }}

        .stAlert {{
            border-radius: 16px !important;
            border: 1px solid rgba(255,255,255,0.10) !important;
        }}

        hr {{
            border-color: rgba(148,163,184,0.20) !important;
        }}

        .aicfo-topbar {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 20px;
            padding: 18px 0 28px 0;
            border-bottom: 1px solid rgba(148,163,184,0.18);
            margin-bottom: 22px;
        }}

        .aicfo-brand {{
            display: flex;
            align-items: center;
            gap: 14px;
        }}

        .aicfo-logo {{
            width: 52px;
            height: 52px;
            border-radius: 17px;
            background: linear-gradient(135deg, #0EA5E9, #2563EB, #7C3AED);
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 24px;
            font-weight: 900;
            box-shadow: 0 18px 38px rgba(37,99,235,0.28);
        }}

        .aicfo-title {{
            font-size: 30px;
            line-height: 1;
            font-weight: 850;
            color: #FFFFFF;
            letter-spacing: -0.04em;
        }}

        .aicfo-subtitle {{
            margin-top: 8px;
            font-size: 14px;
            color: #94A3B8;
            font-weight: 520;
        }}

        .aicfo-nav-wrap {{
            display: grid;
            grid-template-columns: repeat(7, minmax(120px, 1fr));
            gap: 14px;
            margin: 18px 0 26px 0;
        }}

        .aicfo-nav-pill {{
            background: linear-gradient(145deg, #172033, #111827);
            border: 1px solid #334155;
            color: #FFFFFF !important;
            padding: 14px 12px;
            border-radius: 16px;
            text-align: center;
            font-weight: 760;
            box-shadow: 0 12px 25px rgba(0,0,0,0.14);
        }}

        .aicfo-nav-pill.active {{
            background: linear-gradient(135deg, #2563EB, #1D4ED8);
            border-color: #60A5FA;
            box-shadow: 0 16px 40px rgba(37,99,235,0.30);
        }}

        .aicfo-card {{
            background: linear-gradient(145deg, rgba(18,26,42,0.96), rgba(15,23,42,0.92));
            border: 1px solid rgba(148,163,184,0.18);
            border-radius: 22px;
            padding: 24px;
            box-shadow: 0 18px 50px rgba(0,0,0,0.20);
        }}

        .aicfo-card-soft {{
            background: rgba(15,23,42,0.78);
            border: 1px solid rgba(148,163,184,0.15);
            border-radius: 18px;
            padding: 18px;
        }}

        .aicfo-hero {{
            min-height: 360px;
            border-radius: 28px;
            padding: 34px;
            position: relative;
            overflow: hidden;
            background:
                linear-gradient(135deg, rgba(14,116,144,0.84), rgba(37,99,235,0.74)),
                url('https://images.unsplash.com/photo-1554224155-6726b3ff858f?q=80&w=1800&auto=format&fit=crop');
            background-size: cover;
            background-position: center;
            border: 1px solid rgba(255,255,255,0.16);
            box-shadow: 0 22px 60px rgba(0,0,0,0.28);
        }}

        .aicfo-hero h1 {{
            color: white !important;
            max-width: 760px;
            font-size: 48px !important;
            line-height: 1.04 !important;
            margin-top: 20px !important;
        }}

        .aicfo-hero p {{
            color: #E0F2FE !important;
            max-width: 760px;
            font-size: 18px !important;
            line-height: 1.6 !important;
        }}

        .aicfo-badge {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 9px 13px;
            border-radius: 999px;
            background: rgba(255,255,255,0.12);
            border: 1px solid rgba(255,255,255,0.25);
            color: #FFFFFF;
            font-weight: 760;
            font-size: 13px;
        }}

        .aicfo-step-grid {{
            display: grid;
            grid-template-columns: repeat(5, minmax(130px, 1fr));
            gap: 14px;
            margin: 20px 0 24px 0;
        }}

        .aicfo-step {{
            border-radius: 18px;
            border: 1px solid rgba(148,163,184,0.20);
            background: linear-gradient(145deg, #111827, #0F172A);
            padding: 16px;
            color: #FFFFFF;
            font-weight: 800;
            text-align: center;
        }}

        .aicfo-step.done {{
            border-color: rgba(34,197,94,0.65);
            background: linear-gradient(135deg, rgba(22,101,52,0.92), rgba(21,128,61,0.76));
        }}

        .aicfo-step.active {{
            border-color: #60A5FA;
            background: linear-gradient(135deg, #1D4ED8, #2563EB);
            box-shadow: 0 16px 38px rgba(37,99,235,0.25);
        }}

        .aicfo-status-line {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            color: #CBD5E1;
            font-weight: 700;
        }}

        .aicfo-muted {{ color: #94A3B8; }}
        .aicfo-success {{ color: #86EFAC; }}
        .aicfo-warning {{ color: #FCD34D; }}
        .aicfo-danger {{ color: #FCA5A5; }}

        @media (max-width: 900px) {{
            .aicfo-nav-wrap {{ grid-template-columns: repeat(2, minmax(120px, 1fr)); }}
            .aicfo-step-grid {{ grid-template-columns: repeat(2, minmax(130px, 1fr)); }}
            .aicfo-hero h1 {{ font-size: 34px !important; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_app_header(title: str = "AI CFO Copilot", subtitle: str | None = None) -> None:
    subtitle = subtitle or "From ERP to board pack — one AI-powered finance workspace."
    st.markdown(
        f"""
        <div class="aicfo-topbar">
            <div class="aicfo-brand">
                <div class="aicfo-logo">▣</div>
                <div>
                    <div class="aicfo-title">{title}</div>
                    <div class="aicfo-subtitle">{subtitle}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_nav_pills(current_page: str, pages: list[str]) -> None:
    html = '<div class="aicfo-nav-wrap">'
    icon_map = {
        "Home": "🏠",
        "Import Centre": "📥",
        "Dashboard": "📊",
        "Reports": "📈",
        "Working Capital Centre": "💰",
        "Insights": "🧠",
        "Downloads": "📤",
        "AI CFO": "🤖",
        "Administration": "⚙️",
    }
    for page in pages:
        active = " active" if page == current_page else ""
        icon = icon_map.get(page, "•")
        html += f'<div class="aicfo-nav-pill{active}">{icon} {page}</div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_hero(title: str, body: str, badge: str = "AI-powered finance workspace") -> None:
    st.markdown(
        f"""
        <div class="aicfo-hero">
            <div class="aicfo-badge">✨ {badge}</div>
            <h1>{title}</h1>
            <p>{body}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_card(title: str, body: str = "", icon: str = "", footer: str = "") -> None:
    st.markdown(
        f"""
        <div class="aicfo-card">
            <h3>{icon} {title}</h3>
            <p>{body}</p>
            <div class="aicfo-muted">{footer}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_small_card(title: str, value: str, note: str = "", status: str = "") -> None:
    status_class = {
        "success": "aicfo-success",
        "warning": "aicfo-warning",
        "danger": "aicfo-danger",
    }.get(status, "aicfo-muted")
    st.markdown(
        f"""
        <div class="aicfo-card-soft">
            <div class="aicfo-muted" style="font-weight:800;font-size:13px;">{title}</div>
            <div style="font-size:28px;font-weight:900;color:white;margin-top:8px;">{value}</div>
            <div class="{status_class}" style="font-weight:700;margin-top:6px;">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_workflow_steps(steps: list[dict]) -> None:
    html = '<div class="aicfo-step-grid">'
    for step in steps:
        label = step.get("label", "Step")
        state = step.get("state", "")
        cls = "aicfo-step"
        if state == "done":
            cls += " done"
        elif state == "active":
            cls += " active"
        prefix = "✓" if state == "done" else "○" if state == "pending" else "→"
        html += f'<div class="{cls}">{prefix} {label}</div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_section_title(icon: str, title: str, subtitle: str = "") -> None:
    st.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:12px;margin:26px 0 16px 0;">
            <div style="width:42px;height:42px;border-radius:14px;background:linear-gradient(135deg,#0EA5E9,#2563EB);display:flex;align-items:center;justify-content:center;font-size:21px;">{icon}</div>
            <div>
                <div style="font-size:24px;font-weight:850;color:white;letter-spacing:-0.03em;">{title}</div>
                <div style="font-size:14px;color:#94A3B8;margin-top:3px;">{subtitle}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
