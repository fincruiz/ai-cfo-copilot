"""AI CFO Copilot safe design system.

This file intentionally avoids styling Streamlit internals like file uploaders,
expanders, radio buttons and dataframes. It only styles custom HTML classes that
we render ourselves.
"""
from __future__ import annotations
import streamlit as st

BRAND = {
    "bg": "#0B0F17", "panel": "#111827", "border": "#2B3A55",
    "text": "#F9FAFB", "muted": "#CBD5E1",
    "primary": "#2563EB", "primary_2": "#0EA5E9",
    "success": "#22C55E", "warning": "#F59E0B", "danger": "#EF4444",
}

def apply_design_system_safe() -> None:
    """Apply only safe global styles and custom class styles."""
    st.markdown(f"""
    <style>
        html, body, [class*="css"] {{ font-family: Arial, sans-serif; }}
        .block-container {{ padding-top: 2.2rem; padding-bottom: 4rem; max-width: 1500px; }}

        .aicfo-hero {{
            border: 1px solid {BRAND["border"]};
            background: linear-gradient(135deg, rgba(37,99,235,.28), rgba(14,165,233,.14)),
                        linear-gradient(135deg, #101827, #07111f);
            border-radius: 28px; padding: 34px; margin: 16px 0 24px 0;
            box-shadow: 0 20px 60px rgba(0,0,0,.28);
        }}
        .aicfo-hero-title {{
            color: {BRAND["text"]}; font-size: 44px; line-height: 1.05;
            font-weight: 800; letter-spacing: -0.04em; margin: 0 0 12px 0;
        }}
        .aicfo-hero-subtitle {{
            color: {BRAND["muted"]}; font-size: 18px; line-height: 1.6;
            max-width: 900px; margin: 0;
        }}
        .aicfo-section-title {{ display: flex; align-items: center; gap: 12px; margin: 22px 0 14px 0; }}
        .aicfo-section-icon {{
            width: 44px; height: 44px; border-radius: 14px; display: flex;
            align-items: center; justify-content: center;
            background: linear-gradient(135deg, #2563EB, #0EA5E9);
            color: white; font-size: 22px; box-shadow: 0 12px 28px rgba(37,99,235,.25);
        }}
        .aicfo-section-title h2 {{
            margin: 0; color: {BRAND["text"]}; font-size: 28px;
            font-weight: 800; letter-spacing: -0.02em;
        }}
        .aicfo-card {{
            background: linear-gradient(180deg, rgba(17,24,39,.98), rgba(15,23,42,.96));
            border: 1px solid {BRAND["border"]}; border-radius: 22px;
            padding: 22px; min-height: 150px; box-shadow: 0 16px 40px rgba(0,0,0,.22);
        }}
        .aicfo-card h3 {{ color: {BRAND["text"]}; margin: 0 0 8px 0; font-size: 20px; font-weight: 800; }}
        .aicfo-card p {{ color: {BRAND["muted"]}; font-size: 15px; line-height: 1.55; margin: 0 0 10px 0; }}
        .aicfo-badge {{
            display: inline-block; padding: 6px 10px; border-radius: 999px;
            font-size: 12px; font-weight: 700; color: white;
            background: rgba(37,99,235,.9); margin-bottom: 12px;
        }}
        .aicfo-badge-muted {{ background: rgba(148,163,184,.22); color: #E5E7EB; border: 1px solid rgba(148,163,184,.28); }}
        .aicfo-badge-success {{ background: rgba(34,197,94,.18); color: #BBF7D0; border: 1px solid rgba(34,197,94,.35); }}
        .aicfo-badge-warning {{ background: rgba(245,158,11,.16); color: #FDE68A; border: 1px solid rgba(245,158,11,.35); }}

        .aicfo-kpi {{
            background: rgba(17,24,39,.92); border: 1px solid {BRAND["border"]};
            border-radius: 18px; padding: 18px; min-height: 118px;
        }}
        .aicfo-kpi-label {{
            color: {BRAND["muted"]}; font-size: 13px; font-weight: 700;
            text-transform: uppercase; letter-spacing: .06em; margin-bottom: 10px;
        }}
        .aicfo-kpi-value {{ color: {BRAND["text"]}; font-size: 28px; font-weight: 800; letter-spacing: -0.03em; }}
        .aicfo-kpi-note {{ color: {BRAND["muted"]}; font-size: 13px; margin-top: 8px; }}

        .aicfo-progress-wrap {{
            background: rgba(15,23,42,.8); border: 1px solid {BRAND["border"]};
            border-radius: 22px; padding: 20px; margin: 20px 0;
        }}
        .aicfo-progress-row {{ display: flex; gap: 10px; flex-wrap: wrap; margin-top: 14px; }}
        .aicfo-step {{
            flex: 1; min-width: 145px; background: rgba(17,24,39,.9);
            border: 1px solid {BRAND["border"]}; color: {BRAND["muted"]};
            border-radius: 16px; padding: 14px 12px; text-align: center; font-weight: 800;
        }}
        .aicfo-step-done {{
            background: linear-gradient(135deg, rgba(37,99,235,.95), rgba(14,165,233,.78));
            color: white; border-color: rgba(147,197,253,.65);
        }}
        .aicfo-alert {{
            border-radius: 18px; padding: 18px 20px;
            border: 1px solid rgba(245,158,11,.28);
            background: rgba(245,158,11,.10); color: #FDE68A; margin: 16px 0;
        }}
    </style>
    """, unsafe_allow_html=True)

def render_section_title(icon: str, title: str) -> None:
    st.markdown(f"""
    <div class="aicfo-section-title">
        <div class="aicfo-section-icon">{icon}</div>
        <h2>{title}</h2>
    </div>
    """, unsafe_allow_html=True)

def render_hero(title: str, subtitle: str) -> None:
    st.markdown(f"""
    <div class="aicfo-hero">
        <div class="aicfo-hero-title">{title}</div>
        <p class="aicfo-hero-subtitle">{subtitle}</p>
    </div>
    """, unsafe_allow_html=True)

def render_card(title: str, body: str, badge: str | None = None, badge_class: str = "") -> None:
    badge_html = f'<span class="aicfo-badge {badge_class}">{badge}</span>' if badge else ""
    st.markdown(f"""
    <div class="aicfo-card">
        {badge_html}
        <h3>{title}</h3>
        <p>{body}</p>
    </div>
    """, unsafe_allow_html=True)

def render_kpi(label: str, value: str, note: str = "") -> None:
    st.markdown(f"""
    <div class="aicfo-kpi">
        <div class="aicfo-kpi-label">{label}</div>
        <div class="aicfo-kpi-value">{value}</div>
        <div class="aicfo-kpi-note">{note}</div>
    </div>
    """, unsafe_allow_html=True)

def render_workflow_progress(done_steps: list[str] | None = None) -> None:
    done_steps = set(done_steps or [])
    steps = [("Configure", "1"), ("Import", "2"), ("Validate", "3"), ("Reports", "4"), ("AI Review", "5"), ("Forecast", "6"), ("Board Pack", "7")]
    html = '<div class="aicfo-progress-wrap"><b style="color:white;font-size:18px;">Month-End Close Workspace</b><div class="aicfo-progress-row">'
    for name, number in steps:
        cls = "aicfo-step aicfo-step-done" if name in done_steps else "aicfo-step"
        prefix = "✓" if name in done_steps else number
        html += f'<div class="{cls}">{prefix} {name}</div>'
    html += "</div></div>"
    st.markdown(html, unsafe_allow_html=True)
