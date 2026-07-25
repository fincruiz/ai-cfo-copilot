"""Premium V1 UI layer for AI CFO Copilot.

This file is intentionally self-contained and safe: it does not change finance
calculation logic, upload logic, validation logic, or reporting functions.

Use from app.py:
    from ui.premium_ui import apply_premium_theme, render_login_gate, render_premium_home

After st.set_page_config(...):
    apply_premium_theme()

Near the top of app.py before normal navigation/content:
    if not render_login_gate():
        st.stop()

Inside your Home page block, call:
    render_premium_home(profile, readiness_score, page_setter_callback)

If you do not have a page callback, pass None.
"""
from __future__ import annotations

from datetime import datetime
from typing import Callable, Optional

import streamlit as st


BRAND = {
    "bg": "#080D16",
    "panel": "#0F172A",
    "panel2": "#111827",
    "card": "#121A2A",
    "card2": "#162033",
    "border": "rgba(148, 163, 184, 0.22)",
    "text": "#F8FAFC",
    "muted": "#94A3B8",
    "soft": "#CBD5E1",
    "blue": "#2563EB",
    "cyan": "#06B6D4",
    "purple": "#7C3AED",
    "green": "#22C55E",
    "amber": "#F59E0B",
    "red": "#EF4444",
}


def apply_premium_theme() -> None:
    """Apply safe CSS that avoids touching Streamlit file uploader internals."""
    st.markdown(
        f"""
        <style>
        :root {{
            --bg: {BRAND['bg']};
            --panel: {BRAND['panel']};
            --panel2: {BRAND['panel2']};
            --card: {BRAND['card']};
            --border: {BRAND['border']};
            --text: {BRAND['text']};
            --muted: {BRAND['muted']};
            --blue: {BRAND['blue']};
            --cyan: {BRAND['cyan']};
            --purple: {BRAND['purple']};
            --green: {BRAND['green']};
        }}

        .stApp {{
            background:
                radial-gradient(circle at 15% 5%, rgba(37, 99, 235, 0.16), transparent 28%),
                radial-gradient(circle at 85% 15%, rgba(124, 58, 237, 0.14), transparent 30%),
                linear-gradient(180deg, #07101F 0%, #080D16 42%, #050914 100%);
            color: var(--text);
        }}

        .block-container {{
            padding-top: 2.2rem !important;
            padding-bottom: 4rem !important;
            max-width: 1400px !important;
        }}

        .premium-shell {{
            border: 1px solid var(--border);
            border-radius: 28px;
            background: linear-gradient(145deg, rgba(15, 23, 42, 0.92), rgba(17, 24, 39, 0.74));
            box-shadow: 0 26px 90px rgba(0, 0, 0, 0.38);
            overflow: hidden;
            margin-bottom: 24px;
        }}

        .premium-hero {{
            padding: 40px;
            background:
                linear-gradient(135deg, rgba(37, 99, 235, 0.26), rgba(124, 58, 237, 0.16)),
                radial-gradient(circle at 85% 10%, rgba(6, 182, 212, 0.20), transparent 35%);
        }}

        .premium-kicker {{
            display: inline-flex;
            gap: 8px;
            align-items: center;
            color: #DBEAFE;
            background: rgba(59, 130, 246, 0.16);
            border: 1px solid rgba(147, 197, 253, 0.24);
            border-radius: 999px;
            padding: 8px 13px;
            font-size: 13px;
            font-weight: 700;
            margin-bottom: 18px;
        }}

        .premium-title {{
            font-size: clamp(34px, 5vw, 64px);
            line-height: 0.98;
            letter-spacing: -0.055em;
            font-weight: 900;
            color: #FFFFFF;
            margin: 0 0 18px 0;
        }}

        .premium-subtitle {{
            color: #DDE7F6;
            font-size: 18px;
            line-height: 1.55;
            max-width: 780px;
            margin-bottom: 28px;
        }}

        .premium-grid {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 16px;
            margin-top: 26px;
        }}

        .premium-card {{
            border: 1px solid var(--border);
            background: rgba(15, 23, 42, 0.72);
            border-radius: 22px;
            padding: 20px;
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.03);
        }}

        .premium-card-label {{
            color: var(--muted);
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-weight: 800;
            margin-bottom: 9px;
        }}

        .premium-card-value {{
            color: var(--text);
            font-size: 25px;
            font-weight: 900;
            letter-spacing: -0.03em;
        }}

        .premium-section-title {{
            font-size: 24px;
            font-weight: 900;
            color: #FFFFFF;
            letter-spacing: -0.035em;
            margin: 10px 0 8px 0;
        }}

        .premium-section-subtitle {{
            color: var(--muted);
            font-size: 15px;
            margin-bottom: 16px;
        }}

        .choice-card {{
            min-height: 270px;
            border: 1px solid var(--border);
            border-radius: 26px;
            background: linear-gradient(160deg, rgba(18, 26, 42, 0.94), rgba(15, 23, 42, 0.82));
            padding: 26px;
            transition: all 160ms ease;
            box-shadow: 0 18px 60px rgba(0, 0, 0, 0.22);
        }}
        .choice-card:hover {{
            transform: translateY(-3px);
            border-color: rgba(96, 165, 250, 0.55);
            box-shadow: 0 24px 80px rgba(37, 99, 235, 0.18);
        }}
        .choice-badge {{
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            background: rgba(34, 197, 94, 0.12);
            color: #BBF7D0;
            border: 1px solid rgba(34, 197, 94, 0.22);
            font-size: 12px;
            font-weight: 800;
            margin-bottom: 14px;
        }}
        .choice-badge-warn {{
            background: rgba(245, 158, 11, 0.12);
            color: #FDE68A;
            border-color: rgba(245, 158, 11, 0.24);
        }}
        .choice-title {{
            color: #FFFFFF;
            font-size: 24px;
            font-weight: 900;
            margin-bottom: 8px;
            letter-spacing: -0.035em;
        }}
        .choice-copy {{
            color: #CBD5E1;
            font-size: 15px;
            line-height: 1.55;
            margin-bottom: 18px;
        }}
        .connector-row {{
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 14px;
        }}
        .connector-pill {{
            border: 1px solid rgba(148, 163, 184, 0.22);
            background: rgba(2, 6, 23, 0.28);
            color: #E2E8F0;
            padding: 7px 10px;
            border-radius: 999px;
            font-size: 12px;
            font-weight: 700;
        }}

        .progress-track {{
            display: grid;
            grid-template-columns: repeat(6, minmax(0, 1fr));
            gap: 12px;
            margin-top: 14px;
        }}
        .progress-step {{
            position: relative;
            border: 1px solid rgba(148, 163, 184, 0.18);
            background: rgba(15, 23, 42, 0.82);
            border-radius: 18px;
            padding: 16px;
            min-height: 92px;
        }}
        .progress-step.done {{
            border-color: rgba(34, 197, 94, 0.35);
            background: linear-gradient(160deg, rgba(34, 197, 94, 0.13), rgba(15, 23, 42, 0.82));
        }}
        .progress-step.current {{
            border-color: rgba(59, 130, 246, 0.54);
            background: linear-gradient(160deg, rgba(37, 99, 235, 0.20), rgba(15, 23, 42, 0.86));
        }}
        .progress-icon {{
            font-size: 20px;
            margin-bottom: 10px;
        }}
        .progress-label {{
            color: #F8FAFC;
            font-weight: 900;
            font-size: 14px;
        }}
        .progress-status {{
            color: #94A3B8;
            font-size: 12px;
            margin-top: 4px;
        }}

        .quick-action-grid {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 14px;
            margin-top: 12px;
        }}
        .quick-card {{
            border: 1px solid rgba(148, 163, 184, 0.20);
            background: rgba(15, 23, 42, 0.78);
            border-radius: 20px;
            padding: 18px;
            min-height: 118px;
        }}
        .quick-card strong {{
            display: block;
            color: #FFFFFF;
            font-size: 16px;
            margin-bottom: 8px;
        }}
        .quick-card span {{
            color: #94A3B8;
            font-size: 13px;
            line-height: 1.4;
        }}

        .ai-brief {{
            border: 1px solid rgba(124, 58, 237, 0.35);
            background:
                linear-gradient(135deg, rgba(124, 58, 237, 0.20), rgba(37, 99, 235, 0.10)),
                rgba(15, 23, 42, 0.84);
            border-radius: 24px;
            padding: 24px;
            min-height: 230px;
        }}
        .ai-line {{
            display: flex;
            justify-content: space-between;
            gap: 12px;
            border-bottom: 1px solid rgba(148, 163, 184, 0.14);
            padding: 10px 0;
            color: #E2E8F0;
            font-size: 14px;
        }}
        .ai-line:last-child {{ border-bottom: none; }}
        .ai-positive {{ color: #86EFAC; font-weight: 800; }}
        .ai-warning {{ color: #FDE68A; font-weight: 800; }}
        .ai-muted {{ color: #CBD5E1; font-weight: 800; }}

        .login-page {{
            min-height: 84vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }}
        .login-card {{
            width: min(1100px, 100%);
            border: 1px solid var(--border);
            background:
                radial-gradient(circle at 10% 10%, rgba(37, 99, 235, 0.24), transparent 34%),
                linear-gradient(145deg, rgba(15, 23, 42, 0.95), rgba(2, 6, 23, 0.92));
            border-radius: 34px;
            padding: 42px;
            box-shadow: 0 34px 120px rgba(0, 0, 0, 0.42);
        }}
        .login-brand {{
            font-size: 54px;
            line-height: 1.0;
            font-weight: 950;
            letter-spacing: -0.06em;
            color: white;
            margin-bottom: 14px;
        }}
        .login-copy {{
            color: #CBD5E1;
            font-size: 18px;
            line-height: 1.6;
            max-width: 560px;
        }}
        .login-panel {{
            border: 1px solid rgba(148, 163, 184, 0.24);
            background: rgba(15, 23, 42, 0.72);
            border-radius: 26px;
            padding: 24px;
        }}

        @media (max-width: 980px) {{
            .premium-grid, .quick-action-grid {{ grid-template-columns: 1fr; }}
            .progress-track {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
            .premium-hero {{ padding: 28px; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _get_profile_value(profile: dict | None, *keys: str, default: str = "Not set") -> str:
    profile = profile or {}
    for key in keys:
        value = profile.get(key)
        if value not in [None, "", "nan"]:
            return str(value)
    return default


def _set_page(page: str, page_callback: Optional[Callable[[str], None]] = None) -> None:
    if page_callback:
        page_callback(page)
    else:
        st.session_state["selected_page"] = page


def render_login_gate() -> bool:
    """Simple premium demo login gate. Returns True when app can continue."""
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if st.session_state.logged_in:
        return True

    st.markdown(
        """
        <div class="login-page">
          <div class="login-card">
            <div class="premium-kicker">✨ AI-powered Finance Operating System</div>
            <div class="login-brand">AI CFO Copilot</div>
            <div class="login-copy">
              From ERP exports to board-ready reporting, KPI packs, forecasting,
              benchmarking and AI CFO commentary in one guided workspace.
            </div>
            <br>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1.2, 1, 1.2])
    with col2:
        st.markdown("### Sign in")
        email = st.text_input("Email", value="demo@aicfocopilot.com", key="login_email")
        _ = st.text_input("Password", value="demo123", type="password", key="login_password")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Sign in", use_container_width=True, type="primary"):
                st.session_state.logged_in = True
                st.session_state.user_email = email
                st.rerun()
        with c2:
            if st.button("Try demo", use_container_width=True):
                st.session_state.logged_in = True
                st.session_state.user_email = "demo@aicfocopilot.com"
                st.rerun()
        st.caption("Demo login only. Real authentication can be added later with Microsoft/Google/Supabase/Clerk.")

    return False


def render_onboarding_modal_if_needed(page_callback: Optional[Callable[[str], None]] = None) -> None:
    """Show ERP/manual onboarding in a dialog on first visit."""
    if st.session_state.get("onboarding_completed"):
        return

    if hasattr(st, "dialog"):
        @st.dialog("Welcome to AI CFO Copilot")
        def _onboarding_dialog():
            _render_onboarding_content(page_callback)
        _onboarding_dialog()
    else:
        with st.expander("Welcome to AI CFO Copilot", expanded=True):
            _render_onboarding_content(page_callback)


def _render_onboarding_content(page_callback: Optional[Callable[[str], None]] = None) -> None:
    st.markdown("### How would you like to bring your finance data in?")
    st.caption("ERP connection is the preferred future flow. Manual upload remains fully supported for clients who prefer not to connect their ERP.")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            """
            <div class="choice-card">
              <div class="choice-badge">Recommended</div>
              <div class="choice-title">Connect ERP</div>
              <div class="choice-copy">Connect directly to your accounting or ERP system. Connectors are being prepared for the next release.</div>
              <div class="connector-row">
                <div class="connector-pill">Xero</div><div class="connector-pill">MYOB</div><div class="connector-pill">QuickBooks</div>
                <div class="connector-pill">Business Central</div><div class="connector-pill">NetSuite</div><div class="connector-pill">SAP</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Request ERP early access", use_container_width=True, key="onboard_erp"):
            st.session_state.import_method = "ERP Early Access"
            st.session_state.onboarding_completed = True
            _set_page("Import Centre", page_callback)
            st.rerun()

    with c2:
        st.markdown(
            """
            <div class="choice-card">
              <div class="choice-badge choice-badge-warn">Available Now</div>
              <div class="choice-title">Manual Import</div>
              <div class="choice-copy">Upload Excel or CSV exports from any ERP. This is ideal for demos, onboarding and clients who do not want API access.</div>
              <div class="connector-row">
                <div class="connector-pill">GL</div><div class="connector-pill">COA</div><div class="connector-pill">Budget</div>
                <div class="connector-pill">Forecast</div><div class="connector-pill">AR/AP</div><div class="connector-pill">KPI Master</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Continue with Manual Import", use_container_width=True, key="onboard_manual", type="primary"):
            st.session_state.import_method = "Manual Upload"
            st.session_state.onboarding_completed = True
            _set_page("Import Centre", page_callback)
            st.rerun()

    if st.button("Skip for now", use_container_width=True, key="onboard_skip"):
        st.session_state.onboarding_completed = True
        st.rerun()


def render_premium_navigation(current_page: str = "Home", page_callback: Optional[Callable[[str], None]] = None) -> str:
    """Modern top navigation. Returns selected page."""
    pages = [
        ("Home", "🏠"),
        ("Import Centre", "📥"),
        ("Dashboard", "📊"),
        ("Reports", "📈"),
        ("Performance", "⚡"),
        ("Working Capital Centre", "💰"),
        ("AI CFO", "🤖"),
        ("Downloads", "📤"),
    ]
    current_page = st.session_state.get("selected_page", current_page)
    cols = st.columns(len(pages))
    for col, (page, icon) in zip(cols, pages):
        with col:
            label = f"✓ {icon} {page}" if page == current_page else f"{icon} {page}"
            if st.button(label, use_container_width=True, key=f"nav_{page}"):
                _set_page(page, page_callback)
                st.rerun()
    return st.session_state.get("selected_page", current_page)


def render_premium_home(profile: dict | None = None, readiness_score: int = 0, page_callback: Optional[Callable[[str], None]] = None) -> None:
    """Premium Home workspace."""
    profile = profile or {}
    company = _get_profile_value(profile, "Company Name", "Company", default="Your Company")
    industry = _get_profile_value(profile, "Industry", default="Industry not set")
    country = _get_profile_value(profile, "Country", default="Country not set")
    period = _get_profile_value(profile, "Report Period", "Financial Year", default="Period not set")
    user = st.session_state.get("user_email", "")
    first_name = user.split("@")[0].split(".")[0].title() if user else "there"
    readiness_score = int(readiness_score or st.session_state.get("readiness_score", 0) or 0)

    render_onboarding_modal_if_needed(page_callback)

    st.markdown(
        f"""
        <div class="premium-shell">
          <div class="premium-hero">
            <div class="premium-kicker">✨ AI-powered Finance Operating System</div>
            <div class="premium-title">Good morning, {first_name}.<br/>Run month-end from one workspace.</div>
            <div class="premium-subtitle">
              Import ERP data, validate mappings, generate financial statements, review KPIs,
              compare budgets and ask your AI CFO for decision-ready commentary.
            </div>
            <div class="premium-grid">
              <div class="premium-card"><div class="premium-card-label">Company</div><div class="premium-card-value">{company}</div></div>
              <div class="premium-card"><div class="premium-card-label">Period</div><div class="premium-card-value">{period}</div></div>
              <div class="premium-card"><div class="premium-card-label">Readiness</div><div class="premium-card-value">{readiness_score}/100</div></div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<div class='premium-section-title'>Start here</div>", unsafe_allow_html=True)
    st.markdown("<div class='premium-section-subtitle'>Choose ERP connection for the roadmap, or continue with manual import today.</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            """
            <div class="choice-card">
              <div class="choice-badge">V1.1 Roadmap</div>
              <div class="choice-title">Connect ERP</div>
              <div class="choice-copy">Direct connectors will reduce manual uploads and refresh reporting automatically.</div>
              <div class="connector-row">
                <div class="connector-pill">Xero</div><div class="connector-pill">QuickBooks</div><div class="connector-pill">MYOB</div>
                <div class="connector-pill">Business Central</div><div class="connector-pill">NetSuite</div><div class="connector-pill">SAP</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("View ERP Hub", use_container_width=True):
            st.session_state.import_method = "ERP Early Access"
            _set_page("Import Centre", page_callback)
            st.rerun()
    with c2:
        st.markdown(
            """
            <div class="choice-card">
              <div class="choice-badge choice-badge-warn">Available Now</div>
              <div class="choice-title">Manual Import</div>
              <div class="choice-copy">Upload GL, COA, budgets, forecasts, AR/AP ageing and prior period files using controlled templates.</div>
              <div class="connector-row">
                <div class="connector-pill">Excel</div><div class="connector-pill">CSV</div><div class="connector-pill">SAP export</div>
                <div class="connector-pill">Oracle export</div><div class="connector-pill">Tally export</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Open Import Centre", use_container_width=True, type="primary"):
            st.session_state.import_method = "Manual Upload"
            _set_page("Import Centre", page_callback)
            st.rerun()

    st.markdown("<br/>", unsafe_allow_html=True)
    st.markdown("<div class='premium-section-title'>Month-end close workspace</div>", unsafe_allow_html=True)
    _render_progress_workspace(profile, readiness_score)

    left, right = st.columns([1.05, 0.95])
    with left:
        st.markdown("<div class='premium-section-title'>Quick actions</div>", unsafe_allow_html=True)
        _render_quick_actions(page_callback)
    with right:
        st.markdown("<div class='premium-section-title'>Today's AI CFO brief</div>", unsafe_allow_html=True)
        _render_ai_brief(readiness_score)

    st.markdown("<br/>", unsafe_allow_html=True)
    _render_recent_activity(company, industry, country)


def _render_progress_workspace(profile: dict, readiness_score: int) -> None:
    company_set = bool(_get_profile_value(profile, "Company Name", default="") not in ["", "Not set"])
    files_loaded = bool(st.session_state.get("data_loaded") or st.session_state.get("current_gl_uploaded") or st.session_state.get("gl_df") is not None)
    validated = readiness_score >= 70 or bool(st.session_state.get("validation_completed"))
    reports = bool(st.session_state.get("consolidated_pnl") is not None or st.session_state.get("reports_generated"))
    ai = bool(st.session_state.get("ai_commentary") or st.session_state.get("ai_review_done"))
    board = bool(st.session_state.get("board_pack_generated"))

    steps = [
        ("Configure", "Company profile", company_set, not company_set, "🏢"),
        ("Import", "GL / COA / optional files", files_loaded, company_set and not files_loaded, "📥"),
        ("Validate", "Quality and mapping checks", validated, files_loaded and not validated, "✅"),
        ("Reports", "P&L / BS / KPIs", reports, validated and not reports, "📊"),
        ("AI Review", "CFO commentary", ai, reports and not ai, "🤖"),
        ("Board Pack", "Export and share", board, ai and not board, "📤"),
    ]

    html = '<div class="progress-track">'
    for label, caption, done, current, icon in steps:
        klass = "progress-step done" if done else "progress-step current" if current else "progress-step"
        status = "Complete" if done else "Current" if current else "Pending"
        html += f"""
        <div class="{klass}">
          <div class="progress-icon">{icon}</div>
          <div class="progress-label">{label}</div>
          <div class="progress-status">{caption}<br/>{status}</div>
        </div>
        """
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def _render_quick_actions(page_callback: Optional[Callable[[str], None]]) -> None:
    st.markdown(
        """
        <div class="quick-action-grid">
          <div class="quick-card"><strong>📥 Import files</strong><span>Upload GL, COA, budget, forecast and working capital files.</span></div>
          <div class="quick-card"><strong>📊 Generate reports</strong><span>Create financial statements and management packs.</span></div>
          <div class="quick-card"><strong>🤖 Run AI CFO review</strong><span>Generate executive commentary and action points.</span></div>
          <div class="quick-card"><strong>📈 Budget comparison</strong><span>Review actuals against budget and forecast.</span></div>
          <div class="quick-card"><strong>💰 Working capital</strong><span>Analyse receivables, payables, DSO and DPO.</span></div>
          <div class="quick-card"><strong>📤 Download pack</strong><span>Export Excel, PDF and board-ready outputs.</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("Open Import Centre", use_container_width=True, key="qa_import"):
            _set_page("Import Centre", page_callback)
            st.rerun()
    with c2:
        if st.button("Open Reports", use_container_width=True, key="qa_reports"):
            _set_page("Reports", page_callback)
            st.rerun()
    with c3:
        if st.button("Open AI CFO", use_container_width=True, key="qa_ai"):
            st.session_state.chat_open = True
            _set_page("AI CFO", page_callback)
            st.rerun()


def _render_ai_brief(readiness_score: int) -> None:
    if readiness_score >= 80:
        readiness_status = "Healthy"
        readiness_class = "ai-positive"
    elif readiness_score >= 50:
        readiness_status = "Needs review"
        readiness_class = "ai-warning"
    else:
        readiness_status = "Setup required"
        readiness_class = "ai-warning"

    st.markdown(
        f"""
        <div class="ai-brief">
          <div class="premium-kicker">🧠 AI CFO</div>
          <div class="choice-title">Executive brief</div>
          <div class="choice-copy">Once files are validated, AI CFO will summarise what changed, why it changed and what management should do next.</div>
          <div class="ai-line"><span>Data readiness</span><span class="{readiness_class}">{readiness_status}</span></div>
          <div class="ai-line"><span>Revenue movement</span><span class="ai-muted">Waiting for reports</span></div>
          <div class="ai-line"><span>Margin risk</span><span class="ai-muted">Waiting for P&L</span></div>
          <div class="ai-line"><span>Working capital</span><span class="ai-muted">Waiting for AR/AP</span></div>
          <div class="ai-line"><span>Next action</span><span class="ai-warning">Complete import</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _render_recent_activity(company: str, industry: str, country: str) -> None:
    now = datetime.now().strftime("%d %b %Y")
    st.markdown("<div class='premium-section-title'>Workspace snapshot</div>", unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="premium-grid">
          <div class="premium-card"><div class="premium-card-label">Company</div><div class="premium-card-value">{company}</div></div>
          <div class="premium-card"><div class="premium-card-label">Industry / Country</div><div class="premium-card-value">{industry} · {country}</div></div>
          <div class="premium-card"><div class="premium-card-label">Last opened</div><div class="premium-card-value">{now}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_import_centre_header(profile: dict | None = None) -> None:
    profile = profile or {}
    company = _get_profile_value(profile, "Company Name", "Company", default="Company not set")
    period = _get_profile_value(profile, "Report Period", "Financial Year", default="Period not set")
    method = st.session_state.get("import_method", "Manual Upload")
    st.markdown(
        f"""
        <div class="premium-shell">
          <div class="premium-hero" style="padding: 30px;">
            <div class="premium-kicker">📥 Import Centre</div>
            <div class="premium-title" style="font-size: 42px;">Bring finance data into the workspace.</div>
            <div class="premium-subtitle">Current company: <b>{company}</b> · Period: <b>{period}</b> · Method: <b>{method}</b></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

