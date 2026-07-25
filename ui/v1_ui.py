import streamlit as st
from core.common import get_report_period_label


def inject_v1_polished_css():
    """Scoped V1 styling. Does not target Streamlit internal uploader/icon classes."""
    st.markdown(
        """
        <style>
        .aicfo-hero {
            position: relative;
            overflow: hidden;
            border-radius: 30px;
            padding: 2.25rem;
            margin: 1.1rem 0 1.2rem 0;
            background:
                radial-gradient(circle at 15% 15%, rgba(56,189,248,.30), transparent 28%),
                radial-gradient(circle at 85% 10%, rgba(124,58,237,.34), transparent 28%),
                linear-gradient(135deg, #07111f 0%, #0b1220 50%, #111827 100%);
            border: 1px solid rgba(148,163,184,.22);
            box-shadow: 0 30px 90px rgba(0,0,0,.34);
        }
        .aicfo-hero:after {
            content: "";
            position:absolute;
            right:-80px;
            top:-80px;
            width:280px;
            height:280px;
            border-radius:999px;
            background: linear-gradient(135deg, rgba(37,99,235,.40), rgba(14,165,233,.05));
            filter: blur(4px);
        }
        .aicfo-kicker {
            display:inline-flex;
            align-items:center;
            gap:.45rem;
            padding:.45rem .8rem;
            border-radius:999px;
            background:rgba(37,99,235,.18);
            border:1px solid rgba(96,165,250,.28);
            color:#bfdbfe;
            font-size:.82rem;
            font-weight:850;
            letter-spacing:.01em;
        }
        .aicfo-hero-title {
            color:#ffffff;
            font-size:3.15rem;
            line-height:1.02;
            letter-spacing:-.06em;
            font-weight:950;
            max-width:850px;
            margin:.75rem 0 .75rem 0;
        }
        .aicfo-hero-sub {
            color:#cbd5e1;
            font-size:1.05rem;
            line-height:1.55;
            max-width:760px;
            margin-bottom:1.15rem;
        }
        .aicfo-pill-row {display:flex;flex-wrap:wrap;gap:.55rem;margin-top:1rem;}
        .aicfo-pill {
            color:#e0f2fe;
            font-weight:800;
            font-size:.84rem;
            border-radius:999px;
            padding:.48rem .72rem;
            background:rgba(255,255,255,.08);
            border:1px solid rgba(255,255,255,.13);
        }
        .aicfo-grid-2 {display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:1rem;margin:1rem 0;}
        .aicfo-grid-3 {display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:1rem;margin:1rem 0;}
        .aicfo-grid-4 {display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:1rem;margin:1rem 0;}
        .aicfo-card {
            border-radius:24px;
            padding:1.15rem;
            background:linear-gradient(180deg, rgba(17,24,39,.97), rgba(15,23,42,.97));
            border:1px solid rgba(148,163,184,.24);
            box-shadow:0 18px 48px rgba(0,0,0,.23);
            color:#f8fafc;
        }
        .aicfo-card.soft {
            background:linear-gradient(180deg, rgba(15,23,42,.90), rgba(17,24,39,.92));
        }
        .aicfo-card h3 {margin:.25rem 0 .4rem 0;color:#fff;font-size:1.15rem;}
        .aicfo-card p, .aicfo-card li {color:#cbd5e1;font-size:.92rem;line-height:1.45;}
        .aicfo-tag {
            display:inline-flex;
            padding:.28rem .58rem;
            border-radius:999px;
            background:rgba(34,197,94,.12);
            border:1px solid rgba(34,197,94,.28);
            color:#86efac;
            font-size:.76rem;
            font-weight:850;
        }
        .aicfo-tag.blue {background:rgba(37,99,235,.15);border-color:rgba(96,165,250,.30);color:#bfdbfe;}
        .aicfo-tag.amber {background:rgba(245,158,11,.13);border-color:rgba(251,191,36,.28);color:#fde68a;}
        .aicfo-metric {
            border-radius:22px;
            padding:1rem;
            background:rgba(15,23,42,.88);
            border:1px solid rgba(148,163,184,.22);
        }
        .aicfo-metric-label {color:#94a3b8;font-size:.78rem;font-weight:850;margin-bottom:.25rem;}
        .aicfo-metric-value {color:#fff;font-size:1.48rem;font-weight:950;letter-spacing:-.04em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        .aicfo-section-title {display:flex;align-items:center;gap:.65rem;margin:1.4rem 0 .75rem 0;}
        .aicfo-section-icon {width:40px;height:40px;border-radius:14px;display:flex;align-items:center;justify-content:center;background:rgba(37,99,235,.15);border:1px solid rgba(96,165,250,.25);font-size:1.2rem;}
        .aicfo-section-text {font-size:1.35rem;font-weight:950;letter-spacing:-.035em;color:#f8fafc;}
        .aicfo-progress-wrap {height:13px;background:rgba(148,163,184,.22);border-radius:999px;overflow:hidden;border:1px solid rgba(148,163,184,.18);margin:.75rem 0 .35rem 0;}
        .aicfo-progress-bar {height:100%;border-radius:999px;background:linear-gradient(90deg,#14b8a6,#2563eb,#7c3aed);}
        .aicfo-step {border-radius:18px;padding:.88rem 1rem;background:rgba(15,23,42,.86);border:1px solid rgba(148,163,184,.22);min-height:88px;}
        .aicfo-step.done {background:rgba(6,78,59,.55);border-color:rgba(52,211,153,.46);}
        .aicfo-step.active {background:rgba(30,64,175,.55);border-color:rgba(96,165,250,.55);}
        .aicfo-step-title {color:#fff;font-weight:900;margin-bottom:.2rem;}
        .aicfo-step-sub {color:#cbd5e1;font-size:.82rem;line-height:1.3;}
        .aicfo-brief {
            border-radius:24px;
            padding:1.15rem;
            background:linear-gradient(135deg, rgba(37,99,235,.18), rgba(15,23,42,.95));
            border:1px solid rgba(96,165,250,.25);
        }
        .aicfo-brief h3 {color:#fff;margin:.1rem 0 .45rem 0;}
        .aicfo-brief p {color:#dbeafe;margin:.2rem 0;font-size:.93rem;}
        .aicfo-import-banner {
            border-radius:24px;
            padding:1.15rem 1.25rem;
            margin:.8rem 0 1.1rem 0;
            background:linear-gradient(135deg, rgba(15,23,42,.95), rgba(30,64,175,.55));
            border:1px solid rgba(96,165,250,.25);
            color:#f8fafc;
        }
        .aicfo-import-banner b {color:#fff;}
        .aicfo-small {color:#cbd5e1;font-size:.9rem;line-height:1.45;}
        .aicfo-dialog-note {color:#cbd5e1;font-size:.92rem;line-height:1.45;margin-bottom:.7rem;}
        @media(max-width: 950px){
            .aicfo-grid-2,.aicfo-grid-3,.aicfo-grid-4{grid-template-columns:1fr;}
            .aicfo-hero-title{font-size:2.15rem;}
            .aicfo-hero{padding:1.35rem;}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def v1_set_source_choice(choice: str):
    st.session_state["data_source_choice"] = choice
    if choice == "manual":
        st.query_params["page"] = "upload"
    st.rerun()


def _onboarding_content():
    st.markdown("### Choose your data import method")
    st.markdown(
        "<div class='aicfo-dialog-note'>ERP connection is the future preferred method. Manual upload remains available for clients who do not want to connect their accounting system yet.</div>",
        unsafe_allow_html=True,
    )
    left, right = st.columns(2)
    with left:
        st.markdown(
            """
            <div class='aicfo-card'>
                <span class='aicfo-tag blue'>Recommended · Coming Soon</span>
                <h3>🔌 Connect ERP</h3>
                <p>Planned connectors for Xero, MYOB, QuickBooks, Business Central, NetSuite, SAP and Oracle.</p>
                <ul>
                    <li>Automatic monthly refresh</li>
                    <li>Less Excel handling</li>
                    <li>Priority access for beta clients</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Request ERP Early Access", use_container_width=True, key="v1_erp_early_access_dialog"):
            st.session_state["data_source_choice"] = "erp_early_access"
            st.toast("ERP early access interest recorded.")
            st.rerun()
    with right:
        st.markdown(
            """
            <div class='aicfo-card'>
                <span class='aicfo-tag'>Available Now</span>
                <h3>📥 Manual Import</h3>
                <p>Upload Excel exports from any ERP or accounting software and let AI CFO Copilot validate and transform the data.</p>
                <ul>
                    <li>GL + COA mapping</li>
                    <li>Budgets, forecasts, AR/AP ageing</li>
                    <li>Validation before reports</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Continue with Manual Import", use_container_width=True, key="v1_manual_import_dialog"):
            v1_set_source_choice("manual")

    st.caption("You can change this later from the Home workspace.")


def render_v1_onboarding_dialog():
    inject_v1_polished_css()
    if st.session_state.get("data_source_choice"):
        return

    if hasattr(st, "dialog"):
        @st.dialog("Welcome to AI CFO Copilot")
        def dialog():
            _onboarding_content()
        dialog()
    else:
        with st.expander("Welcome to AI CFO Copilot", expanded=True):
            _onboarding_content()


def render_v1_home_intro(profile, score, profile_done, data_loaded, validation_ok, reports_ready, insights_ready):
    inject_v1_polished_css()
    company = profile.get("Company Name") or "Your Company"
    industry = profile.get("Industry") or "Industry not set"
    country = profile.get("Country") or "Country not set"
    period = get_report_period_label(profile)
    completed = sum([profile_done, data_loaded, validation_ok, reports_ready, insights_ready])
    progress = int(completed / 5 * 100)

    st.markdown(
        f"""
        <div class='aicfo-hero'>
            <div class='aicfo-kicker'>✨ AI-powered Finance Operating System</div>
            <div class='aicfo-hero-title'>From ERP data to board-ready decisions.</div>
            <div class='aicfo-hero-sub'>
                A guided workspace for month-end close, validation, financial statements, KPIs, forecasts, benchmarks and AI CFO commentary.
            </div>
            <div class='aicfo-grid-4'>
                <div class='aicfo-metric'><div class='aicfo-metric-label'>Company</div><div class='aicfo-metric-value'>{company}</div></div>
                <div class='aicfo-metric'><div class='aicfo-metric-label'>Industry</div><div class='aicfo-metric-value'>{industry}</div></div>
                <div class='aicfo-metric'><div class='aicfo-metric-label'>Period</div><div class='aicfo-metric-value'>{period}</div></div>
                <div class='aicfo-metric'><div class='aicfo-metric-label'>Readiness</div><div class='aicfo-metric-value'>{score}/100</div></div>
            </div>
            <div class='aicfo-progress-wrap'><div class='aicfo-progress-bar' style='width:{progress}%;'></div></div>
            <div class='aicfo-small'>Month-end close progress: <b>{progress}% complete</b> · {country}</div>
            <div class='aicfo-pill-row'>
                <span class='aicfo-pill'>Financial Statements</span>
                <span class='aicfo-pill'>KPI Packs</span>
                <span class='aicfo-pill'>Budget vs Actual</span>
                <span class='aicfo-pill'>AI CFO Brief</span>
                <span class='aicfo-pill'>Board Pack Ready</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    step_data = [
        ("1 Configure", "Company profile and reporting period", profile_done),
        ("2 Import", "ERP or manual finance data", data_loaded),
        ("3 Validate", "Mapping, duplicates and data quality", validation_ok),
        ("4 Reports", "P&L, Balance Sheet and KPIs", reports_ready),
        ("5 AI Review", "Executive commentary and actions", insights_ready),
    ]
    cards = []
    for idx, (title, sub, done) in enumerate(step_data):
        cls = "done" if done else ("active" if idx == completed else "")
        tick = "✓" if done else "○"
        cards.append(
            f"<div class='aicfo-step {cls}'><div class='aicfo-step-title'>{tick} {title}</div><div class='aicfo-step-sub'>{sub}</div></div>"
        )
    st.markdown("<div class='aicfo-grid-3'>" + "".join(cards) + "</div>", unsafe_allow_html=True)

    c1, c2 = st.columns([1.15, .85])
    with c1:
        render_v1_data_source_cards(compact=True)
    with c2:
        st.markdown(
            """
            <div class='aicfo-brief'>
                <span class='aicfo-tag blue'>AI CFO Brief</span>
                <h3>Today’s focus</h3>
                <p>• Complete company setup and import current period files.</p>
                <p>• Run validation before generating reports.</p>
                <p>• Use AI CFO to explain variances, mapping risks and management actions.</p>
                <p>• ERP connectors are displayed as upcoming to support client curiosity.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Open / Change Import Method", use_container_width=True, key="v1_change_import_method"):
            st.session_state["data_source_choice"] = None
            st.rerun()


def render_v1_data_source_cards(compact: bool = False):
    inject_v1_polished_css()
    if not compact:
        st.markdown(
            "<div class='aicfo-section-title'><div class='aicfo-section-icon'>🚀</div><div class='aicfo-section-text'>Get Started</div></div>",
            unsafe_allow_html=True,
        )
    choice = st.session_state.get("data_source_choice", "")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            """
            <div class='aicfo-card'>
                <span class='aicfo-tag blue'>Recommended · Coming Soon</span>
                <h3>🔌 Connect ERP</h3>
                <p>Show clients the direction: direct connections to Xero, MYOB, QuickBooks, Business Central, NetSuite, SAP and Oracle.</p>
                <p><b>V1 status:</b> early access / waitlist only.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Request ERP Early Access", use_container_width=True, key="v1_erp_home_card"):
            st.session_state["data_source_choice"] = "erp_early_access"
            st.toast("ERP early access interest recorded.")
    with c2:
        st.markdown(
            """
            <div class='aicfo-card'>
                <span class='aicfo-tag'>Available Now</span>
                <h3>📥 Manual Import</h3>
                <p>Upload GL, COA, budget, forecast, previous year, opening balance sheet and AR/AP ageing files exported from any ERP.</p>
                <p><b>V1 status:</b> live.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Open Import Centre", use_container_width=True, key="v1_manual_home_card"):
            v1_set_source_choice("manual")
    if choice:
        st.caption(f"Current data source choice: {choice.replace('_', ' ').title()}")


def render_v1_import_intro(profile):
    inject_v1_polished_css()
    company = profile.get("Company Name", "") or "Not set"
    industry = profile.get("Industry", "") or "Not set"
    country = profile.get("Country", "") or "Not set"
    reporting = profile.get("Reporting Structure", "Consolidated Only")
    st.markdown(
        "<div class='aicfo-section-title'><div class='aicfo-section-icon'>📥</div><div class='aicfo-section-text'>Import Centre</div></div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class='aicfo-import-banner'>
            <b>Importing for:</b> {company} &nbsp; | &nbsp;
            <b>Industry:</b> {industry} &nbsp; | &nbsp;
            <b>Country:</b> {country} &nbsp; | &nbsp;
            <b>Reporting:</b> {reporting}<br>
            <span class='aicfo-small'>Manual uploads are available now. ERP connection is shown on Home as coming soon / early access.</span>
        </div>
        """,
        unsafe_allow_html=True,
    )



def _fmt_money(value):
    try:
        value = float(value or 0)
        sign = "-" if value < 0 else ""
        value = abs(value)
        if value >= 1_000_000:
            return f"{sign}{value/1_000_000:,.2f}M"
        if value >= 1_000:
            return f"{sign}{value/1_000:,.2f}K"
        return f"{sign}{value:,.2f}"
    except Exception:
        return "0.00"


def _fmt_pct(value):
    try:
        return f"{float(value or 0):.2f}%"
    except Exception:
        return "0.00%"


def _get_kpi(kpi_df, name, default=0.0):
    try:
        if kpi_df is None or kpi_df.empty:
            return default
        row = kpi_df[kpi_df["KPI"].astype(str).str.lower() == str(name).lower()]
        if row.empty:
            return default
        return float(row.iloc[0]["Value"])
    except Exception:
        return default


def inject_v1_dashboard_css():
    inject_v1_polished_css()
    st.markdown(
        """
        <style>
        .dash-shell {
            border-radius: 32px;
            padding: 1.55rem;
            background:
              radial-gradient(circle at 16% 8%, rgba(37,99,235,.22), transparent 30%),
              radial-gradient(circle at 86% 16%, rgba(20,184,166,.20), transparent 32%),
              linear-gradient(135deg, #050816 0%, #0b1020 48%, #111827 100%);
            border: 1px solid rgba(148,163,184,.22);
            box-shadow: 0 28px 90px rgba(2,6,23,.38);
            margin-top: .6rem;
        }
        .dash-topbar {display:flex;justify-content:space-between;gap:1rem;align-items:flex-start;margin-bottom:1rem;}
        .dash-kicker {display:inline-flex;gap:.45rem;align-items:center;border-radius:999px;padding:.42rem .72rem;background:rgba(37,99,235,.17);border:1px solid rgba(96,165,250,.24);color:#bfdbfe;font-size:.78rem;font-weight:900;}
        .dash-title {color:#fff;font-size:2.35rem;line-height:1.04;font-weight:950;letter-spacing:-.055em;margin:.55rem 0 .3rem 0;}
        .dash-sub {color:#cbd5e1;font-size:.96rem;line-height:1.45;max-width:780px;}
        .dash-status-card {min-width:250px;border-radius:24px;padding:1rem;background:rgba(15,23,42,.76);border:1px solid rgba(148,163,184,.22);box-shadow:0 16px 42px rgba(0,0,0,.22);}
        .dash-status-label {color:#94a3b8;font-weight:850;font-size:.78rem;margin-bottom:.3rem;}
        .dash-score {color:#fff;font-size:2.05rem;font-weight:950;letter-spacing:-.05em;}
        .dash-progress {height:11px;background:rgba(148,163,184,.22);border-radius:999px;overflow:hidden;margin:.65rem 0 .35rem 0;}
        .dash-progress-bar {height:100%;background:linear-gradient(90deg,#14b8a6,#2563eb,#7c3aed);border-radius:999px;}
        .dash-grid-4 {display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.9rem;margin:.9rem 0;}
        .dash-grid-3 {display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.9rem;margin:.9rem 0;}
        .dash-grid-2 {display:grid;grid-template-columns:1.2fr .8fr;gap:.9rem;margin:.9rem 0;}
        .dash-kpi {position:relative;overflow:hidden;border-radius:24px;padding:1.05rem;background:linear-gradient(180deg,rgba(15,23,42,.92),rgba(17,24,39,.96));border:1px solid rgba(148,163,184,.20);box-shadow:0 18px 45px rgba(0,0,0,.25);min-height:138px;}
        .dash-kpi:after {content:"";position:absolute;right:-35px;top:-35px;width:110px;height:110px;border-radius:999px;background:rgba(37,99,235,.13);}
        .dash-kpi-label {color:#94a3b8;font-size:.78rem;font-weight:900;text-transform:uppercase;letter-spacing:.055em;}
        .dash-kpi-value {color:#fff;font-size:1.75rem;font-weight:950;letter-spacing:-.045em;margin:.55rem 0 .35rem 0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        .dash-kpi-note {color:#cbd5e1;font-size:.83rem;line-height:1.35;}
        .dash-kpi-chip {display:inline-flex;margin-top:.55rem;padding:.28rem .55rem;border-radius:999px;font-size:.76rem;font-weight:850;background:rgba(34,197,94,.12);border:1px solid rgba(34,197,94,.25);color:#86efac;}
        .dash-panel {border-radius:26px;padding:1.2rem;background:linear-gradient(180deg,rgba(15,23,42,.88),rgba(17,24,39,.94));border:1px solid rgba(148,163,184,.22);box-shadow:0 20px 50px rgba(0,0,0,.23);}
        .dash-panel h3 {color:#fff;font-size:1.2rem;margin:.1rem 0 .6rem 0;letter-spacing:-.025em;}
        .dash-panel p, .dash-panel li {color:#cbd5e1;font-size:.92rem;line-height:1.45;}
        .dash-ai {background:linear-gradient(135deg,rgba(37,99,235,.24),rgba(15,23,42,.96));border-color:rgba(96,165,250,.30);}
        .dash-step-row {display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:.7rem;margin-top:.75rem;}
        .dash-step {border-radius:18px;padding:.8rem;background:rgba(15,23,42,.75);border:1px solid rgba(148,163,184,.22);min-height:86px;}
        .dash-step.done {background:rgba(6,78,59,.48);border-color:rgba(52,211,153,.44);}
        .dash-step.active {background:rgba(30,64,175,.50);border-color:rgba(96,165,250,.52);}
        .dash-step b {color:#fff;font-size:.86rem;display:block;margin-bottom:.2rem;}
        .dash-step span {color:#cbd5e1;font-size:.76rem;line-height:1.25;}
        .dash-action-card {border-radius:22px;padding:1rem;background:rgba(15,23,42,.82);border:1px solid rgba(148,163,184,.22);min-height:116px;}
        .dash-action-title {color:#fff;font-weight:920;margin-bottom:.25rem;}
        .dash-action-sub {color:#cbd5e1;font-size:.82rem;line-height:1.3;}
        .dash-activity {display:flex;gap:.72rem;align-items:flex-start;border-bottom:1px solid rgba(148,163,184,.14);padding:.65rem 0;}
        .dash-dot {width:10px;height:10px;border-radius:99px;background:#22c55e;margin-top:.38rem;box-shadow:0 0 0 4px rgba(34,197,94,.14);}
        .dash-activity b {color:#fff;font-size:.88rem;}
        .dash-activity span {color:#94a3b8;font-size:.78rem;display:block;margin-top:.05rem;}
        .dash-muted {color:#94a3b8;font-size:.82rem;}
        .dash-empty {border-radius:26px;padding:1.2rem;background:rgba(15,23,42,.72);border:1px dashed rgba(148,163,184,.35);color:#cbd5e1;}
        @media(max-width: 1050px){.dash-grid-4,.dash-grid-3,.dash-grid-2,.dash-step-row{grid-template-columns:1fr}.dash-topbar{display:block}.dash-status-card{margin-top:1rem;min-width:0}.dash-title{font-size:2rem}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_v1_dashboard_command_center(profile, readiness_score, profile_done, data_loaded, validation_ok, reports_ready, insights_ready, session_state):
    """Premium dashboard-first command centre. Presentation only; no finance calculations changed."""
    inject_v1_dashboard_css()
    company = profile.get("Company Name") or "Create your workspace"
    industry = profile.get("Industry") or "Industry pending"
    period = get_report_period_label(profile)
    completed = sum([profile_done, data_loaded, validation_ok, reports_ready, insights_ready])
    progress = int(completed / 5 * 100)

    kpis = session_state.get("consolidated_kpis")
    ar = session_state.get("ar_summary") or {}
    ap = session_state.get("ap_summary") or {}
    report = session_state.get("last_validation_report") or {}
    critical = len(report.get("critical", []) or [])
    warnings = len(report.get("warnings", []) or [])

    revenue = _get_kpi(kpis, "Revenue")
    gross_profit = _get_kpi(kpis, "Gross Profit")
    gross_margin = _get_kpi(kpis, "Gross Margin %")
    operating_profit = _get_kpi(kpis, "Operating Profit")
    operating_margin = _get_kpi(kpis, "Operating Margin %")
    ar_total = ar.get("total", 0) if isinstance(ar, dict) else 0
    ap_total = ap.get("total", 0) if isinstance(ap, dict) else 0

    st.markdown(
        f"""
        <div class='dash-shell'>
            <div class='dash-topbar'>
                <div>
                    <div class='dash-kicker'>● Finance command centre</div>
                    <div class='dash-title'>Good morning. Your {period} close is {progress}% complete.</div>
                    <div class='dash-sub'>
                        Workspace: <b>{company}</b> · {industry}. Start with import and validation, then move into reports, AI review and board-ready packs.
                    </div>
                </div>
                <div class='dash-status-card'>
                    <div class='dash-status-label'>DATA READINESS</div>
                    <div class='dash-score'>{readiness_score}/100</div>
                    <div class='dash-progress'><div class='dash-progress-bar' style='width:{readiness_score}%;'></div></div>
                    <div class='dash-muted'>{critical} critical · {warnings} warning(s)</div>
                </div>
            </div>
            <div class='dash-grid-4'>
                <div class='dash-kpi'><div class='dash-kpi-label'>Revenue</div><div class='dash-kpi-value'>{_fmt_money(revenue)}</div><div class='dash-kpi-note'>Current period sales performance.</div><span class='dash-kpi-chip'>Statement-linked</span></div>
                <div class='dash-kpi'><div class='dash-kpi-label'>Gross Profit</div><div class='dash-kpi-value'>{_fmt_money(gross_profit)}</div><div class='dash-kpi-note'>After direct costs / COGS.</div><span class='dash-kpi-chip'>{_fmt_pct(gross_margin)} margin</span></div>
                <div class='dash-kpi'><div class='dash-kpi-label'>Operating Profit</div><div class='dash-kpi-value'>{_fmt_money(operating_profit)}</div><div class='dash-kpi-note'>After overheads and operating expenses.</div><span class='dash-kpi-chip'>{_fmt_pct(operating_margin)} margin</span></div>
                <div class='dash-kpi'><div class='dash-kpi-label'>Working Capital</div><div class='dash-kpi-value'>{_fmt_money(ar_total - ap_total)}</div><div class='dash-kpi-note'>AR less AP uploaded ageing balance.</div><span class='dash-kpi-chip'>AR {_fmt_money(ar_total)} · AP {_fmt_money(ap_total)}</span></div>
            </div>
            <div class='dash-panel'>
                <h3>Month-end close progress</h3>
                <div class='dash-progress'><div class='dash-progress-bar' style='width:{progress}%;'></div></div>
                <div class='dash-step-row'>
                    <div class='dash-step {'done' if profile_done else 'active'}'><b>{'✓' if profile_done else '○'} Company</b><span>Profile, period and reporting structure.</span></div>
                    <div class='dash-step {'done' if data_loaded else ('active' if profile_done else '')}'><b>{'✓' if data_loaded else '○'} Import</b><span>GL, COA and optional packs.</span></div>
                    <div class='dash-step {'done' if validation_ok else ('active' if data_loaded else '')}'><b>{'✓' if validation_ok else '○'} Validate</b><span>Mapping, duplicates and date checks.</span></div>
                    <div class='dash-step {'done' if reports_ready else ('active' if validation_ok else '')}'><b>{'✓' if reports_ready else '○'} Reports</b><span>P&L, BS, KPIs and variances.</span></div>
                    <div class='dash-step {'done' if insights_ready else ('active' if reports_ready else '')}'><b>{'✓' if insights_ready else '○'} AI Review</b><span>Executive narrative and actions.</span></div>
                    <div class='dash-step'><b>○ Board Pack</b><span>V1 preview / next release export flow.</span></div>
                </div>
            </div>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns([1.18, .82])
    with left:
        st.markdown(
            """
            <div class='dash-panel dash-ai'>
                <div class='dash-kicker'>🤖 AI CFO Brief</div>
                <h3>Today’s recommended focus</h3>
                <p>• Finish import and validation before relying on the financial pack.</p>
                <p>• Review revenue, gross margin and working capital movements first.</p>
                <p>• Use AI CFO to explain variances, benchmark gaps and management actions.</p>
                <p>• ERP connectors are visible as roadmap items to support client curiosity.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        q1, q2, q3 = st.columns(3)
        if q1.button("📥 Import Files", use_container_width=True, key="dash_go_import"):
            st.query_params["page"] = "upload"
            st.rerun()
        if q2.button("📈 View Reports", use_container_width=True, key="dash_go_reports"):
            st.query_params["page"] = "reports"
            st.rerun()
        if q3.button("🤖 Open AI CFO", use_container_width=True, key="dash_open_ai"):
            st.session_state["ai_cfo_panel_open"] = True
            st.rerun()

        st.markdown(
            """
            <div class='dash-grid-3'>
                <div class='dash-action-card'><div class='dash-action-title'>Generate board pack</div><div class='dash-action-sub'>Coming next: one-click PDF / PowerPoint pack with AI narrative.</div></div>
                <div class='dash-action-card'><div class='dash-action-title'>Forecast review</div><div class='dash-action-sub'>Compare uploaded forecast P&L / BS against actuals.</div></div>
                <div class='dash-action-card'><div class='dash-action-title'>Tax centre preview</div><div class='dash-action-sub'>GST/BAS/VAT estimate module planned after reporting V1.</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with right:
        st.markdown("<div class='dash-panel'><h3>Recent activity</h3>", unsafe_allow_html=True)
        activity = []
        if profile_done:
            activity.append(("Workspace configured", "Company profile and reporting period saved."))
        if data_loaded:
            activity.append(("Files imported", "GL and COA have been processed."))
        if validation_ok:
            activity.append(("Validation passed", "Reports are ready to use."))
        if reports_ready:
            activity.append(("Reports generated", "P&L, BS and KPI pack available."))
        if insights_ready:
            activity.append(("AI review generated", "Executive commentary available."))
        if not activity:
            activity = [("Start your workspace", "Choose ERP early access or manual import."), ("Upload GL + COA", "Use templates in the Import Centre."), ("Validate data", "Review mapping and duplicate checks.")]
        for title, sub in activity[-5:][::-1]:
            st.markdown(f"<div class='dash-activity'><div class='dash-dot'></div><div><b>{title}</b><span>{sub}</span></div></div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    if not data_loaded:
        st.markdown(
            """
            <div class='dash-empty'>
                <b>Next step:</b> open the Import Centre and upload GL + COA. ERP connectors are shown as coming soon, but manual import is live for V1 demos.
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("</div>", unsafe_allow_html=True)
