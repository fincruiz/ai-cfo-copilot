from __future__ import annotations

from datetime import date
from typing import Any

import streamlit as st

from services.demo_data import load_demo_workspace


INDUSTRIES = [
    "Manufacturing", "Wholesale / Distribution", "Retail", "Professional Services",
    "Construction", "Logistics", "Hospitality", "Healthcare", "Technology", "Other",
]
COUNTRIES = ["Australia", "India", "United States", "United Kingdom", "Canada", "New Zealand", "Other"]
CURRENCIES = ["AUD", "INR", "USD", "GBP", "CAD", "NZD", "EUR", "Other"]
DEFAULT_MODULES = [
    "Financial Statements", "KPI Dashboard", "AI CFO", "Forecasting",
    "Working Capital", "Benchmarking", "Board Pack",
]


def _initialise_auth_state() -> None:
    defaults: dict[str, Any] = {
        "app_logged_in": False,
        "auth_mode": None,
        "auth_view": "login",
        "onboarding_step": 1,
        "workspace_modules": DEFAULT_MODULES.copy(),
        "company_profile": {},
        "onboarding_draft": {},
        "demo_choice": None,
        "demo_intro_pending": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state or st.session_state[key] is None:
            st.session_state[key] = value.copy() if isinstance(value, (dict, list)) else value


def _inject_login_css() -> None:
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"], [data-testid="collapsedControl"] {display:none !important;}
        [data-testid="stHeader"] {background:transparent !important;}
        [data-testid="stToolbar"] {right: .5rem !important;}
        html, body, .stApp {height:100%; overflow-x:hidden;}
        .stApp {
          min-height:100vh;
          background:
            radial-gradient(circle at 10% 20%, rgba(37,99,235,.24), transparent 28%),
            radial-gradient(circle at 92% 14%, rgba(124,58,237,.22), transparent 26%),
            radial-gradient(circle at 70% 90%, rgba(8,145,178,.15), transparent 30%),
            linear-gradient(135deg,#050914 0%,#081120 48%,#070913 100%) !important;
        }
        .block-container {
          max-width:1440px !important;
          padding:.6rem 1rem !important;
          min-height:calc(100vh - 1.2rem);
          display:flex;
          align-items:center;
        }
        .st-key-login_shell {width:100%; animation:loginEnter .55s cubic-bezier(.2,.8,.2,1) both;}
        .enterprise-login {
            width:min(100%,1480px);margin:0 auto;
          width:100%; min-height:min(820px, calc(100vh - 2rem));
          display:grid; grid-template-columns:minmax(0,1.18fr) minmax(390px,.82fr);
          border-radius:32px; overflow:hidden;
          border:1px solid rgba(148,163,184,.18);
          background:rgba(5,10,20,.90);
          box-shadow:0 38px 120px rgba(0,0,0,.58);
          position:relative;
        }
        .enterprise-login:before {
          content:""; position:absolute; inset:0; pointer-events:none; z-index:4;
          background:linear-gradient(115deg,transparent 0%,rgba(255,255,255,.025) 45%,transparent 58%);
          transform:translateX(-110%); animation:sheen 8s ease-in-out infinite;
        }
        .login-visual {overflow:hidden;isolation:isolate;
          min-height:100%; position:relative; overflow:hidden; padding:2.2rem 2.25rem;
          display:flex; flex-direction:column; justify-content:space-between;
          background:
            linear-gradient(115deg,rgba(3,10,25,.94) 0%,rgba(9,20,42,.78) 48%,rgba(15,23,42,.43) 100%),
            url('https://images.unsplash.com/photo-1554224155-6726b3ff858f?auto=format&fit=crop&w=1800&q=88');
          background-size:cover; background-position:center;
        }
        .login-visual:after {
          content:""; position:absolute; inset:0; pointer-events:none;
          background-image:linear-gradient(rgba(148,163,184,.045) 1px,transparent 1px),linear-gradient(90deg,rgba(148,163,184,.045) 1px,transparent 1px);
          background-size:44px 44px; mask-image:linear-gradient(to bottom,black,transparent 90%);
          animation:gridDrift 18s linear infinite;
        }
        .login-logo {position:relative;z-index:2;display:flex;align-items:center;gap:.78rem;color:#fff;font-weight:950;font-size:1.15rem;}
        .login-logo-mark {width:46px;height:46px;border-radius:15px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#06b6d4,#2563eb,#7c3aed);box-shadow:0 15px 38px rgba(37,99,235,.42);animation:logoPulse 3.4s ease-in-out infinite;}
        .login-kicker {position:relative;z-index:2;display:inline-flex;width:max-content;padding:.43rem .75rem;border-radius:999px;color:#dbeafe;background:rgba(37,99,235,.20);border:1px solid rgba(147,197,253,.30);font-size:.74rem;font-weight:900;letter-spacing:.08em;text-transform:uppercase;backdrop-filter:blur(10px);}
        .login-headline {position:relative;z-index:2;font-size:clamp(2.65rem,4.2vw,4.7rem);line-height:.98;letter-spacing:-.065em;font-weight:950;color:#fff;margin:.75rem 0 .85rem;max-width:760px;text-shadow:0 12px 40px rgba(0,0,0,.35);}
        .login-copy {position:relative;z-index:2;font-size:1.02rem;line-height:1.65;color:#dbeafe;max-width:650px;}
        .floating-finance-card {display:none !important;}
        .ffc-1{right:7%;top:19%;}.ffc-2{right:13%;bottom:17%;animation-delay:-2.2s}.ffc-3{left:8%;bottom:10%;animation-delay:-3.6s}
        .ffc-label{font-size:.67rem;color:#93c5fd;text-transform:uppercase;letter-spacing:.08em;font-weight:900}.ffc-value{font-size:1.24rem;color:#fff;font-weight:950;margin-top:.15rem}.ffc-trend{font-size:.72rem;color:#86efac;font-weight:800;margin-top:.12rem}
        .login-feature-row {position:relative;z-index:2;display:flex;flex-wrap:wrap;gap:.55rem;margin-top:1.05rem;}
        .login-feature {padding:.52rem .72rem;border-radius:999px;border:1px solid rgba(191,219,254,.20);background:rgba(5,12,24,.55);color:#e2e8f0;font-weight:800;font-size:.78rem;backdrop-filter:blur(10px);transition:transform .2s ease,border-color .2s ease;}
        .login-feature:hover{transform:translateY(-2px);border-color:rgba(96,165,250,.65)}
        .login-proof {position:relative;z-index:2;display:grid;grid-template-columns:repeat(3,1fr);gap:.7rem;max-width:650px;}
        .proof-item {padding:.78rem;border-radius:16px;border:1px solid rgba(148,163,184,.16);background:rgba(2,6,23,.48);backdrop-filter:blur(12px);}
        .proof-value {font-size:1.22rem;color:#fff;font-weight:950}.proof-label {font-size:.68rem;color:#94a3b8;font-weight:800;margin-top:.1rem}
        .login-actions {min-height:100%;padding:2.05rem 2rem;display:flex;flex-direction:column;justify-content:center;background:linear-gradient(160deg,rgba(7,12,23,.97),rgba(10,15,29,.99));position:relative;z-index:5;}
        .login-actions:before {content:"";position:absolute;width:260px;height:260px;right:-120px;top:-130px;border-radius:50%;background:rgba(124,58,237,.14);filter:blur(45px);animation:orbFloat 7s ease-in-out infinite alternate;}
        .login-card-title {position:relative;font-size:2.15rem;font-weight:950;letter-spacing:-.05em;color:#fff;margin-bottom:.35rem;}
        .login-card-copy {position:relative;color:#94a3b8;font-size:.9rem;line-height:1.5;margin-bottom:1rem;}
        .login-primary-box {position:relative;padding:1rem 1.05rem;border-radius:18px;background:linear-gradient(135deg,rgba(37,99,235,.20),rgba(124,58,237,.18));border:1px solid rgba(129,140,248,.34);margin:.25rem 0 .7rem;box-shadow:0 14px 36px rgba(37,99,235,.10);}
        .login-primary-box b{color:#fff}.login-primary-box span{display:block;color:#c7d2fe;font-size:.8rem;margin-top:.2rem;line-height:1.4}
        .login-divider {display:flex;align-items:center;gap:.65rem;color:#64748b;font-size:.7rem;font-weight:850;margin:.65rem 0;}
        .login-divider:before,.login-divider:after{content:"";height:1px;flex:1;background:rgba(148,163,184,.18)}
        .login-trust {display:flex;align-items:center;gap:.55rem;color:#64748b;font-size:.71rem;margin-top:.6rem}.trust-dot{width:7px;height:7px;border-radius:50%;background:#22c55e;box-shadow:0 0 12px rgba(34,197,94,.65)}
        .st-key-login_actions div[data-testid="stButton"] > button {border-radius:14px !important;min-height:2.8rem !important;border:1px solid rgba(148,163,184,.25) !important;background:rgba(15,23,42,.94) !important;color:#f8fafc !important;font-weight:850 !important;transition:all .2s ease !important;}
        .st-key-login_actions div[data-testid="stButton"] > button:hover {transform:translateY(-2px) scale(1.005);border-color:#60a5fa !important;background:linear-gradient(135deg,#1d4ed8,#4f46e5) !important;box-shadow:0 15px 38px rgba(37,99,235,.28) !important;}
        .st-key-create_workspace button {background:linear-gradient(135deg,#0284c7,#2563eb,#7c3aed) !important;color:#fff !important;box-shadow:0 15px 38px rgba(79,70,229,.30) !important;border-color:rgba(191,219,254,.42) !important;}
        .st-key-open_demo_choice button {background:linear-gradient(135deg,rgba(8,145,178,.18),rgba(79,70,229,.18)) !important;border-color:rgba(34,211,238,.35) !important;}

        div[data-testid="stDialog"] {background:rgba(2,6,23,.66) !important;backdrop-filter:blur(10px);}
        div[data-testid="stDialog"] > div {
          width:min(900px,calc(100vw - 2rem)) !important; max-width:900px !important;
          border-radius:28px !important;border:1px solid rgba(96,165,250,.25) !important;
          background:linear-gradient(160deg,#090f1c,#111827) !important;
          box-shadow:0 45px 140px rgba(0,0,0,.68) !important;
          animation:dialogEnter .32s cubic-bezier(.2,.8,.2,1) both;
        }
        div[data-testid="stDialog"] h1,div[data-testid="stDialog"] h2,div[data-testid="stDialog"] h3,div[data-testid="stDialog"] p,div[data-testid="stDialog"] label{color:#f8fafc !important;}
        div[data-testid="stDialog"] input,div[data-testid="stDialog"] textarea{background:#0f172a !important;color:#fff !important;border-radius:13px !important;}
        div[data-testid="stDialog"] [data-baseweb="select"] > div{background:#0f172a !important;color:#fff !important;border-radius:13px !important;}
        .wizard-progress{display:flex;gap:.42rem;margin:.15rem 0 1rem}.wizard-progress span{height:6px;flex:1;border-radius:999px;background:#253147}.wizard-progress span.done{background:linear-gradient(90deg,#06b6d4,#2563eb,#7c3aed);box-shadow:0 0 18px rgba(59,130,246,.40);animation:progressGlow 2s ease-in-out infinite alternate}
        .wizard-label{font-size:.75rem;color:#93c5fd;font-weight:900;text-transform:uppercase;letter-spacing:.1em;margin-bottom:.2rem}.wizard-summary{padding:1rem;border-radius:16px;background:rgba(34,197,94,.08);border:1px solid rgba(34,197,94,.22);color:#d1fae5}
        .demo-choice-grid {display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin:.7rem 0 1rem;}
        .demo-choice-card {padding:1.15rem;border-radius:19px;border:1px solid rgba(148,163,184,.20);background:linear-gradient(145deg,rgba(15,23,42,.92),rgba(17,24,39,.76));min-height:165px;transition:transform .2s ease,border-color .2s ease,box-shadow .2s ease;}
        .demo-choice-card:hover{transform:translateY(-4px);border-color:rgba(96,165,250,.55);box-shadow:0 20px 50px rgba(37,99,235,.15)}
        .demo-choice-icon{font-size:1.65rem}.demo-choice-title{color:#fff;font-weight:950;font-size:1.06rem;margin:.45rem 0 .28rem}.demo-choice-copy{color:#cbd5e1;font-size:.82rem;line-height:1.48}.demo-choice-time{color:#67e8f9;font-size:.72rem;font-weight:850;margin-top:.65rem}
        @keyframes loginEnter{from{opacity:0;transform:translateY(14px) scale(.988)}to{opacity:1;transform:none}}
        @keyframes dialogEnter{from{opacity:0;transform:translateY(14px) scale(.965)}to{opacity:1;transform:none}}
        @keyframes sheen{0%,72%{transform:translateX(-110%)}100%{transform:translateX(110%)}}
        @keyframes gridDrift{from{background-position:0 0,0 0}to{background-position:44px 44px,44px 44px}}
        @keyframes logoPulse{0%,100%{transform:scale(1);box-shadow:0 15px 38px rgba(37,99,235,.42)}50%{transform:scale(1.045);box-shadow:0 20px 50px rgba(124,58,237,.48)}}
        @keyframes cardFloat{0%,100%{transform:translateY(0) rotate(0)}50%{transform:translateY(-12px) rotate(.5deg)}}
        @keyframes orbFloat{from{transform:translate(0,0)}to{transform:translate(-35px,28px)}}
        @keyframes progressGlow{from{filter:brightness(.9)}to{filter:brightness(1.25)}}
        @media(max-width:980px){
          .block-container{display:block;padding:.45rem !important}.enterprise-login{grid-template-columns:1fr;min-height:auto}.login-visual{min-height:430px;padding:1.35rem}.login-actions{padding:1.3rem;min-height:auto}.floating-finance-card{display:none}.login-proof{display:none}.login-headline{font-size:2.55rem}.login-feature-row{margin-bottom:.4rem}
        }
        @media(prefers-reduced-motion:reduce){*,*:before,*:after{animation-duration:.01ms !important;animation-iteration-count:1 !important;transition-duration:.01ms !important}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _set_logged_in(mode: str) -> None:
    st.session_state["app_logged_in"] = True
    st.session_state["auth_mode"] = mode
    st.session_state["auth_view"] = "login"
    st.query_params["page"] = "dashboard"


@st.dialog("Sign in to AI CFO Copilot")
def _email_sign_in_dialog() -> None:
    st.markdown('<div class="wizard-label">Secure workspace access</div>', unsafe_allow_html=True)
    st.caption("V1 preview authentication. Production identity will be connected before public release.")
    left, right = st.columns([1.15, .85], gap="large")
    with left:
        with st.form("email_sign_in_form"):
            email = st.text_input("Work email", placeholder="name@company.com")
            password = st.text_input("Password", type="password", placeholder="••••••••")
            remember = st.checkbox("Remember this workspace")
            submitted = st.form_submit_button("Sign in securely", use_container_width=True)
        if submitted:
            if not email.strip() or not password:
                st.error("Enter your email and password.")
            else:
                st.session_state["login_email"] = email.strip()
                st.session_state["remember_workspace"] = remember
                _set_logged_in("email_preview")
                st.rerun()
    with right:
        st.markdown(
            """
            <div style="padding:1rem;border-radius:18px;background:linear-gradient(145deg,rgba(37,99,235,.15),rgba(124,58,237,.12));border:1px solid rgba(129,140,248,.28);height:100%">
              <div style="color:#fff;font-weight:950;font-size:1.05rem">Your finance workspace</div>
              <div style="color:#cbd5e1;font-size:.82rem;line-height:1.5;margin-top:.4rem">Access company reporting, validation, forecasts, working capital and AI CFO analysis from one secure workspace.</div>
              <div style="margin-top:.9rem;color:#93c5fd;font-size:.76rem;font-weight:850">V1 PREVIEW</div>
              <div style="color:#94a3b8;font-size:.74rem;line-height:1.45;margin-top:.25rem">Identity providers, password reset and multi-factor authentication are scheduled before public release.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


@st.dialog("Choose your demo experience")
def _demo_choice_dialog() -> None:
    st.caption("Both options use the same complete sample company. The difference is how the experience is delivered.")
    st.markdown(
        """
        <div class="demo-choice-grid">
          <div class="demo-choice-card"><div class="demo-choice-icon">🎬</div><div class="demo-choice-title">Interactive Product Tour</div><div class="demo-choice-copy">A coached six-step journey. The application moves you through Dashboard, Import Centre, Reports, Working Capital, AI Insights and Downloads with a persistent guide.</div><div class="demo-choice-time">6–8 minutes · recommended for first visit</div></div>
          <div class="demo-choice-card"><div class="demo-choice-icon">🚀</div><div class="demo-choice-title">Explore Sample Workspace</div><div class="demo-choice-copy">Jump into a fully populated sample company with no prompts or forced navigation. Browse any page, open the AI CFO and inspect the reports at your own pace.</div><div class="demo-choice-time">Open-ended · complete freedom</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns(2)
    if c1.button("Start interactive tour →", type="primary", use_container_width=True, key="choose_guided_demo"):
        load_demo_workspace(force=True)
        st.session_state["demo_tour_active"] = True
        st.session_state["demo_tour_step"] = 0
        st.session_state["demo_intro_pending"] = True
        st.query_params["page"] = "dashboard"
        st.rerun()
    if c2.button("Explore freely →", use_container_width=True, key="choose_free_demo"):
        load_demo_workspace(force=True)
        st.session_state["demo_tour_active"] = False
        st.session_state["demo_tour_step"] = 0
        st.session_state["demo_intro_pending"] = False
        st.query_params["page"] = "dashboard"
        st.rerun()


@st.dialog("Create your finance workspace")
def _workspace_wizard_dialog() -> None:
    step = int(st.session_state.get("onboarding_step", 1))
    draft = st.session_state.get("onboarding_draft", {}) or {}
    bars = "".join('<span class="done"></span>' if idx <= step else '<span></span>' for idx in range(1, 6))
    st.markdown(f'<div class="wizard-progress">{bars}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="wizard-label">Step {step} of 5</div>', unsafe_allow_html=True)

    if step == 1:
        st.subheader("Tell us about the company")
        with st.form("workspace_company_form"):
            company = st.text_input("Company name *", value=draft.get("Company Name", ""))
            c1, c2 = st.columns(2)
            industry = c1.selectbox("Industry", INDUSTRIES, index=INDUSTRIES.index(draft.get("Industry")) if draft.get("Industry") in INDUSTRIES else 0)
            country = c2.selectbox("Country", COUNTRIES, index=COUNTRIES.index(draft.get("Country")) if draft.get("Country") in COUNTRIES else 0)
            currency = c1.selectbox("Currency", CURRENCIES, index=CURRENCIES.index(draft.get("Currency")) if draft.get("Currency") in CURRENCIES else 0)
            tax_id = c2.text_input("Tax ID / ABN / GSTIN (optional)", value=draft.get("Tax Identifier", ""))
            nxt = st.form_submit_button("Continue →", use_container_width=True)
        if nxt:
            if not company.strip():
                st.error("Company name is required.")
            else:
                draft.update({"Company Name": company.strip(), "Industry": industry, "Country": country, "Currency": currency, "Tax Identifier": tax_id})
                st.session_state["onboarding_draft"] = draft
                st.session_state["onboarding_step"] = 2
                st.rerun()
    elif step == 2:
        st.subheader("Choose your data source")
        with st.form("workspace_source_form"):
            source = st.radio("Import method", ["Connect ERP — Early access", "Upload Excel / CSV — Available now"], index=0 if draft.get("Data Source", "Connect").startswith("Connect") else 1)
            erp = st.selectbox("Preferred ERP", ["Microsoft Business Central", "Xero", "MYOB", "QuickBooks", "NetSuite", "SAP", "Oracle"]) if source.startswith("Connect") else "Manual Import"
            st.info("ERP preference is recorded for early access. Manual import remains available until the connector launches.")
            back, nxt = st.columns(2)
            b = back.form_submit_button("← Back", use_container_width=True)
            n = nxt.form_submit_button("Continue →", use_container_width=True)
        if b:
            st.session_state["onboarding_step"] = 1
            st.rerun()
        if n:
            draft.update({"Data Source": source, "Preferred ERP": erp})
            st.session_state["onboarding_draft"] = draft
            st.session_state["onboarding_step"] = 3
            st.rerun()
    elif step == 3:
        st.subheader("Set the reporting period")
        with st.form("workspace_period_form"):
            c1, c2 = st.columns(2)
            freq = c1.selectbox("Reporting frequency", ["Monthly", "Quarterly", "Annual"])
            period = c2.text_input("First report period *", value=draft.get("Report Period", ""), placeholder="Example: June 2026")
            start = c1.date_input("Period start", value=date.today().replace(day=1))
            end = c2.date_input("Period end", value=date.today())
            structure = st.radio("Reporting structure", ["Consolidated Only", "Branch / Business Unit Reporting"], horizontal=True)
            back, nxt = st.columns(2)
            b = back.form_submit_button("← Back", use_container_width=True)
            n = nxt.form_submit_button("Continue →", use_container_width=True)
        if b:
            st.session_state["onboarding_step"] = 2
            st.rerun()
        if n:
            if not period.strip():
                st.error("Report period is required.")
            elif start > end:
                st.error("Period start cannot be after period end.")
            else:
                draft.update({"Reporting Period": freq, "Report Period": period.strip(), "Period Start Date": str(start), "Period End Date": str(end), "Reporting Structure": structure})
                st.session_state["onboarding_draft"] = draft
                st.session_state["onboarding_step"] = 4
                st.rerun()
    elif step == 4:
        st.subheader("Choose workspace modules")
        with st.form("workspace_modules_form"):
            modules = st.multiselect("Enabled modules", DEFAULT_MODULES + ["Tax Centre Preview"], default=st.session_state.get("workspace_modules") or DEFAULT_MODULES)
            back, nxt = st.columns(2)
            b = back.form_submit_button("← Back", use_container_width=True)
            n = nxt.form_submit_button("Review workspace →", use_container_width=True)
        if b:
            st.session_state["onboarding_step"] = 3
            st.rerun()
        if n:
            st.session_state["workspace_modules"] = modules
            st.session_state["onboarding_step"] = 5
            st.rerun()
    else:
        st.subheader("Your workspace is ready")
        st.markdown(f'<div class="wizard-summary"><b>{draft.get("Company Name", "Your company")}</b><br>{draft.get("Country", "")} · {draft.get("Industry", "")} · {draft.get("Report Period", "")}</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        if c1.button("← Back", use_container_width=True, key="wizard_back_final"):
            st.session_state["onboarding_step"] = 4
            st.rerun()
        if c2.button("Launch dashboard →", type="primary", use_container_width=True, key="wizard_launch"):
            st.session_state["company_profile"] = draft.copy()
            st.session_state["reporting_structure"] = draft.get("Reporting Structure", "Consolidated Only")
            _set_logged_in("new_workspace")
            st.rerun()


def _render_login_screen() -> None:
    _inject_login_css()
    with st.container(key="login_shell"):
        left, right = st.columns([1.18, .82], gap=None)
        # Build a single visual shell while keeping Streamlit buttons functional in the right column.
        st.markdown('<div style="display:none">enterprise shell marker</div>', unsafe_allow_html=True)
        with left:
            st.markdown(
                """
                <div class="login-visual">
                  <div class="login-logo"><div class="login-logo-mark">▣</div><div>AI CFO Copilot</div></div>
                  <div>
                    <div class="login-kicker">AI-powered Finance Operating System</div>
                    <div class="login-headline">Finance clarity.<br>Board-ready speed.</div>
                    <div class="login-copy">Move from ERP entries to validated reports, forecasts, working-capital action and AI CFO commentary in one controlled workspace.</div>
                    <div class="login-feature-row">
                      <span class="login-feature">Financial statements</span><span class="login-feature">Budget vs actual</span><span class="login-feature">Three-way forecast</span><span class="login-feature">AI decision support</span>
                    </div>
                  </div>
                  <div class="login-proof">
                    <div class="proof-item"><div class="proof-value">Minutes</div><div class="proof-label">from upload to insight</div></div>
                    <div class="proof-item"><div class="proof-value">1 workspace</div><div class="proof-label">close, plan and explain</div></div>
                    <div class="proof-item"><div class="proof-value">ERP-first</div><div class="proof-label">manual import ready</div></div>
                  </div>
                  <div class="floating-finance-card ffc-1"><div class="ffc-label">Revenue</div><div class="ffc-value">$12.8M</div><div class="ffc-trend">▲ 8.4% vs plan</div></div>
                  <div class="floating-finance-card ffc-2"><div class="ffc-label">Close readiness</div><div class="ffc-value">92%</div><div class="ffc-trend">2 tasks remaining</div></div>
                  <div class="floating-finance-card ffc-3"><div class="ffc-label">AI CFO</div><div class="ffc-value">3 actions</div><div class="ffc-trend">Margin risk detected</div></div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with right:
            with st.container(key="login_actions"):
                st.markdown(
                    """
                    <div class="login-actions">
                      <div class="login-card-title">Welcome to your finance command centre</div>
                      <div class="login-card-copy">Create a workspace, sign in securely, or experience a fully populated sample company.</div>
                      <div class="login-primary-box"><b>New to AI CFO Copilot?</b><span>Company, reporting period, data source and modules are configured in one guided enterprise setup.</span></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("Create finance workspace →", type="primary", use_container_width=True, key="create_workspace"):
                    st.session_state["auth_view"] = "onboarding"
                    st.session_state["onboarding_step"] = 1
                    st.rerun()
                st.markdown('<div class="login-divider">already have a workspace?</div>', unsafe_allow_html=True)
                if st.button("Sign in with work email", use_container_width=True, key="email_signin"):
                    st.session_state["auth_view"] = "email"
                    st.rerun()
                c1, c2 = st.columns(2)
                if c1.button("Microsoft", use_container_width=True, key="ms_signin"):
                    _set_logged_in("microsoft_preview")
                    st.rerun()
                if c2.button("Google", use_container_width=True, key="google_signin"):
                    _set_logged_in("google_preview")
                    st.rerun()
                st.markdown('<div class="login-divider">see the product first</div>', unsafe_allow_html=True)
                if st.button("Experience sample company →", use_container_width=True, key="open_demo_choice"):
                    st.session_state["auth_view"] = "demo_choice"
                    st.rerun()
                st.markdown('<div class="login-trust"><span class="trust-dot"></span><span>Preview environment · production identity and MFA planned before public release</span></div>', unsafe_allow_html=True)


def render_login_and_workspace_gate() -> bool:
    _initialise_auth_state()
    if bool(st.session_state.get("app_logged_in")):
        return True
    view = st.session_state.get("auth_view", "login")
    _render_login_screen()
    if view == "onboarding":
        _workspace_wizard_dialog()
    elif view == "email":
        _email_sign_in_dialog()
    elif view == "demo_choice":
        _demo_choice_dialog()
    return False
