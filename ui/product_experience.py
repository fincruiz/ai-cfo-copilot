from __future__ import annotations

from html import escape

import streamlit as st

from services.demo_data import clear_demo_workspace, load_demo_workspace


TOUR_STEPS = [
    {
        "page": "dashboard",
        "eyebrow": "Step 1 · Executive command centre",
        "title": "See the financial story in under five seconds",
        "copy": "The dashboard combines readiness, performance KPIs, month-end status and AI priorities. In a live workspace these values update after validation.",
        "focus": ["Executive KPIs", "Close readiness", "AI priorities"],
        "action": "Scan the KPI cards and identify the strongest and weakest performance signal.",
    },
    {
        "page": "upload",
        "eyebrow": "Step 2 · Controlled data intake",
        "title": "Bring ERP exports into one validated finance model",
        "copy": "The Import Centre accepts GL, COA, budget, forecast, opening balance sheet and ageing files. Every file is checked before reports are generated.",
        "focus": ["GL & COA", "Validation", "Mapping review"],
        "action": "Open the templates and notice which files are mandatory versus optional.",
    },
    {
        "page": "reports",
        "eyebrow": "Step 3 · Board-ready reporting",
        "title": "Move from transactions to structured financial statements",
        "copy": "Review P&L, balance sheet, KPI, trend and variance views. Exact account-code mapping keeps similar-looking GL accounts separate.",
        "focus": ["P&L", "Balance sheet", "Variance"],
        "action": "Open P&L detail and compare the report hierarchy with the account-level drill-down.",
    },
    {
        "page": "working_capital",
        "eyebrow": "Step 4 · Cash discipline",
        "title": "Turn ageing data into collection and payment priorities",
        "copy": "Receivables, payables and cash-conversion signals help finance teams focus on exposures that affect liquidity now.",
        "focus": ["Receivables", "Payables", "Cash conversion"],
        "action": "Identify the largest ageing exposure and the recommended management action.",
    },
    {
        "page": "insights",
        "eyebrow": "Step 5 · AI CFO intelligence",
        "title": "Translate movements into explanations and actions",
        "copy": "AI commentary, anomaly detection and benchmark context explain what changed, why it matters and what management should do next.",
        "focus": ["Anomalies", "Commentary", "Recommendations"],
        "action": "Open the AI CFO and ask: What should management prioritise this month?",
    },
    {
        "page": "downloads",
        "eyebrow": "Step 6 · Management output",
        "title": "Take the analysis into meetings and decision workflows",
        "copy": "Export financial statements, KPI packs, working-capital analysis and management files for review outside the platform.",
        "focus": ["Report exports", "Management pack", "Audit trail"],
        "action": "Review the available outputs. Demo files remain temporary and are never saved as a real company.",
    },
]


def apply_product_motion() -> None:
    st.markdown(
        """
        <style>
        :root{--motion-fast:160ms;--motion-normal:300ms;--motion-slow:560ms}
        .block-container{animation:pageEnter .42s cubic-bezier(.2,.8,.2,1) both}
        div[data-testid="stMetric"]{border:1px solid rgba(148,163,184,.17);border-radius:18px;padding:.82rem .92rem;background:linear-gradient(180deg,rgba(255,255,255,.06),rgba(255,255,255,.02));box-shadow:0 11px 32px rgba(2,6,23,.12);transition:transform var(--motion-fast) ease,border-color var(--motion-fast) ease,box-shadow var(--motion-fast) ease;animation:cardRise .48s cubic-bezier(.2,.8,.2,1) both}
        div[data-testid="stMetric"]:hover{transform:translateY(-4px);border-color:rgba(96,165,250,.50);box-shadow:0 20px 46px rgba(37,99,235,.16)}
        div[data-testid="stDataFrame"],div[data-testid="stTable"]{animation:softFade .45s ease both;border-radius:16px;overflow:hidden}
        div[data-testid="stExpander"]{border-radius:16px !important;transition:transform var(--motion-fast) ease,border-color var(--motion-fast) ease}
        div[data-testid="stExpander"]:hover{transform:translateY(-1px);border-color:rgba(96,165,250,.36) !important}
        div[data-testid="stButton"]>button,div[data-testid="stFormSubmitButton"]>button,div[data-testid="stDownloadButton"]>button{transition:transform var(--motion-fast) ease,box-shadow var(--motion-fast) ease,border-color var(--motion-fast) ease !important}
        div[data-testid="stButton"]>button:hover,div[data-testid="stFormSubmitButton"]>button:hover,div[data-testid="stDownloadButton"]>button:hover{transform:translateY(-2px)}
        div[data-testid="stTabs"] button[role="tab"]{transition:color .18s ease,background .18s ease,transform .18s ease}div[data-testid="stTabs"] button[role="tab"]:hover{transform:translateY(-1px)}
        .stAlert{animation:toastSlide .36s cubic-bezier(.2,.8,.2,1) both}
        .demo-guide-shell{position:fixed;right:22px;bottom:22px;width:min(440px,calc(100vw - 32px));z-index:999990;border-radius:24px;padding:1rem 1.05rem;background:linear-gradient(145deg,rgba(8,15,28,.98),rgba(30,41,59,.97));border:1px solid rgba(96,165,250,.40);box-shadow:0 28px 80px rgba(0,0,0,.52);animation:guideEnter .38s cubic-bezier(.2,.8,.2,1) both;overflow:hidden}
        .demo-guide-shell:after{content:"";position:absolute;width:190px;height:190px;border-radius:50%;right:-90px;top:-110px;background:rgba(124,58,237,.20);filter:blur(28px);animation:guideOrb 5.5s ease-in-out infinite alternate;pointer-events:none}
        .guide-eyebrow{position:relative;z-index:2;color:#93c5fd;font-weight:900;font-size:.7rem;letter-spacing:.1em;text-transform:uppercase}.guide-title{position:relative;z-index:2;color:#fff;font-weight:950;font-size:1.22rem;letter-spacing:-.035em;margin:.18rem 0 .3rem}.guide-copy{position:relative;z-index:2;color:#dbeafe;font-size:.84rem;line-height:1.48}.guide-focus{position:relative;z-index:2;display:flex;flex-wrap:wrap;gap:.38rem;margin-top:.58rem}.guide-focus span{padding:.28rem .5rem;border-radius:999px;background:rgba(255,255,255,.09);border:1px solid rgba(255,255,255,.16);color:#fff;font-size:.69rem;font-weight:850}.guide-action{position:relative;z-index:2;margin-top:.58rem;color:#a5f3fc;font-size:.76rem;font-weight:800}
        .demo-ribbon{margin:.25rem 0 .8rem;padding:.72rem .86rem;border-radius:17px;background:linear-gradient(135deg,rgba(8,145,178,.13),rgba(79,70,229,.12));border:1px solid rgba(34,211,238,.28);display:flex;align-items:center;justify-content:space-between;gap:1rem;animation:softFade .38s ease both}.demo-ribbon b{color:#fff}.demo-ribbon span{color:#bae6fd;font-size:.79rem}.demo-mode-chip{padding:.3rem .55rem;border-radius:999px;background:rgba(34,211,238,.12);border:1px solid rgba(34,211,238,.28);color:#a5f3fc;font-size:.69rem;font-weight:900;white-space:nowrap}
        .st-key-demo_exit_primary button{background:linear-gradient(135deg,#991b1b,#ef4444) !important;color:#fff !important;border:1px solid rgba(254,202,202,.68) !important;box-shadow:0 13px 34px rgba(220,38,38,.32) !important;font-weight:950 !important}
        .st-key-demo_tour_next button,.st-key-demo_tour_start button{background:linear-gradient(135deg,#0284c7,#4f46e5) !important;color:#fff !important;font-weight:900 !important}
        @keyframes pageEnter{from{opacity:0;transform:translateY(9px)}to{opacity:1;transform:none}}@keyframes cardRise{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}@keyframes softFade{from{opacity:0}to{opacity:1}}@keyframes toastSlide{from{opacity:0;transform:translateX(14px)}to{opacity:1;transform:none}}@keyframes guideEnter{from{opacity:0;transform:translateY(18px) scale(.97)}to{opacity:1;transform:none}}@keyframes guideOrb{from{transform:translate(0,0)}to{transform:translate(-35px,30px)}}
        @media(max-width:768px){.demo-guide-shell{left:12px;right:12px;bottom:12px;width:auto}}
        @media(prefers-reduced-motion:reduce){*,*:before,*:after{animation-duration:.01ms !important;animation-iteration-count:1 !important;transition-duration:.01ms !important}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _tour_index() -> int:
    raw = int(st.session_state.get("demo_tour_step", 0) or 0)
    return max(0, min(raw, len(TOUR_STEPS) - 1))


def _go_to_step(index: int) -> None:
    index = max(0, min(index, len(TOUR_STEPS) - 1))
    st.session_state["demo_tour_step"] = index
    st.query_params["page"] = TOUR_STEPS[index]["page"]
    st.rerun()


@st.dialog("Welcome to the interactive product tour")
def _tour_intro_dialog() -> None:
    st.markdown(
        """
        <div style="padding:.2rem .1rem .7rem">
          <div style="font-size:2.1rem">🎬</div>
          <div style="color:#fff;font-size:1.35rem;font-weight:950;letter-spacing:-.035em;margin:.4rem 0">Six guided stops. One complete finance story.</div>
          <div style="color:#cbd5e1;line-height:1.55">The tour automatically moves through the sample dashboard, controlled import, board-ready reports, working capital, AI CFO insights and downloadable output.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns(2)
    if c1.button("Start tour →", type="primary", use_container_width=True, key="tour_intro_start"):
        st.session_state["demo_intro_pending"] = False
        _go_to_step(0)
    if c2.button("Explore freely instead", use_container_width=True, key="tour_intro_free"):
        st.session_state["demo_intro_pending"] = False
        st.session_state["demo_tour_active"] = False
        st.rerun()


def render_demo_experience() -> None:
    if st.session_state.get("auth_mode") != "demo":
        return
    if not st.session_state.get("demo_data_loaded"):
        load_demo_workspace(force=True)

    if st.session_state.get("demo_intro_pending"):
        _tour_intro_dialog()

    active = bool(st.session_state.get("demo_tour_active", False))
    step = _tour_index()
    item = TOUR_STEPS[step]

    if active:
        st.markdown(
            f'<div class="demo-ribbon"><div><b>Northstar Manufacturing · Interactive tour</b><br><span>Step {step + 1} of {len(TOUR_STEPS)} · sample data only</span></div><div class="demo-mode-chip">GUIDED MODE</div></div>',
            unsafe_allow_html=True,
        )
        focus_html = "".join(f"<span>{escape(str(value))}</span>" for value in item["focus"])
        st.markdown(
            f"""
            <div class="demo-guide-shell">
              <div class="guide-eyebrow">{escape(item['eyebrow'])}</div>
              <div class="guide-title">{escape(item['title'])}</div>
              <div class="guide-copy">{escape(item['copy'])}</div>
              <div class="guide-focus">{focus_html}</div>
              <div class="guide-action">Try this now: {escape(item['action'])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.progress((step + 1) / len(TOUR_STEPS), text=f"Interactive product tour · {step + 1} of {len(TOUR_STEPS)}")
        c1, c2, c3, c4, c5 = st.columns([.8, .9, 1.05, 1.05, 1.35])
        if c1.button("← Back", disabled=step == 0, use_container_width=True, key="demo_tour_back"):
            _go_to_step(step - 1)
        if c2.button("Next →", disabled=step == len(TOUR_STEPS) - 1, use_container_width=True, key="demo_tour_next"):
            _go_to_step(step + 1)
        if c3.button("Explore freely", use_container_width=True, key="demo_tour_free"):
            st.session_state["demo_tour_active"] = False
            st.rerun()
        if c4.button("Restart", use_container_width=True, key="demo_tour_restart"):
            _go_to_step(0)
        if c5.button("Exit Demo & Sign In", use_container_width=True, key="demo_exit_primary"):
            clear_demo_workspace(return_to_login=True)
            st.rerun()
    else:
        st.markdown(
            '<div class="demo-ribbon"><div><b>Northstar Manufacturing · Sample workspace</b><br><span>Browse freely. No coach prompts or forced navigation.</span></div><div class="demo-mode-chip">FREE EXPLORE</div></div>',
            unsafe_allow_html=True,
        )
        c1, c2, c3 = st.columns([1.15, 1, 1.35])
        if c1.button("Start interactive tour", use_container_width=True, key="demo_tour_start"):
            st.session_state["demo_tour_active"] = True
            st.session_state["demo_intro_pending"] = True
            st.rerun()
        if c2.button("Reset sample data", use_container_width=True, key="demo_reset_sample"):
            load_demo_workspace(force=True)
            st.session_state["demo_tour_active"] = False
            st.rerun()
        if c3.button("Exit Demo & Sign In", use_container_width=True, key="demo_exit_primary"):
            clear_demo_workspace(return_to_login=True)
            st.rerun()
