import os
import re
import json
from io import BytesIO
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.parse import urlencode

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


# Modular imports
from core.common import *
from core.normalizers import *
from core.excel_templates import *
from core.pipeline import *
from services.external_data import *
from services.downloads import *
from services.history_service import *
from services.ai_cfo import *
from services.research_service import *
from modules.reporting import *
from ui.validation_ui import *
from ui.v1_ui import *
from ui.login_flow import render_login_and_workspace_gate
from ui.product_experience import apply_product_motion, render_demo_experience
from ui.business_analytics_ui import render_business_analytics_page, render_market_research_page

st.set_page_config(page_title="AI CFO Copilot", layout="wide")

# Global application alignment and overlay fixes
st.markdown("""
<style>
html, body { overflow-x: hidden; }
[data-testid="stAppViewContainer"] > .main { width: 100%; }
.main .block-container {
    width: min(100%, 1500px) !important;
    max-width: 1500px !important;
    margin-left: auto !important;
    margin-right: auto !important;
    padding-left: clamp(1rem, 2.5vw, 2.5rem) !important;
    padding-right: clamp(1rem, 2.5vw, 2.5rem) !important;
}
[data-testid="stDialog"] > div, div[role="dialog"] > div {
    margin-left: auto !important;
    margin-right: auto !important;
}
/* Keep global AI launcher visible on every page and independent of page flow. */
.st-key-open_ai_cfo_global {
    position: fixed !important; right: 24px !important; bottom: 24px !important;
    z-index: 1000000 !important; width: 68px !important; height: 68px !important;
}
.st-key-open_ai_cfo_global > div { width:68px !important; height:68px !important; }
.st-key-open_ai_cfo_global button {
    position: static !important; width:68px !important; height:68px !important; min-height:68px !important;
    border-radius:50% !important; margin:0 !important; padding:0 !important;
}
/* Prevent tour coach from colliding with the AI launcher. */
.demo-guide-shell { right: 24px !important; bottom: 112px !important; }
@media (max-width: 900px) {
  .main .block-container {padding-left:.85rem !important;padding-right:.85rem !important;}
  .st-key-open_ai_cfo_global {right:14px !important;bottom:14px !important;}
  .demo-guide-shell {left:12px !important;right:12px !important;bottom:94px !important;width:auto !important;}
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
html, body, [class*="css"] {font-family: Arial, sans-serif;}
h1 {font-family: Arial, sans-serif !important; font-size: 34px !important; font-weight: 700 !important;}
h2 {font-family: Arial, sans-serif !important; font-size: 24px !important; font-weight: 700 !important;}
h3 {font-family: Arial, sans-serif !important; font-size: 20px !important; font-weight: 700 !important;}
div[data-testid="stDataFrame"] * {font-family: Arial, sans-serif !important; font-size: 13px !important;}
div[data-testid="stMetric"] * {font-family: Arial, sans-serif !important;}
button {font-family: Arial, sans-serif !important; font-size: 14px !important; font-weight: 600 !important;}
</style>
""", unsafe_allow_html=True)

# Floating animated AI CFO launcher
st.markdown("""
<style>
.floating-ai-cfo-wrap {
    position: fixed;
    right: 26px;
    bottom: 28px;
    z-index: 999999;
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 8px;
}
.floating-ai-cfo-label {
    background: rgba(17, 24, 39, 0.92);
    color: #ffffff;
    padding: 8px 12px;
    border-radius: 999px;
    font-size: 13px;
    box-shadow: 0 12px 30px rgba(0,0,0,0.18);
    animation: aiLabelFloat 3.2s ease-in-out infinite;
}
.floating-ai-cfo-button {
    width: 72px;
    height: 72px;
    border-radius: 50%;
    background: linear-gradient(135deg, #0f766e, #2563eb, #7c3aed);
    color: #fff !important;
    text-decoration: none !important;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 34px;
    box-shadow: 0 18px 45px rgba(37, 99, 235, 0.38);
    animation: aiFloat 2.8s ease-in-out infinite, aiPulse 1.8s ease-in-out infinite;
    border: 2px solid rgba(255,255,255,0.9);
}
.floating-ai-cfo-button:hover {
    transform: scale(1.08);
    box-shadow: 0 20px 55px rgba(124, 58, 237, 0.45);
}
.floating-ai-cfo-dot {
    position: absolute;
    right: 8px;
    bottom: 8px;
    width: 16px;
    height: 16px;
    background: #22c55e;
    border-radius: 50%;
    border: 2px solid white;
    animation: aiDot 1.4s ease-in-out infinite;
}
@keyframes aiFloat {
    0%, 100% { transform: translateY(0px) rotate(0deg); }
    50% { transform: translateY(-12px) rotate(2deg); }
}
@keyframes aiPulse {
    0% { box-shadow: 0 0 0 0 rgba(37,99,235,0.45), 0 18px 45px rgba(37,99,235,0.38); }
    70% { box-shadow: 0 0 0 18px rgba(37,99,235,0), 0 18px 45px rgba(37,99,235,0.38); }
    100% { box-shadow: 0 0 0 0 rgba(37,99,235,0), 0 18px 45px rgba(37,99,235,0.38); }
}
@keyframes aiLabelFloat {
    0%, 100% { transform: translateY(0px); opacity: 0.95; }
    50% { transform: translateY(-6px); opacity: 1; }
}
@keyframes aiDot {
    0%, 100% { transform: scale(1); opacity: 1; }
    50% { transform: scale(1.35); opacity: 0.75; }
}
@media (max-width: 768px) {
    .floating-ai-cfo-wrap { right: 16px; bottom: 18px; }
    .floating-ai-cfo-label { display: none; }
    .floating-ai-cfo-button { width: 62px; height: 62px; font-size: 29px; }
}
</style>
""", unsafe_allow_html=True)



# ----------------------------
# Generic helpers
# ----------------------------


# ----------------------------
# External FX / benchmark helpers
# ----------------------------



# ----------------------------
# Excel / template helpers
# ----------------------------


# ----------------------------
# Standardizers / normalizers
# ----------------------------


# ----------------------------
# Finance calculations
# ----------------------------










# ----------------------------
# Session defaults
# ----------------------------
for key in [
    "gl", "coa", "kpi_master", "latest_bs", "mapped", "pnl_mapped", "bs_mapped", "unmapped", "consolidated_pnl", "consolidated_bs", "consolidated_kpis", "branch_outputs", "branch_summary", "detected_branches", "validation_passed", "company_profile", "bs_disclaimer", "ai_commentary", "prior_pnl", "prior_bs", "prior_kpis", "save_run_preference", "anomaly_flags", "ar_df", "ap_df", "ar_summary", "ap_summary", "budget_df", "budget_compare", "budget_summary", "benchmark_df", "py_compare", "benchmark_compare", "monthly_actuals", "monthly_branch_actuals", "executive_summary_df", "forecast_pnl", "forecast_bs", "previous_year_pnl", "forecast_pnl_compare", "previous_year_pnl_compare", "fx_rate_info", "country_indicators", "external_benchmark_df", "consolidated_pnl_detail", "consolidated_bs_detail", "coa_duplicate_rows", "coa_mapping_review", "financial_logic_review", "last_validation_report", "reporting_structure", "ai_cfo_chat_messages", "app_logged_in", "auth_mode", "onboarding_step", "workspace_modules", "external_research_pack", "custom_research_result"
]:
    if key not in st.session_state:
        st.session_state[key] = None
if st.session_state["company_profile"] is None:
    st.session_state["company_profile"] = {}
if st.session_state["save_run_preference"] is None:
    st.session_state["save_run_preference"] = False
if st.session_state["ai_cfo_chat_messages"] is None:
    st.session_state["ai_cfo_chat_messages"] = []
if "ai_cfo_panel_open" not in st.session_state or st.session_state["ai_cfo_panel_open"] is None:
    st.session_state["ai_cfo_panel_open"] = False

# ----------------------------
# Global AI CFO overlay (available before and after login)
# ----------------------------
st.markdown("""
<style>
/* Floating AI CFO launcher. Clicking toggles the panel. */
.st-key-open_ai_cfo_global, div[class*="st-key-open_ai_cfo_global"] {
    position: fixed !important;
    right: 24px !important;
    bottom: 24px !important;
    width:72px !important;
    height:72px !important;
    z-index:2147483000 !important;
    margin:0 !important;
    padding:0 !important;
}
.st-key-open_ai_cfo_global button, div[class*="st-key-open_ai_cfo_global"] button {
    position: static !important;
    right: auto !important;
    bottom: auto !important;
    z-index: 2147483000 !important;
    width: 72px !important;
    height: 72px !important;
    min-height: 72px !important;
    border-radius: 50% !important;
    background: linear-gradient(135deg, #0f766e, #2563eb, #7c3aed) !important;
    color: #ffffff !important;
    font-size: 30px !important;
    border: 2px solid rgba(255,255,255,0.9) !important;
    box-shadow: 0 18px 45px rgba(37,99,235,0.38) !important;
    animation: aiFloat 2.8s ease-in-out infinite, aiPulse 1.8s ease-in-out infinite;
}
.st-key-open_ai_cfo_global button p { color: #ffffff !important; font-size: 30px !important; }
.st-key-open_ai_cfo_global button:hover {
    transform: scale(1.08) !important;
    box-shadow: 0 20px 55px rgba(124,58,237,0.45) !important;
}

/* Chatbot overlay panel. It is fixed and does NOT change page layout. */
.st-key-ai_cfo_overlay_panel {
    position: fixed !important;
    right: 24px !important;
    bottom: 112px !important;
    width: min(430px, calc(100vw - 32px)) !important;
    max-height: calc(100vh - 150px) !important;
    z-index: 2147482999 !important;
    overflow: auto !important;
    border: 1px solid rgba(96,165,250,0.38) !important;
    background: linear-gradient(180deg, rgba(15,23,42,0.98), rgba(17,24,39,0.98)) !important;
    border-radius: 24px !important;
    padding: 1rem !important;
    box-shadow: 0 24px 70px rgba(0,0,0,0.48) !important;
}
.st-key-ai_cfo_overlay_panel * { color: #f8fafc; }
.st-key-ai_cfo_overlay_panel label, .st-key-ai_cfo_overlay_panel p { color: #dbeafe !important; }
.st-key-ai_cfo_overlay_panel textarea, .st-key-ai_cfo_overlay_panel input {
    background: rgba(255,255,255,0.08) !important;
    color: #ffffff !important;
    border: 1px solid rgba(148,163,184,0.45) !important;
    border-radius: 12px !important;
}
.st-key-ai_cfo_overlay_panel textarea::placeholder, .st-key-ai_cfo_overlay_panel input::placeholder { color: #cbd5e1 !important; }
.ai-overlay-title {font-size:1.2rem;font-weight:850;color:#f8fafc;margin-bottom:0.25rem;}
.ai-overlay-sub {color:#cbd5e1;font-size:0.9rem;margin-bottom:0.85rem;line-height:1.35;}
.ai-bubble-user {background:#2563eb;color:#fff;border-radius:16px 16px 4px 16px;padding:0.68rem 0.78rem;margin:0.35rem 0 0.35rem auto;max-width:88%;font-size:0.92rem;}
.ai-bubble-assistant {background:rgba(255,255,255,0.08);color:#f8fafc;border:1px solid rgba(148,163,184,0.24);border-radius:16px 16px 16px 4px;padding:0.68rem 0.78rem;margin:0.35rem auto 0.35rem 0;max-width:94%;font-size:0.92rem;}
.ai-panel-note {color:#93c5fd;font-size:0.78rem;margin-top:0.45rem;}
@media (max-width: 768px) {
    .st-key-open_ai_cfo_global button { right: 16px !important; bottom: 18px !important; width: 62px !important; height: 62px !important; min-height: 62px !important; }
    .st-key-ai_cfo_overlay_panel { right: 12px !important; left: 12px !important; bottom: 92px !important; width: auto !important; max-height: calc(100vh - 120px) !important; }
}
</style>
""", unsafe_allow_html=True)

if st.button("🤖", key="open_ai_cfo_global", help="Open / close AI CFO Assistant"):
    st.session_state["ai_cfo_panel_open"] = not st.session_state.get("ai_cfo_panel_open", False)
    st.rerun()

# Global AI CFO overlay panel. Fixed position, so it does not disturb page alignment.
if st.session_state.get("ai_cfo_panel_open"):
    with st.container(key="ai_cfo_overlay_panel"):
        st.markdown('<div class="ai-overlay-title">🤖 AI CFO Assistant</div>', unsafe_allow_html=True)
        st.markdown('<div class="ai-overlay-sub">Ask about your financials, industry conditions, benchmarks, economic changes or management actions. Market research is automatically included in AI commentary when configured.</div>', unsafe_allow_html=True)

        research_pack = st.session_state.get("external_research_pack")
        research_status = "Market context ready" if research_pack else "Market context not loaded"
        st.caption(f"🌐 {research_status}")
        research_col, market_col = st.columns(2)
        if research_col.button("Refresh market context", use_container_width=True, key="ai_refresh_market_context"):
            with st.spinner("Researching current industry and economic context..."):
                pack = ensure_external_research_context(force=True, scan_type="Executive scan")
            if pack:
                st.success("Market context refreshed and connected to AI CFO.")
            else:
                st.warning("Market research could not be loaded. Check TAVILY_API_KEY and company profile.")
            st.rerun()
        if market_col.button("Open research page", use_container_width=True, key="ai_open_market_research"):
            st.session_state["ai_cfo_panel_open"] = False
            st.query_params["page"] = "research"
            st.rerun()

        close_col, clear_col = st.columns(2)
        if close_col.button("Close", use_container_width=True, key="close_ai_cfo_panel"):
            st.session_state["ai_cfo_panel_open"] = False
            st.rerun()
        if clear_col.button("Clear", use_container_width=True, key="clear_ai_cfo_panel"):
            st.session_state["ai_cfo_chat_messages"] = []
            st.rerun()

        chat_mode_inline = st.selectbox(
            "Mode",
            ["Auto", "General Help", "Data-specific CFO Analysis", "Internet & Benchmark Research"],
            key="global_ai_cfo_mode"
        )

        if not st.session_state.get("ai_cfo_chat_messages"):
            st.markdown('<div class="ai-bubble-assistant">Hi, I’m your AI CFO. Ask me about upload formats, mapping, validation, benchmarks, forecasts, or your uploaded financial data.</div>', unsafe_allow_html=True)

        for msg in st.session_state.get("ai_cfo_chat_messages", [])[-8:]:
            role = msg.get("role", "assistant")
            content = msg.get("content", "")
            if role == "user":
                st.markdown(f'<div class="ai-bubble-user">{content}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="ai-bubble-assistant">{content}</div>', unsafe_allow_html=True)

        with st.form("ai_cfo_overlay_form", clear_on_submit=True):
            user_question = st.text_area("Ask a question", placeholder="Example: Why is gross margin down?", height=85, key="ai_cfo_overlay_question")
            send_col, hint_col = st.columns([0.42, 0.58])
            send_clicked = send_col.form_submit_button("Send", use_container_width=True)
            hint_col.markdown('<div class="ai-panel-note">Use Close or click 🤖 again to hide this chat.</div>', unsafe_allow_html=True)

        if send_clicked and user_question.strip():
            st.session_state["ai_cfo_chat_messages"].append({"role": "user", "content": user_question.strip()})
            with st.spinner("AI CFO is thinking..."):
                inline_answer = answer_ai_cfo_question(user_question.strip(), mode=chat_mode_inline)
            st.session_state["ai_cfo_chat_messages"].append({"role": "assistant", "content": inline_answer})
            st.rerun()


# ----------------------------
# V1 Login / Workspace onboarding gate
# ----------------------------
if st.session_state.get("app_logged_in") is None:
    st.session_state["app_logged_in"] = False
if st.session_state.get("onboarding_step") is None:
    st.session_state["onboarding_step"] = 1
if st.session_state.get("workspace_modules") is None:
    st.session_state["workspace_modules"] = []

if not render_login_and_workspace_gate():
    st.stop()

# Safe motion system for the authenticated product experience.
apply_product_motion()

# ----------------------------
# UI - Product-style navigation
# ----------------------------
st.markdown("""
<style>
:root {--brand-1:#0f766e;--brand-2:#2563eb;--brand-3:#7c3aed;--ink:#111827;--muted:#6b7280;--line:#e5e7eb;--soft:#f8fafc;}
.block-container {max-width:1320px; padding-top:1.2rem;}
.main-card {padding:1.15rem 1.25rem;border:1px solid rgba(226,232,240,0.9);border-radius:18px;background:rgba(255,255,255,0.94);box-shadow:0 10px 30px rgba(15,23,42,0.06);margin-bottom:1rem;}
.hero-card {position:relative;overflow:hidden;min-height:295px;padding:2.1rem;border-radius:28px;color:white;background:linear-gradient(90deg, rgba(15,23,42,0.94) 0%, rgba(15,118,110,0.86) 48%, rgba(37,99,235,0.68) 100%),url('https://images.unsplash.com/photo-1554224155-6726b3ff858f?auto=format&fit=crop&w=1400&q=80');background-size:cover;background-position:center;box-shadow:0 22px 70px rgba(15,23,42,0.24);margin-bottom:1.25rem;}
.hero-kicker {display:inline-flex;align-items:center;gap:0.45rem;padding:0.42rem 0.72rem;background:rgba(255,255,255,0.14);border:1px solid rgba(255,255,255,0.24);border-radius:999px;font-weight:700;font-size:0.86rem;margin-bottom:0.9rem;}
.hero-title {font-size:2.65rem;line-height:1.05;font-weight:800;letter-spacing:-0.045em;max-width:780px;margin:0 0 0.8rem 0;}
.hero-subtitle {font-size:1.04rem;color:rgba(255,255,255,0.88);max-width:720px;margin-bottom:1.2rem;}
.hero-pill-row {display:flex;flex-wrap:wrap;gap:0.55rem;margin-top:1rem;}
.hero-pill {padding:0.52rem 0.78rem;border-radius:999px;background:rgba(255,255,255,0.14);border:1px solid rgba(255,255,255,0.2);color:white;font-weight:700;font-size:0.86rem;}
.feature-card {padding:1.1rem;border-radius:18px;border:1px solid rgba(226,232,240,0.95);background:linear-gradient(180deg,#ffffff 0%,#f8fafc 100%);box-shadow:0 8px 26px rgba(15,23,42,0.05);height:100%;}
.feature-icon {width:42px;height:42px;border-radius:14px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,rgba(15,118,110,0.12),rgba(37,99,235,0.12));color:#0f766e;font-size:1.35rem;margin-bottom:0.72rem;}
.feature-title {font-weight:800;color:var(--ink);margin-bottom:0.25rem;}
.feature-text {color:var(--muted);font-size:0.91rem;line-height:1.45;}
.setup-panel {border-radius:24px;padding:1.25rem;border:1px solid rgba(37,99,235,0.14);background:linear-gradient(180deg,rgba(239,246,255,0.82),rgba(255,255,255,0.96));box-shadow:0 12px 34px rgba(37,99,235,0.08);}
.section-title-row {display:flex;align-items:center;gap:0.65rem;margin:1.1rem 0 0.7rem 0;}
.section-title-icon {width:34px;height:34px;border-radius:12px;display:flex;align-items:center;justify-content:center;background:#eff6ff;color:#2563eb;font-size:1.08rem;}
.section-title-text {font-size:1.22rem;font-weight:800;color:#f8fafc;text-shadow:0 1px 1px rgba(0,0,0,0.25);}
.workflow-step {padding:0.8rem 1rem;border-radius:16px;background:#111827;border:1px solid rgba(148,163,184,0.30);text-align:center;font-weight:800;color:#f8fafc;box-shadow:0 6px 18px rgba(0,0,0,0.18);}
.workflow-step-done {background:#064e3b;border-color:#34d399;color:#d1fae5;}
.workflow-step-active {background:#1e3a8a;border-color:#60a5fa;color:#eff6ff;}
.small-muted {color:#6b7280;font-size:0.9rem;}
.alert-card {padding:0.9rem 1rem;border-radius:14px;background:#3f3f0f;border:1px solid #a3a329;margin-bottom:0.55rem;box-shadow:0 4px 14px rgba(251,146,60,0.08);color:#fef9c3;}
.alert-card b {color:#ffffff;}
.status-strip {display:grid;grid-template-columns:repeat(4,1fr);gap:0.8rem;margin-bottom:1rem;}
.status-mini {background:#ffffff;border:1px solid #e5e7eb;border-radius:16px;padding:0.85rem 1rem;box-shadow:0 7px 22px rgba(15,23,42,0.04);}
.status-mini b {font-size:0.82rem;color:#6b7280;display:block;margin-bottom:0.25rem;}
.status-mini span {font-weight:800;color:#111827;}
.upload-intro {padding:1rem 1.1rem;border-radius:18px;background:#111827;border:1px solid rgba(148,163,184,0.35);margin-bottom:1rem;color:#f8fafc;}
.upload-intro b {color:#ffffff;}
.upload-intro .small-muted {color:#cbd5e1;}
@media (max-width:900px) {.hero-title{font-size:2rem;}.hero-card{padding:1.4rem;min-height:auto;}.status-strip{grid-template-columns:repeat(2,1fr);}}
</style>
""", unsafe_allow_html=True)

# ----------------------------
# V1 Enterprise UX helpers (safe CSS - does not style Streamlit internals)
# ----------------------------
st.markdown("""
<style>
.v1-hero {
  border-radius: 28px; padding: 2.1rem; margin: 1.2rem 0;
  background: radial-gradient(circle at top left, rgba(37,99,235,.28), transparent 38%),
              linear-gradient(135deg, #0b1220 0%, #111827 45%, #0f172a 100%);
  border: 1px solid rgba(148,163,184,.22); box-shadow: 0 26px 70px rgba(0,0,0,.26);
}
.v1-hero h1 {font-size: 3rem !important; line-height: 1.02 !important; letter-spacing: -.055em !important; color: #fff !important; margin: .55rem 0 .65rem 0 !important;}
.v1-hero p {color: #cbd5e1; font-size: 1.08rem; max-width: 760px;}
.v1-badge {display:inline-flex; align-items:center; gap:.45rem; padding:.45rem .78rem; border-radius:999px; background:rgba(37,99,235,.16); border:1px solid rgba(96,165,250,.25); color:#bfdbfe; font-weight:800; font-size:.86rem;}
.v1-grid-2 {display:grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap:1rem; margin: 1rem 0;}
.v1-grid-3 {display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:1rem; margin: 1rem 0;}
.v1-card {border-radius:22px; padding:1.15rem; background:linear-gradient(180deg, rgba(17,24,39,.98), rgba(15,23,42,.98)); border:1px solid rgba(148,163,184,.24); box-shadow:0 16px 44px rgba(0,0,0,.18); color:#f8fafc;}
.v1-card h3 {margin:.25rem 0 .45rem 0 !important; color:#fff !important;}
.v1-card p, .v1-card li {color:#cbd5e1; font-size:.95rem;}
.v1-step {padding: .9rem 1rem; border-radius: 18px; background:#101827; border:1px solid rgba(148,163,184,.24); color:#e5e7eb;}
.v1-step.done {background:rgba(6,78,59,.56); border-color:rgba(52,211,153,.55);}
.v1-step.active {background:rgba(30,64,175,.55); border-color:rgba(96,165,250,.65);}
.v1-step b {display:block; color:#fff; margin-bottom:.15rem;}
.v1-kpi {border-radius:20px; padding:1rem; background:#0f172a; border:1px solid rgba(148,163,184,.22);}
.v1-kpi .label {color:#94a3b8; font-size:.82rem; font-weight:800; margin-bottom:.28rem;}
.v1-kpi .value {color:#fff; font-size:1.55rem; font-weight:900; letter-spacing:-.04em;}
.v1-divider {height:1px; background:linear-gradient(90deg, transparent, rgba(148,163,184,.38), transparent); margin:1rem 0;}
.v1-source-card {min-height:205px;}
.v1-source-card .tag {display:inline-flex; padding:.28rem .55rem; border-radius:999px; background:rgba(34,197,94,.12); border:1px solid rgba(34,197,94,.26); color:#86efac; font-size:.78rem; font-weight:800;}
@media (max-width: 900px) {.v1-grid-2,.v1-grid-3{grid-template-columns:1fr}.v1-hero h1{font-size:2.2rem!important}}
</style>
""", unsafe_allow_html=True)



st.markdown("""
<div style="display:flex;align-items:center;gap:0.65rem;margin-bottom:0.2rem;">
  <div style="width:40px;height:40px;border-radius:14px;background:linear-gradient(135deg,#0f766e,#2563eb);display:flex;align-items:center;justify-content:center;color:white;font-size:1.35rem;font-weight:800;">▣</div>
  <div>
    <div style="font-size:2rem;font-weight:850;letter-spacing:-0.04em;color:#f8fafc;line-height:1;">AI CFO Copilot</div>
    <div style="color:#cbd5e1;font-size:0.95rem;margin-top:0.2rem;">Finance reporting, validation, forecasting, benchmarking and AI analysis in one guided workspace</div>
  </div>
</div>
""", unsafe_allow_html=True)

pages = [
    "📊 Dashboard",
    "📉 Business Analytics",
    "🌐 Market Research",
    "🏠 Home",
    "📥 Import Centre",
    "📈 Reports",
    "💰 Working Capital Centre",
    "🧠 Insights",
    "📤 Downloads",
]

page_query = st.query_params.get("page", "")
page_key_map = {
    "home": "🏠 Home",
    "upload": "📥 Import Centre",
    "dashboard": "📊 Dashboard",
    "analytics": "📉 Business Analytics",
    "research": "🌐 Market Research",
    "reports": "📈 Reports",
    "working_capital": "💰 Working Capital Centre",
    "insights": "🧠 Insights",
    "downloads": "📤 Downloads",
}
page_slug_map = {v: k for k, v in page_key_map.items()}
selected_page = page_key_map.get(page_query, "📊 Dashboard")

# Demo mode has two genuinely different journeys: an interactive guided tour or free exploration.
render_demo_experience()

# Hide Streamlit's default sidebar completely. Navigation is handled by top action buttons.
st.markdown("""
<style>
[data-testid="stSidebar"] {display: none !important;}
[data-testid="collapsedControl"] {display: none !important;}
.block-container {padding-top: 1.15rem; max-width: 1440px;}
.top-nav-row {
    padding: 0.65rem 0.75rem;
    border: 1px solid rgba(148,163,184,0.20);
    border-radius: 22px;
    margin: 0.55rem 0 1rem 0;
    background: linear-gradient(135deg, rgba(15,23,42,0.96), rgba(17,24,39,0.92));
    box-shadow: 0 16px 45px rgba(2,6,23,0.24);
}
.nav-hint {
    color: #93c5fd;
    font-size: 0.78rem;
    font-weight: 900;
    letter-spacing: .08em;
    text-transform: uppercase;
    margin-bottom: 0.45rem;
}
div[data-testid="stButton"] > button {
    border-radius: 999px !important;
    border: 1px solid rgba(148,163,184,0.24) !important;
    background: rgba(15,23,42,0.86) !important;
    color: #e5e7eb !important;
    box-shadow: none !important;
    min-height: 2.35rem !important;
    transition: all .18s ease !important;
}
div[data-testid="stButton"] > button:hover {
    border-color: #60a5fa !important;
    background: linear-gradient(135deg, #1d4ed8, #2563eb) !important;
    color: #ffffff !important;
    transform: translateY(-1px);
    box-shadow: 0 10px 28px rgba(37,99,235,0.22) !important;
}
div[data-testid="stButton"] > button p, div[data-testid="stButton"] > button span {
    color: inherit !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="top-nav-row">', unsafe_allow_html=True)
st.markdown('<div class="nav-hint">CFO Workflow</div>', unsafe_allow_html=True)
primary_pages = ["📊 Dashboard", "📉 Business Analytics", "🌐 Market Research", "📈 Reports", "🧠 Insights"]
secondary_pages = ["🏠 Home", "📥 Import Centre", "💰 Working Capital Centre", "📤 Downloads"]
for row_index, row_pages in enumerate([primary_pages, secondary_pages]):
    nav_cols = st.columns(len(row_pages))
    for nav_page, nav_col in zip(row_pages, nav_cols):
        is_current = selected_page == nav_page
        button_label = ("✓ " if is_current else "") + nav_page
        if nav_col.button(button_label, key=f"nav_{row_index}_{page_slug_map[nav_page]}", use_container_width=True):
            st.query_params["page"] = page_slug_map[nav_page]
            st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

profile = st.session_state.get("company_profile", {}) or {}
report = st.session_state.get("last_validation_report") or {}
score = report.get("score", 100 if isinstance(st.session_state.get("mapped"), pd.DataFrame) and not st.session_state.get("mapped").empty else 0)
meta_cols = st.columns(4)
meta_cols[0].caption(f"Company: {profile.get('Company Name', 'Not set')}")
meta_cols[1].caption(f"Industry: {profile.get('Industry', 'Not set')}")
meta_cols[2].caption(f"Period: {get_report_period_label(profile)}")
meta_cols[3].caption(f"Readiness: {score}/100")

# Workflow status is now shown inside the Dashboard command centre instead of as large global blocks.
profile_done = bool((st.session_state.get("company_profile") or {}).get("Company Name"))
data_loaded = st.session_state.get("mapped") is not None
validation_ok = bool(st.session_state.get("validation_passed")) if data_loaded else False
reports_ready = st.session_state.get("consolidated_pnl") is not None
insights_ready = bool(st.session_state.get("ai_commentary"))

if selected_page == "🏠 Home":
    render_v1_onboarding_dialog()
    profile = st.session_state.get("company_profile", {}) or {}
    report = st.session_state.get("last_validation_report") or {}
    score = report.get("score", 100 if isinstance(st.session_state.get("mapped"), pd.DataFrame) and not st.session_state.get("mapped").empty else 0)
    profile_done = bool(profile.get("Company Name"))
    data_loaded = st.session_state.get("mapped") is not None
    validation_ok = bool(st.session_state.get("validation_passed")) if data_loaded else False
    reports_ready = st.session_state.get("consolidated_pnl") is not None
    insights_ready = bool(st.session_state.get("ai_commentary"))
    render_v1_home_intro(profile, score, profile_done, data_loaded, validation_ok, reports_ready, insights_ready)
    render_v1_data_source_cards()
    profile = st.session_state.get("company_profile", {}) or {}
    report = st.session_state.get("last_validation_report") or {}
    critical = report.get("critical", [])
    warnings = report.get("warnings", [])
    recommendations = report.get("recommendations", [])
    score = report.get("score", 100 if isinstance(st.session_state.get("mapped"), pd.DataFrame) and not st.session_state.get("mapped").empty else 0)

    st.markdown("""
    <div class="hero-card">
        <div class="hero-kicker">✨ AI-powered finance workspace</div>
        <div class="hero-title">Turn messy finance data into board-ready decisions.</div>
        <div class="hero-subtitle">
            Upload GL, COA, budgets, forecast packs and AR/AP ageing. The app validates your data, flags mapping risks,
            builds P&L, balance sheet, KPI packs, dashboards and lets users ask an AI CFO questions.
        </div>
        <div class="hero-pill-row">
            <span class="hero-pill">📊 Management dashboards</span>
            <span class="hero-pill">✅ Data validation centre</span>
            <span class="hero-pill">🧠 AI CFO chatbot</span>
            <span class="hero-pill">🌍 FX & benchmarks</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="status-strip">
        <div class="status-mini"><b>Data Readiness</b><span>{score}/100</span></div>
        <div class="status-mini"><b>Critical Errors</b><span>{len(critical)}</span></div>
        <div class="status-mini"><b>Warnings</b><span>{len(warnings)}</span></div>
        <div class="status-mini"><b>Recommendations</b><span>{len(recommendations)}</span></div>
    </div>
    """, unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        st.markdown('<div class="feature-card"><div class="feature-icon">🏢</div><div class="feature-title">1. Company Setup</div><div class="feature-text">Capture company, country, industry, period and reporting structure before any upload.</div></div>', unsafe_allow_html=True)
    with f2:
        st.markdown('<div class="feature-card"><div class="feature-icon">📁</div><div class="feature-title">2. Validate Uploads</div><div class="feature-text">Upload GL and COA. The Validation Centre checks missing columns, duplicates and mapping risks.</div></div>', unsafe_allow_html=True)
    with f3:
        st.markdown('<div class="feature-card"><div class="feature-icon">📈</div><div class="feature-title">3. Generate Reports</div><div class="feature-text">Create P&L, balance sheet, branch packs, KPI summary, variance and working capital views.</div></div>', unsafe_allow_html=True)
    with f4:
        st.markdown('<div class="feature-card"><div class="feature-icon">🤖</div><div class="feature-title">4. Ask AI CFO</div><div class="feature-text">Ask generic upload questions before data, then data-specific CFO questions after upload.</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title-row"><div class="section-title-icon">🏢</div><div class="section-title-text">Company Setup</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="setup-panel">', unsafe_allow_html=True)

    industry_options = ["Select Industry", "Manufacturing", "Wholesale / Distribution", "Retail", "Professional Services", "Construction", "Logistics", "Hospitality", "Healthcare", "Technology", "Other"]
    country_options = ["Select Country", "Australia", "India", "United States", "United Kingdom", "Canada", "New Zealand", "Other"]
    currency_options = ["Select Currency", "AUD", "INR", "USD", "GBP", "CAD", "NZD", "Other"]
    period_options = ["Monthly", "Quarterly", "Annual"]
    structure_options = ["Consolidated Only", "Branch / Business Unit Reporting"]

    def option_index(options, value, default=0):
        return options.index(value) if value in options else default

    c1, c2 = st.columns(2)
    with c1:
        company_name = st.text_input("Company Name *", value=profile.get("Company Name", ""), key="home_company_name")
        industry = st.selectbox("Industry", industry_options, index=option_index(industry_options, profile.get("Industry", "Select Industry")), key="home_industry")
        country = st.selectbox("Country", country_options, index=option_index(country_options, profile.get("Country", "Select Country")), key="home_country")
        state_region = st.text_input("State / Region", value=profile.get("State / Region", ""), key="home_state_region")
        financial_year = st.text_input("Financial Year", value=profile.get("Financial Year", ""), placeholder="Example: FY2026 or 2025-26", key="home_financial_year")
        report_period = st.text_input("Report Period *", value=profile.get("Report Period", ""), placeholder="Example: April 2026 or Q1 FY2026", key="home_report_period")
        period_start_date = st.date_input("Period Start Date", value=pd.to_datetime(profile.get("Period Start Date"), errors="coerce").date() if profile.get("Period Start Date") else None, key="home_period_start_date")
    with c2:
        currency = st.selectbox("Currency", currency_options, index=option_index(currency_options, profile.get("Currency", "Select Currency")), key="home_currency")
        tax_identifier = st.text_input("Tax Identifier / ABN / GSTIN (Optional)", value=profile.get("Tax Identifier", ""), key="home_tax_identifier")
        reporting_period = st.selectbox("Period Type", period_options, index=option_index(period_options, profile.get("Reporting Period", "Monthly")), key="home_reporting_period")
        period_end_date = st.date_input("Period End Date", value=pd.to_datetime(profile.get("Period End Date"), errors="coerce").date() if profile.get("Period End Date") else None, key="home_period_end_date")
        reporting_structure = st.radio("Reporting Structure", structure_options, index=option_index(structure_options, profile.get("Reporting Structure", "Consolidated Only")), key="home_reporting_structure", help="If Consolidated Only is selected, Branch is optional in GL and the app uses a default Consolidated unit.")
        benchmark_group = st.text_input("Benchmark Group (Optional)", value=profile.get("Benchmark Group", ""), key="home_benchmark_group")

    business_notes = st.text_area("Business Notes (Optional)", value=profile.get("Business Notes", ""), key="home_business_notes")
    save_run_preference = st.checkbox("Save this run for future comparison", value=st.session_state["save_run_preference"], key="home_save_run")

    h1, h2, h3 = st.columns([1.2, 1, 1])
    with h1:
        if st.button("Save Company Profile", use_container_width=True, key="home_save_company_profile"):
            if not company_name.strip():
                st.error("Company Name is mandatory.")
            elif industry == "Select Industry" or country == "Select Country":
                st.error("Please select at least Industry and Country.")
            elif not report_period.strip():
                st.error("Report Period is mandatory. Example: April 2026 or Q1 FY2026.")
            elif period_start_date and period_end_date and pd.to_datetime(period_start_date) > pd.to_datetime(period_end_date):
                st.error("Period Start Date cannot be after Period End Date.")
            else:
                st.session_state["company_profile"] = {"Company Name": company_name.strip(), "Industry": industry, "Country": country, "State / Region": state_region, "Financial Year": financial_year, "Report Period": report_period.strip(), "Period Start Date": str(period_start_date) if period_start_date else "", "Period End Date": str(period_end_date) if period_end_date else "", "Currency": currency if currency != "Select Currency" else "", "Tax Identifier": tax_identifier, "Reporting Period": reporting_period, "Reporting Structure": reporting_structure, "Benchmark Group": benchmark_group, "Business Notes": business_notes}
                st.session_state["reporting_structure"] = reporting_structure
                st.session_state["save_run_preference"] = save_run_preference
                st.success("Company profile saved. You can now go to Import Centre.")
    with h2:
        if st.button("Go to Import Centre", use_container_width=True, key="home_go_upload"):
            st.query_params["page"] = "upload"
            st.rerun()
    with h3:
        if st.button("Ask AI CFO", use_container_width=True, key="home_go_ai"):
            st.session_state["ai_cfo_panel_open"] = True
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("External Data & Benchmarks (Optional)", expanded=False):
        profile = st.session_state.get("company_profile", {})
        selected_country = profile.get("Country", "Australia") if profile else "Australia"
        selected_industry = profile.get("Industry", "Other") if profile else "Other"
        selected_currency = profile.get("Currency", "AUD") if profile else "AUD"
        default_target_currency = selected_currency if selected_currency and selected_currency != "Select Currency" else currency_for_country(selected_country)

        st.info("Optional setup for FX, country indicators and starter industry benchmarks. Uploaded benchmark files are still added from Import Centre, but external benchmark setup belongs here with the company profile.")
        fx1, fx2, fx3 = st.columns(3)
        with fx1:
            fx_base = st.selectbox("FX Base Currency", ["AUD", "INR", "USD", "GBP", "CAD", "NZD", "EUR"], index=0)
        with fx2:
            currency_options = ["AUD", "INR", "USD", "GBP", "CAD", "NZD", "EUR"]
            target_index = currency_options.index(default_target_currency) if default_target_currency in currency_options else 0
            fx_target = st.selectbox("FX Target Currency", currency_options, index=target_index)
        with fx3:
            fx_date = st.text_input("FX Date", value="latest", help="Use latest or YYYY-MM-DD")

        b1, b2, b3 = st.columns(3)
        with b1:
            if st.button("Fetch FX Rate", use_container_width=True):
                try:
                    st.session_state["fx_rate_info"] = fetch_fx_rate(fx_base, fx_target, fx_date)
                    st.success("FX rate fetched.")
                except Exception as e:
                    st.error(f"FX fetch failed: {e}")
        with b2:
            if st.button("Fetch Country Indicators", use_container_width=True):
                try:
                    st.session_state["country_indicators"] = fetch_country_indicators(selected_country)
                    st.success("Country indicators fetched.")
                except Exception as e:
                    st.error(f"Country indicator fetch failed: {e}")
        with b3:
            if st.button("Load Starter Industry Benchmarks", use_container_width=True):
                st.session_state["external_benchmark_df"] = get_builtin_industry_benchmarks(selected_industry, selected_country)
                st.success("Starter benchmark set loaded.")

        if st.session_state.get("fx_rate_info"):
            st.markdown("**FX Rate**")
            st.dataframe(pd.DataFrame([st.session_state["fx_rate_info"]]), use_container_width=True, hide_index=True)
        if st.session_state.get("country_indicators") is not None:
            st.markdown("**Country Indicators**")
            st.dataframe(st.session_state["country_indicators"], use_container_width=True, hide_index=True)
        if st.session_state.get("external_benchmark_df") is not None:
            st.markdown("**Loaded Benchmark Data**")
            st.dataframe(st.session_state["external_benchmark_df"], use_container_width=True, hide_index=True)



    profile = st.session_state.get("company_profile", {}) or {}
    if profile:
        st.markdown('<div class="section-title-row"><div class="section-title-icon">📋</div><div class="section-title-text">Current Company Profile</div></div>', unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(profile.items(), columns=["Field", "Value"]), use_container_width=True, hide_index=True)

    if critical or warnings or recommendations:
        st.markdown('<div class="section-title-row"><div class="section-title-icon">⚠️</div><div class="section-title-text">Items Needing Attention</div></div>', unsafe_allow_html=True)
        for item in (critical + warnings + recommendations)[:8]:
            st.markdown(f'<div class="alert-card"><b>{item.get("Area", "Review")}</b><br>{item.get("Issue", "")}</div>', unsafe_allow_html=True)
    elif st.session_state.get("mapped") is not None:
        st.success("No validation errors and no recommendations found in the last validation run.")
    else:
        st.info("Save company profile, then upload data to activate the Validation Centre.")

elif selected_page == "📥 Import Centre":
    profile = st.session_state.get("company_profile", {}) or {}
    if not profile or not profile.get("Company Name"):
        st.warning("Please complete Company Setup on the Home page before importing files.")
        if st.button("Go to Home Setup", use_container_width=True, key="upload_go_home_setup"):
            st.query_params["page"] = "home"
            st.rerun()
    else:
        render_v1_import_intro(profile)
        st.markdown(f"""
        <div class="upload-intro">
            <b>Uploading for:</b> {profile.get("Company Name", "")} &nbsp; | &nbsp;
            <b>Industry:</b> {profile.get("Industry", "")} &nbsp; | &nbsp;
            <b>Country:</b> {profile.get("Country", "")} &nbsp; | &nbsp;
            <b>Reporting:</b> {profile.get("Reporting Structure", "Consolidated Only")}
            <br><span class="small-muted">Company details are managed on the Home page. This section is only for files, templates, FX, benchmarks and validation.</span>
        </div>
        """, unsafe_allow_html=True)

    with st.expander("Current Period Uploads", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            gl_file = st.file_uploader("Current GL Report", type=["xlsx"])
            mapping_file = st.file_uploader("COA Mapping", type=["xlsx"])
            budget_file = st.file_uploader("Budget Data (Optional)", type=["xlsx"])
        with c2:
            kpi_file = st.file_uploader("KPI Master (Optional)", type=["xlsx"])
            latest_bs_file = st.file_uploader("Latest Previous Balance Sheet (Optional)", type=["xlsx"])
            forecast_pnl_file = st.file_uploader("Forecast P&L (Optional)", type=["xlsx"])
        with c3:
            forecast_bs_file = st.file_uploader("Forecast Balance Sheet (Optional)", type=["xlsx"])
            ar_file = st.file_uploader("AR Ageing (Optional)", type=["xlsx"])
            ap_file = st.file_uploader("AP Ageing (Optional)", type=["xlsx"])
            benchmark_file = st.file_uploader("Industry Benchmark File (Optional)", type=["xlsx"])
        previous_year_pnl_file = st.file_uploader("Previous Year P&L (Optional)", type=["xlsx"])

        current_reporting_structure = st.session_state.get("company_profile", {}).get("Reporting Structure", "Consolidated Only")
        if current_reporting_structure == "Consolidated Only":
            st.info("Branch / Business Unit is optional for this company. If missing in GL, the app will use Consolidated automatically.")
        else:
            st.info("Branch / Business Unit Reporting is enabled. Branch column is mandatory in GL.")

        duplicate_resolution_confirmed = st.checkbox(
            "If duplicate COA Account codes are found, I have reviewed them and approve keeping the first row for each duplicate Account code",
            value=False,
            help="The original Excel file is not changed. This only cleans a processing copy after you approve."
        )

        if st.button("Validate & Upload Files", use_container_width=True):
            critical_items, warning_items, recommendation_items, info_items = [], [], [], []
            validation_success, loaded_files, previews = [], {}, {}

            profile = st.session_state.get("company_profile", {})
            reporting_structure = profile.get("Reporting Structure", "Consolidated Only")
            branch_required = reporting_structure == "Branch / Business Unit Reporting"

            def add_item(bucket, area, issue, recommendation=""):
                bucket.append({"Area": area, "Issue": str(issue), "Recommendation": recommendation})

            def log_success(file):
                validation_success.append({"File": file, "Status": "Valid"})

            if not profile or not profile.get("Company Name", "").strip():
                add_item(critical_items, "Company Profile", "Company profile is not saved.", "Save Company Profile before uploading files.")
            if gl_file is None:
                add_item(critical_items, "Current GL Report", "Mandatory file missing.", "Upload the Current GL Report template/file.")
            if mapping_file is None:
                add_item(critical_items, "COA Mapping", "Mandatory file missing.", "Upload the COA Mapping template/file.")

            gl_required_cols = ["Account code", "Debit", "Credit"] + (["Branch"] if branch_required else [])
            if not branch_required:
                add_item(info_items, "Reporting Structure", "Consolidated Only selected.", "Branch column is optional. Missing/blank Branch values will be treated as Consolidated.")
            else:
                add_item(info_items, "Reporting Structure", "Branch / Business Unit Reporting selected.", "Branch column is mandatory in the GL.")

            file_checks = [
                ("Current GL Report", gl_file, lambda f: standardize_key_columns(pd.read_excel(f), pd.DataFrame())[0], gl_required_cols, "gl"),
                ("COA Mapping", mapping_file, lambda f: standardize_key_columns(pd.DataFrame(), pd.read_excel(f))[1], ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], "coa"),
                ("KPI Master", kpi_file, lambda f: standardize_key_columns(pd.DataFrame(), pd.DataFrame(), pd.read_excel(f))[2], ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"], "kpi"),
                ("Latest Previous Balance Sheet", latest_bs_file, lambda f: normalize_uploaded_bs(pd.read_excel(f), "Latest Previous Balance Sheet"), ["Reporting Group", "Reporting Subgroup", "Balance"], "latest_bs"),
                ("Budget Data", budget_file, lambda f: normalize_plan_df(pd.read_excel(f), "Budget Data"), ["Month", "Reporting Group", "Amount"], "budget"),
                ("Forecast P&L", forecast_pnl_file, lambda f: normalize_uploaded_pnl(pd.read_excel(f), "Forecast P&L"), ["Reporting Group", "Reporting Subgroup", "Report Value"], "forecast_pnl"),
                ("Forecast Balance Sheet", forecast_bs_file, lambda f: normalize_uploaded_bs(pd.read_excel(f), "Forecast Balance Sheet"), ["Reporting Group", "Reporting Subgroup", "Balance"], "forecast_bs"),
                ("Previous Year P&L", previous_year_pnl_file, lambda f: normalize_uploaded_pnl(pd.read_excel(f), "Previous Year P&L"), ["Reporting Group", "Reporting Subgroup", "Report Value"], "previous_year_pnl"),
                ("AR Ageing", ar_file, lambda f: normalize_ageing_df(pd.read_excel(f), "AR"), ["Party Name", "Outstanding Amount"], "ar"),
                ("AP Ageing", ap_file, lambda f: normalize_ageing_df(pd.read_excel(f), "AP"), ["Party Name", "Outstanding Amount"], "ap"),
                ("Industry Benchmark File", benchmark_file, lambda f: normalize_benchmark_df(pd.read_excel(f)), ["Metric", "Benchmark Value"], "benchmark"),
            ]

            for file_label, file_obj, loader, required, key in file_checks:
                if file_obj is None:
                    if key in ["kpi", "latest_bs", "budget", "forecast_pnl", "forecast_bs", "previous_year_pnl", "ar", "ap", "benchmark"]:
                        add_item(info_items, file_label, "Optional file not uploaded.", "Upload this file if you want this analysis included.")
                    continue
                try:
                    df = loader(file_obj)
                    validate_required_columns(df, required, file_label)
                    if key == "gl" and not branch_required and "Branch" not in df.columns:
                        df["Branch"] = "Consolidated"
                    if key == "gl":
                        for period_issue in validate_gl_dates_against_profile(df, profile):
                            add_item(warning_items, period_issue["Area"], period_issue["Issue"], period_issue["Recommendation"])
                    if key == "coa":
                        duplicate_rows = find_coa_duplicate_rows(df)
                        st.session_state["coa_duplicate_rows"] = duplicate_rows
                        if not duplicate_rows.empty:
                            add_item(warning_items, "COA Mapping", f"Duplicate Account code rows found: {duplicate_rows['Account code'].nunique()} duplicated code(s).", "Review duplicate mapping rows. Tick the duplicate confirmation box to keep the first row for each duplicate for this run.")
                            if not duplicate_resolution_confirmed:
                                add_item(critical_items, "COA Mapping", "Duplicate Account code rows require user confirmation before processing.", "Fix the COA file or tick the duplicate confirmation checkbox after reviewing duplicates.")
                        mapping_review = build_coa_mapping_review(resolve_coa_duplicate_rows(df, keep="first") if not duplicate_rows.empty else df)
                        loaded_files["coa_mapping_review"] = mapping_review
                        if mapping_review is not None and not mapping_review.empty:
                            for _, r in mapping_review.head(10).iterrows():
                                add_item(recommendation_items, "COA Mapping Review", f"{r.get('Account code', '')} {r.get('Account Name', '')}: mapped to {r.get('Current Mapping', '')}", f"Suggested review: {r.get('Suggested Mapping', '')}. {r.get('Reason', '')}")
                    loaded_files[key] = df
                    previews[file_label] = df
                    log_success(file_label)
                except Exception as e:
                    raw_df = None
                    try:
                        raw_df = clean_columns(pd.read_excel(file_obj))
                    except Exception:
                        pass
                    found_cols = f" Found columns: {list(raw_df.columns)}" if raw_df is not None else ""
                    add_item(critical_items, file_label, f"{e}.{found_cols}", "Correct the file using the sample template and upload again.")

            if critical_items:
                render_validation_centre(critical_items, warning_items, recommendation_items, info_items, previews, block_processing=True)
                st.stop()

            try:
                gl, coa, kpi_master, latest_bs, mapped, pnl_mapped, bs_mapped, unmapped = prepare_data(
                    gl_file,
                    mapping_file,
                    kpi_file,
                    latest_bs_file,
                    allow_duplicate_coa_cleanup=duplicate_resolution_confirmed,
                    reporting_structure=reporting_structure,
                )
                consolidated_pnl = build_pnl(pnl_mapped)
                consolidated_pnl_detail = build_pnl_detail(pnl_mapped)
                coa_mapping_review = build_coa_mapping_review(coa)
                financial_logic_review = build_financial_logic_review(consolidated_pnl)
                if financial_logic_review is not None and not financial_logic_review.empty:
                    for _, r in financial_logic_review.head(10).iterrows():
                        add_item(recommendation_items, "Financial Logic Review", r.get("Issue", "Financial logic item"), r.get("Recommendation", "Review financial classification and source data."))
                if unmapped is not None and not unmapped.empty:
                    add_item(warning_items, "GL Mapping", f"{len(unmapped)} GL row(s) are unmapped.", "Review unmapped account codes and update COA Mapping.")

                current_bs = build_balance_sheet_from_gl(bs_mapped)
                consolidated_bs_detail = build_balance_sheet_detail(bs_mapped)
                bs_disclaimer = None
                if latest_bs is not None:
                    consolidated_bs = combine_opening_and_current_bs(latest_bs, current_bs)
                else:
                    consolidated_bs = current_bs
                    bs_disclaimer = "Balance Sheet may not fully match because opening balances were not provided."
                    add_item(info_items, "Balance Sheet", "Latest previous Balance Sheet not uploaded.", "Upload opening/latest BS if you want stronger balance-sheet continuity.")
                consolidated_kpis = build_kpis(pnl_mapped, kpi_master) if kpi_master is not None else None
                detected_branches = sorted(pnl_mapped["Branch"].dropna().unique().tolist()) if not pnl_mapped.empty else []
                branch_outputs, branch_summary_rows = {}, []
                for branch in detected_branches:
                    branch_df = pnl_mapped[pnl_mapped["Branch"] == branch].copy()
                    branch_pnl = build_pnl(branch_df)
                    branch_pnl_detail = build_pnl_detail(branch_df)
                    branch_kpis = build_kpis(branch_df, kpi_master) if kpi_master is not None else None
                    branch_outputs[branch] = {"pnl": branch_pnl, "pnl_detail": branch_pnl_detail, "kpis": branch_kpis}
                    if branch_kpis is not None:
                        row = {"Branch": branch}
                        for _, r in branch_kpis.iterrows():
                            row[r["KPI"]] = r["Display Value"]
                        branch_summary_rows.append(row)
                branch_summary = pd.DataFrame(branch_summary_rows) if branch_summary_rows else pd.DataFrame()
                ar_df, ap_df = loaded_files.get("ar"), loaded_files.get("ap")
                ar_summary = build_ageing_summary(ar_df, "AR") if ar_df is not None else None
                ap_summary = build_ageing_summary(ap_df, "AP") if ap_df is not None else None
                budget_df = loaded_files.get("budget")
                uploaded_benchmark_df = loaded_files.get("benchmark")
                benchmark_df = merge_benchmark_sources(uploaded_benchmark_df, st.session_state.get("external_benchmark_df"))
                forecast_pnl = loaded_files.get("forecast_pnl")
                forecast_bs = loaded_files.get("forecast_bs")
                previous_year_pnl = loaded_files.get("previous_year_pnl")
                actuals_df = build_actuals_by_branch_reporting_group(pnl_mapped)
                budget_compare = compare_plan_vs_actual(actuals_df, budget_df, "Budget") if budget_df is not None else None
                budget_summary = summarize_plan_vs_actual(budget_compare, "Budget") if budget_compare is not None else None
                forecast_pnl_compare = compare_pnl_to_forecast(consolidated_pnl, forecast_pnl) if forecast_pnl is not None else None
                previous_year_pnl_compare = compare_pnl_to_previous_year(consolidated_pnl, previous_year_pnl) if previous_year_pnl is not None else None
                py_compare = build_py_comparison(consolidated_kpis, st.session_state.get("prior_kpis"))
                benchmark_compare = build_benchmark_comparison(consolidated_kpis, benchmark_df, ar_summary, ap_summary)
                monthly_actuals = build_monthly_actuals(pnl_mapped)
                monthly_branch_actuals = build_monthly_branch_actuals(pnl_mapped)
                executive_summary_df = build_executive_summary(consolidated_kpis, ar_summary=ar_summary, ap_summary=ap_summary, budget_summary=budget_summary, benchmark_compare=benchmark_compare, forecast_pnl_compare=forecast_pnl_compare, previous_year_pnl_compare=previous_year_pnl_compare)
                anomaly_flags = detect_anomalies(consolidated_kpis, prior_kpis=st.session_state.get("prior_kpis"), ar_summary=ar_summary, ap_summary=ap_summary, budget_summary=budget_summary, forecast_pnl_compare=forecast_pnl_compare) if consolidated_kpis is not None else []
                for flag in anomaly_flags:
                    add_item(recommendation_items, "Anomaly Review", flag, "Review source data and mapping before relying on final reports.")

                last_validation_report = {
                    "critical": critical_items,
                    "warnings": warning_items,
                    "recommendations": recommendation_items,
                    "info": info_items,
                    "score": calculate_validation_score(len(critical_items), len(warning_items), len(recommendation_items)),
                }

                for k, v in {
                    "gl": gl, "coa": coa, "kpi_master": kpi_master, "latest_bs": latest_bs, "mapped": mapped, "pnl_mapped": pnl_mapped, "bs_mapped": bs_mapped, "unmapped": unmapped, "consolidated_pnl": consolidated_pnl, "consolidated_pnl_detail": consolidated_pnl_detail, "consolidated_bs": consolidated_bs, "consolidated_bs_detail": consolidated_bs_detail, "consolidated_kpis": consolidated_kpis, "branch_outputs": branch_outputs, "branch_summary": branch_summary, "detected_branches": detected_branches, "validation_passed": unmapped.empty, "bs_disclaimer": bs_disclaimer, "ai_commentary": None, "ar_df": ar_df, "ap_df": ap_df, "ar_summary": ar_summary, "ap_summary": ap_summary, "budget_df": budget_df, "budget_compare": budget_compare, "budget_summary": budget_summary, "benchmark_df": benchmark_df, "benchmark_compare": benchmark_compare, "py_compare": py_compare, "monthly_actuals": monthly_actuals, "monthly_branch_actuals": monthly_branch_actuals, "executive_summary_df": executive_summary_df, "forecast_pnl": forecast_pnl, "forecast_bs": forecast_bs, "previous_year_pnl": previous_year_pnl, "forecast_pnl_compare": forecast_pnl_compare, "previous_year_pnl_compare": previous_year_pnl_compare, "anomaly_flags": anomaly_flags, "coa_mapping_review": coa_mapping_review, "financial_logic_review": financial_logic_review, "last_validation_report": last_validation_report, "reporting_structure": reporting_structure,
                }.items():
                    st.session_state[k] = v
                if st.session_state["save_run_preference"]:
                    save_run_to_history(st.session_state["company_profile"], consolidated_pnl, consolidated_bs, consolidated_kpis, branch_summary)

                render_validation_centre(critical_items, warning_items, recommendation_items, info_items, previews, block_processing=False)
                if unmapped.empty:
                    st.success("Files loaded successfully. Validation Centre has been opened.")
                else:
                    st.warning("Files loaded, but unmapped GL rows were found. Review the Validation Centre.")
            except Exception as e:
                add_item(critical_items, "Processing", f"Files passed upload validation, but processing failed: {e}", "Review the error details, mapping file, and source files.")
                render_validation_centre(critical_items, warning_items, recommendation_items, info_items, previews, block_processing=True)
                st.exception(e)

    with st.expander("Prior Period / Restore"):
        company_name_for_history = st.session_state["company_profile"].get("Company Name", "").strip()
        if not company_name_for_history:
            st.warning("Please save Company Profile first.")
        else:
            saved_runs = list_saved_company_runs(company_name_for_history)
            if saved_runs:
                selected_run = st.selectbox("Select Saved Run", saved_runs)
                if st.button("Restore Selected Run", use_container_width=True):
                    restored = restore_run_from_history(company_name_for_history, selected_run)
                    st.session_state["prior_pnl"] = restored.get("prior_pnl")
                    st.session_state["prior_bs"] = restored.get("prior_bs")
                    st.session_state["prior_kpis"] = restored.get("prior_kpis")
                    st.success(f"Restored: {selected_run}")
            else:
                st.info("No saved history found for this company.")
        c1, c2 = st.columns(2)
        with c1:
            prior_pnl_file = st.file_uploader("Prior Period P&L (Optional)", type=["xlsx"])
        with c2:
            prior_bs_file = st.file_uploader("Prior Period Balance Sheet (Optional)", type=["xlsx"])
            prior_kpi_file = st.file_uploader("Prior Period KPI Pack (Optional)", type=["xlsx"])
        if st.button("Load Prior Period Inputs", use_container_width=True):
            try:
                loaded_any = False
                if prior_pnl_file is not None:
                    st.session_state["prior_pnl"] = normalize_uploaded_pnl(pd.read_excel(prior_pnl_file), "Prior Period P&L")
                    loaded_any = True
                if prior_bs_file is not None:
                    st.session_state["prior_bs"] = normalize_uploaded_bs(pd.read_excel(prior_bs_file), "Prior Period Balance Sheet")
                    loaded_any = True
                if prior_kpi_file is not None:
                    pk = clean_columns(pd.read_excel(prior_kpi_file))
                    pk.rename(columns={"Kpi": "KPI", "Display value": "Display Value"}, inplace=True)
                    validate_required_columns(pk, ["KPI", "Value"], "Prior Period KPI Pack")
                    st.session_state["prior_kpis"] = pk
                    loaded_any = True
                st.success("Prior period data loaded successfully.") if loaded_any else st.info("No prior period file uploaded.")
            except Exception as e:
                st.error(f"Error loading prior period data: {e}")

    if st.session_state.get("gl") is not None:
        with st.expander("Validation Summary"):
            report = st.session_state.get("last_validation_report") or {}
            if report:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Readiness Score", f"{report.get('score', 100)}/100")
                c2.metric("Critical", len(report.get("critical", [])))
                c3.metric("Warnings", len(report.get("warnings", [])))
                c4.metric("Recommendations", len(report.get("recommendations", [])))
                if st.button("Open Validation Centre Again", use_container_width=True):
                    render_validation_centre(report.get("critical", []), report.get("warnings", []), report.get("recommendations", []), report.get("info", []), {}, block_processing=False)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("GL Rows", len(st.session_state["gl"]))
            m2.metric("Mapped Rows", len(st.session_state["mapped"]))
            m3.metric("Unmapped Rows", len(st.session_state["unmapped"]))
            m4.metric("Reporting Units", len(st.session_state["detected_branches"] or []))
            unmapped = st.session_state.get("unmapped")
            if unmapped is not None and not unmapped.empty:
                cols_to_show = [c for c in ["Account code", "Account Name", "Description", "Branch", "Debit", "Credit", "Net"] if c in unmapped.columns]
                st.dataframe(style_dataframe(unmapped[cols_to_show]), use_container_width=True)

    with st.expander("Required Columns Guide"):
        g1, g2 = st.columns(2)
        with g1:
            show_required_columns("Current GL Report", ["Account code", "Debit", "Credit"], ["Branch / Business Unit", "Net", "Date", "Period", "Description", "Account Name"])
            show_required_columns("COA Mapping", ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], ["Account Name", "Sign Convention", "Display Order"])
            show_required_columns("KPI Master", ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"], [])
            show_required_columns("Latest Previous Balance Sheet", ["Reporting Group", "Reporting Subgroup", "Balance"], [])
            show_required_columns("Budget Data", ["Month", "Reporting Group", "Amount"], ["Branch / Business Unit"])
            show_required_columns("Forecast P&L", ["Reporting Group", "Reporting Subgroup", "Report Value"], ["Period"])
        with g2:
            show_required_columns("Forecast Balance Sheet", ["Reporting Group", "Reporting Subgroup", "Balance"], [])
            show_required_columns("Previous Year P&L", ["Reporting Group", "Reporting Subgroup", "Report Value"], ["Period"])
            show_required_columns("AR Ageing", ["Party Name", "Outstanding Amount"], ["Document Number", "Document Date", "Due Date", "Branch", "Age Bucket"])
            show_required_columns("AP Ageing", ["Party Name", "Outstanding Amount"], ["Document Number", "Document Date", "Due Date", "Branch", "Age Bucket"])
            show_required_columns("Industry Benchmark File", ["Metric", "Benchmark Value"], [])
            show_required_columns("Prior Period KPI Pack", ["KPI", "Value"], ["Display Value", "Output Type"])

    with st.expander("Download Sample Templates", expanded=True):
        st.info("Download a template, replace the sample rows with your own data, and upload the same file back into the app.")
        cols = st.columns(3)
        for idx, (name, df) in enumerate(get_sample_templates().items()):
            with cols[idx % 3]:
                st.download_button(
                    label=f"Download {name}",
                    data=make_sample_template_bytes(df),
                    file_name=f"{name.lower().replace(' ', '_').replace('&', 'and')}_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"tpl_{name}",
                )

elif selected_page == "📊 Dashboard":
    profile = st.session_state.get("company_profile", {}) or {}
    report = st.session_state.get("last_validation_report") or {}
    score = report.get("score", 100 if isinstance(st.session_state.get("mapped"), pd.DataFrame) and not st.session_state.get("mapped").empty else 0)
    profile_done = bool(profile.get("Company Name"))
    data_loaded = st.session_state.get("mapped") is not None
    validation_ok = bool(st.session_state.get("validation_passed")) if data_loaded else False
    reports_ready = st.session_state.get("consolidated_pnl") is not None
    insights_ready = bool(st.session_state.get("ai_commentary"))
    render_v1_dashboard_command_center(
        profile,
        score,
        profile_done,
        data_loaded,
        validation_ok,
        reports_ready,
        insights_ready,
        st.session_state,
    )

elif selected_page == "📉 Business Analytics":
    render_business_analytics_page()

elif selected_page == "🌐 Market Research":
    render_market_research_page()

elif selected_page == "📈 Reports":
    st.subheader("Reports")
    st.caption(f"Reporting period: {get_report_period_label(st.session_state.get('company_profile', {}))}")
    sub_pnl, sub_bs, sub_kpi, sub_trends, sub_variance = st.tabs(["P&L", "Balance Sheet", "KPIs", "Trends", "Variance"])
    with sub_pnl:
        if st.session_state["consolidated_pnl"] is None:
            st.info("No P&L available yet.")
        else:
            st.markdown("### Consolidated P&L Summary")
            st.dataframe(style_dataframe(st.session_state["consolidated_pnl"]), use_container_width=True)
            if st.session_state.get("consolidated_pnl_detail") is not None and not st.session_state["consolidated_pnl_detail"].empty:
                st.markdown("### P&L Detail by GL Account")
                st.info("This view maps by exact Account code, so similar accounts such as Freight Domestic and Freight International remain separate.")
                st.dataframe(style_dataframe(st.session_state["consolidated_pnl_detail"]), use_container_width=True)
            if st.session_state["branch_outputs"]:
                st.markdown("### Branch P&L")
                for branch, reports in st.session_state["branch_outputs"].items():
                    with st.expander(str(branch)):
                        st.markdown("**Summary**")
                        st.dataframe(style_dataframe(reports["pnl"]), use_container_width=True)
                        if reports.get("pnl_detail") is not None and not reports["pnl_detail"].empty:
                            st.markdown("**Detail by GL Account**")
                            st.dataframe(style_dataframe(reports["pnl_detail"]), use_container_width=True)
            if st.session_state["forecast_pnl"] is not None:
                st.markdown("### Forecast P&L")
                st.dataframe(style_dataframe(st.session_state["forecast_pnl"]), use_container_width=True)
            if st.session_state["previous_year_pnl"] is not None:
                st.markdown("### Previous Year P&L")
                st.dataframe(style_dataframe(st.session_state["previous_year_pnl"]), use_container_width=True)
    with sub_bs:
        if st.session_state["consolidated_bs"] is None or st.session_state["consolidated_bs"].empty:
            st.info("No Balance Sheet available yet.")
        else:
            if st.session_state["bs_disclaimer"]:
                st.warning(st.session_state["bs_disclaimer"])
            st.dataframe(style_dataframe(st.session_state["consolidated_bs"]), use_container_width=True)
            if st.session_state.get("consolidated_bs_detail") is not None and not st.session_state["consolidated_bs_detail"].empty:
                st.markdown("### Balance Sheet Detail by GL Account")
                st.dataframe(style_dataframe(st.session_state["consolidated_bs_detail"]), use_container_width=True)
        if st.session_state["forecast_bs"] is not None:
            st.markdown("### Forecast Balance Sheet")
            st.dataframe(style_dataframe(st.session_state["forecast_bs"]), use_container_width=True)
    with sub_kpi:
        if st.session_state["consolidated_kpis"] is None:
            st.info("No KPI master uploaded.")
        else:
            st.markdown("### Consolidated KPIs")
            st.dataframe(style_dataframe(st.session_state["consolidated_kpis"][["KPI", "Display Value"]]), use_container_width=True)
            if st.session_state["branch_summary"] is not None and not st.session_state["branch_summary"].empty:
                st.markdown("### Branch KPI Summary")
                st.dataframe(style_dataframe(st.session_state["branch_summary"]), use_container_width=True)
    with sub_trends:
        monthly_actuals = st.session_state.get("monthly_actuals")
        monthly_branch_actuals = st.session_state.get("monthly_branch_actuals")
        if monthly_actuals is None or monthly_actuals.empty:
            st.info("No monthly trend data available. Upload GL with a valid Date column.")
        else:
            for group, title in [("revenue", "Revenue Trend"), ("gross profit", "Gross Profit Trend"), ("operating profit", "Operating Profit Trend")]:
                temp = monthly_actuals[monthly_actuals["Reporting Group"].astype(str).str.strip().str.lower() == group]
                if not temp.empty:
                    st.markdown(f"### {title}")
                    st.line_chart(temp.set_index("Month")[["Amount"]])
            if monthly_branch_actuals is not None and not monthly_branch_actuals.empty:
                st.markdown("### Branch Revenue Trend")
                st.line_chart(monthly_branch_actuals.pivot(index="Month", columns="Branch", values="Amount").fillna(0))
            st.markdown("### Monthly Trend Data")
            st.dataframe(style_dataframe(monthly_actuals), use_container_width=True)
    with sub_variance:
        if st.session_state["budget_compare"] is not None and not st.session_state["budget_compare"].empty:
            st.markdown("### Budget vs Actual")
            st.dataframe(style_dataframe(st.session_state["budget_summary"]), use_container_width=True)
            st.dataframe(style_dataframe(st.session_state["budget_compare"]), use_container_width=True)
        else:
            st.info("No budget data uploaded.")
        if st.session_state["forecast_pnl_compare"] is not None and not st.session_state["forecast_pnl_compare"].empty:
            st.markdown("### Actual vs Forecast P&L")
            st.dataframe(style_dataframe(st.session_state["forecast_pnl_compare"]), use_container_width=True)
        else:
            st.info("No forecast P&L uploaded.")
        if st.session_state["previous_year_pnl_compare"] is not None and not st.session_state["previous_year_pnl_compare"].empty:
            st.markdown("### Actual vs Previous Year P&L")
            st.dataframe(style_dataframe(st.session_state["previous_year_pnl_compare"]), use_container_width=True)
        else:
            st.info("No previous year P&L uploaded.")
        if st.session_state["benchmark_compare"] is not None and not st.session_state["benchmark_compare"].empty:
            st.markdown("### Benchmark Comparison")
            st.dataframe(style_dataframe(st.session_state["benchmark_compare"]), use_container_width=True)

elif selected_page == "💰 Working Capital Centre":
    st.subheader("Working Capital Centre")
    st.caption(f"Reporting period: {get_report_period_label(st.session_state.get('company_profile', {}))}")
    wc_ar, wc_ap = st.tabs(["AR", "AP"])
    with wc_ar:
        if st.session_state["ar_summary"] is None:
            st.info("Upload AR file to view AR ageing.")
        else:
            ar = st.session_state["ar_summary"]
            x1, x2, x3 = st.columns(3)
            x1.metric("Total AR", f"{ar['total']:,.2f}")
            x2.metric("Overdue AR", f"{ar['overdue']:,.2f}")
            x3.metric("Overdue AR %", f"{ar['overdue_pct']:.2f}%")
            if not ar["by_bucket"].empty:
                st.bar_chart(ar["by_bucket"].set_index("Age Bucket")[["Outstanding Amount"]])
            st.dataframe(style_dataframe(ar["by_bucket"]), use_container_width=True)
            st.dataframe(style_dataframe(ar["by_branch"]), use_container_width=True)
            st.dataframe(style_dataframe(ar["top_parties"]), use_container_width=True)
    with wc_ap:
        if st.session_state["ap_summary"] is None:
            st.info("Upload AP file to view AP ageing.")
        else:
            ap = st.session_state["ap_summary"]
            y1, y2, y3 = st.columns(3)
            y1.metric("Total AP", f"{ap['total']:,.2f}")
            y2.metric("Overdue AP", f"{ap['overdue']:,.2f}")
            y3.metric("Overdue AP %", f"{ap['overdue_pct']:.2f}%")
            if not ap["by_bucket"].empty:
                st.bar_chart(ap["by_bucket"].set_index("Age Bucket")[["Outstanding Amount"]])
            st.dataframe(style_dataframe(ap["by_bucket"]), use_container_width=True)
            st.dataframe(style_dataframe(ap["by_branch"]), use_container_width=True)
            st.dataframe(style_dataframe(ap["top_parties"]), use_container_width=True)

elif selected_page == "🧠 Insights":
    st.subheader("Insights")
    st.caption("Insights now focuses only on performance anomalies and AI commentary. Mapping, duplicates, and upload recommendations are handled in the Validation Centre during upload.")

    insight_anom, insight_ai = st.tabs(["Anomalies", "AI Commentary"])

    with insight_anom:
        flags = st.session_state.get("anomaly_flags", [])
        if flags:
            for flag in flags:
                st.warning(flag)
        else:
            st.success("No major financial anomalies detected based on current rules.")

        logic_review = st.session_state.get("financial_logic_review")
        if logic_review is not None and not logic_review.empty:
            st.markdown("### Financial Logic Notes")
            st.info("These are calculation-level checks. Upload-format, duplicate COA, and COA mapping recommendations stay inside the Validation Centre.")
            st.dataframe(style_dataframe(logic_review), use_container_width=True)

    with insight_ai:
        if st.session_state["mapped"] is None:
            st.warning("Please upload and validate data first.")
        elif not st.session_state["validation_passed"]:
            st.error("Resolve unmapped accounts before generating AI insights.")
        else:
            if st.button("Generate AI Insights", use_container_width=True):
                with st.spinner("Analyzing financials..."):
                    st.session_state["ai_commentary"] = generate_ai_commentary(st.session_state["consolidated_pnl"], st.session_state["consolidated_kpis"], st.session_state["consolidated_bs"], st.session_state["company_profile"], anomaly_flags=st.session_state.get("anomaly_flags", []), ar_summary=st.session_state.get("ar_summary"), ap_summary=st.session_state.get("ap_summary"), budget_summary=st.session_state.get("budget_summary"), forecast_pnl_compare=st.session_state.get("forecast_pnl_compare"))
            if st.session_state["ai_commentary"]:
                st.write(st.session_state["ai_commentary"])

elif selected_page == "💬 Ask AI CFO":
    st.subheader("AI CFO Chatbot")

    has_data = st.session_state.get("mapped") is not None
    profile = st.session_state.get("company_profile", {}) or {}

    st.markdown("""
    <div class="section-card">
        <h3>💬 Ask questions before or after upload</h3>
        <p class="small-muted">
        Before upload, ask about templates, required columns, validation, forecast setup, or benchmarks.
        After upload, ask data-specific questions about P&L, KPIs, AR/AP, budget, forecast, branch performance and mapping review.
        </p>
    </div>
    """, unsafe_allow_html=True)

    status_cols = st.columns(4)
    with status_cols[0]:
        st.metric("Data Status", "Loaded" if has_data else "Not uploaded")
    with status_cols[1]:
        st.metric("Company", profile.get("Company Name", "Not set"))
    with status_cols[2]:
        st.metric("Industry", profile.get("Industry", "Not set"))
    with status_cols[3]:
        st.metric("Country", profile.get("Country", "Not set"))

    chat_mode = st.radio(
        "Chat mode",
        ["Auto", "General Help", "Data-specific CFO Analysis", "Internet & Benchmark Research"],
        horizontal=True,
        help="Auto chooses based on your question. Internet & Benchmark Research uses available external APIs and optional Tavily search if configured.",
    )

    if not has_data:
        st.info("You can ask generic questions now. Upload and validate GL + COA to unlock company-specific CFO analysis.")
    else:
        st.success("Uploaded financial data is available for data-specific CFO questions.")

    st.markdown("### Suggested questions")
    qcols = st.columns(3)
    with qcols[0]:
        if st.button("What files should I upload?", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "What files should I upload and which columns are mandatory?"
        if st.button("How should I map freight accounts?", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "How should I decide whether freight domestic and freight international should be COGS or overheads?"
    with qcols[1]:
        if st.button("Compare me with benchmarks", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "Compare the company against available country and industry benchmarks. Tell me what data is missing if needed."
        if st.button("Explain forecast uploads", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "How should I upload forecast P&L and forecast balance sheet, and how will the app compare them?"
    with qcols[2]:
        if st.button("Why is GP changing?", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "Why is gross profit or gross margin changing based on the uploaded data?"
        if st.button("What should management focus on?", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "What are the top management focus areas based on the uploaded data and validation review?"

    tool_cols = st.columns([1, 1, 3])
    with tool_cols[0]:
        if st.button("Clear Chat", use_container_width=True):
            st.session_state["ai_cfo_chat_messages"] = []
            st.rerun()
    with tool_cols[1]:
        if st.button("Fetch Benchmark Context", use_container_width=True):
            try:
                selected_country = profile.get("Country", "Australia") or "Australia"
                selected_industry = profile.get("Industry", "Other") or "Other"
                st.session_state["country_indicators"] = fetch_country_indicators(selected_country)
                st.session_state["external_benchmark_df"] = get_builtin_industry_benchmarks(selected_industry, selected_country)
                st.success("Benchmark context loaded. Ask the chatbot to compare benchmarks now.")
            except Exception as exc:
                st.error(f"Benchmark context fetch failed: {exc}")

    # Chat transcript
    st.markdown("### Conversation")
    if not st.session_state.get("ai_cfo_chat_messages"):
        with st.chat_message("assistant"):
            st.markdown(
                "Hi, I’m your AI CFO assistant. I can help with upload formats, mapping decisions, validation issues, benchmarks, forecasts, and uploaded-data analysis. What would you like to check?"
            )

    for msg in st.session_state.get("ai_cfo_chat_messages", []):
        with st.chat_message(msg.get("role", "assistant")):
            st.markdown(msg.get("content", ""))

    pending_prompt = st.session_state.pop("_ai_cfo_pending_prompt", None)
    typed_prompt = st.chat_input("Ask about uploads, benchmarks, mapping, forecasts, or your uploaded financial data...")
    user_question = pending_prompt or typed_prompt

    if user_question:
        st.session_state["ai_cfo_chat_messages"].append({"role": "user", "content": user_question})
        with st.chat_message("user"):
            st.markdown(user_question)
        with st.chat_message("assistant"):
            with st.spinner("AI CFO is thinking..."):
                answer = answer_ai_cfo_question(user_question, mode=chat_mode)
            st.markdown(answer)
        st.session_state["ai_cfo_chat_messages"].append({"role": "assistant", "content": answer})

    with st.expander("What can the chatbot use?"):
        st.write("Before upload: app guidance, template rules, validation rules, generic finance logic and available external benchmark APIs.")
        st.write("After upload: your uploaded P&L, balance sheet, KPIs, branch data, AR/AP, budget, forecast, benchmarks, validation issues and mapping review.")
        st.write("Live web search is optional and requires `TAVILY_API_KEY` in deployment secrets. Without it, the app uses FX APIs, World Bank indicators, starter benchmarks and uploaded benchmark files.")
        st.write("The chatbot does not permanently train itself and does not replace finance review.")


elif selected_page == "📤 Downloads":
    st.subheader("Downloads")
    if st.session_state["mapped"] is None:
        st.warning("Please validate and load files first.")
    elif not st.session_state["validation_passed"]:
        st.error("Resolve unmapped GL rows before downloading reports.")
    else:
        full_pack_bytes = create_excel_pack(consolidated_pnl=st.session_state["consolidated_pnl"], consolidated_bs=st.session_state["consolidated_bs"], consolidated_kpis=st.session_state["consolidated_kpis"], branch_summary=st.session_state["branch_summary"], branch_outputs=st.session_state["branch_outputs"], unmapped=st.session_state["unmapped"], executive_summary=st.session_state["executive_summary_df"], monthly_actuals=st.session_state["monthly_actuals"], monthly_branch_actuals=st.session_state["monthly_branch_actuals"], ar_df=st.session_state["ar_df"], ap_df=st.session_state["ap_df"], budget_compare=st.session_state["budget_compare"], forecast_compare=st.session_state["forecast_pnl_compare"], py_compare=st.session_state["previous_year_pnl_compare"], benchmark_compare=st.session_state["benchmark_compare"], forecast_bs=st.session_state["forecast_bs"], fx_rate_info=st.session_state.get("fx_rate_info"), country_indicators=st.session_state.get("country_indicators"), external_benchmark_df=st.session_state.get("external_benchmark_df"))
        st.download_button(label="Download Full Management Pack", data=full_pack_bytes, file_name="full_management_pack.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        if st.session_state["unmapped"] is not None and not st.session_state["unmapped"].empty:
            st.download_button(label="Download Unmapped GL", data=st.session_state["unmapped"].to_csv(index=False).encode("utf-8"), file_name="unmapped_gl.csv", mime="text/csv", use_container_width=True)