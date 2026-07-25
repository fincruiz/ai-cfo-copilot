import os
import json
from urllib.request import Request, urlopen

import pandas as pd
import streamlit as st
try:
    from openai import OpenAI
except Exception:
    OpenAI = None
from core.common import get_report_period_label
from services.external_data import fetch_country_indicators, get_builtin_industry_benchmarks
from services.research_service import (
    research_company_environment,
    research_pack_to_context,
    search_web,
    tavily_is_configured,
)


def _secret(name: str) -> str:
    value = os.getenv(name, "")
    if value:
        return value
    try:
        return str(st.secrets.get(name, ""))
    except Exception:
        return ""


def ensure_external_research_context(force: bool = False, scan_type: str = "Executive scan"):
    """Load one source-backed research snapshot for AI and commentary use.

    The snapshot is cached in session state so ordinary chatbot questions and
    commentary generation do not repeatedly consume Tavily credits.
    """
    existing = st.session_state.get("external_research_pack")
    if existing and not force:
        return existing
    profile = st.session_state.get("company_profile", {}) or {}
    if not profile.get("Industry") or not profile.get("Country"):
        return existing
    if not tavily_is_configured():
        return existing
    try:
        pack = research_company_environment(
            profile, search_depth="basic", scan_type=scan_type, prefer_authoritative=True
        )
        st.session_state["external_research_pack"] = pack
        return pack
    except Exception as exc:
        st.session_state["external_research_error"] = str(exc)
        return existing

def generate_ai_commentary(pnl_df, kpi_df, bs_df, profile, anomaly_flags=None, ar_summary=None, ap_summary=None, budget_summary=None, forecast_pnl_compare=None):
    if OpenAI is None:
        return "AI Commentary failed: openai package is not installed. Add openai to requirements.txt."
    if not _secret("OPENAI_API_KEY"):
        return "AI Commentary failed: OPENAI_API_KEY is not set in Streamlit secrets/environment."
    try:
        client = OpenAI(api_key=_secret("OPENAI_API_KEY"))
        model_name = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
        pnl_summary = pnl_df.to_string(index=False)[:3000] if pnl_df is not None and not pnl_df.empty else "No P&L data available."
        kpi_summary = kpi_df[["KPI", "Display Value"]].to_string(index=False)[:2000] if kpi_df is not None and not kpi_df.empty else "No KPI data available."
        bs_summary = bs_df.to_string(index=False)[:2000] if bs_df is not None and not bs_df.empty else "No Balance Sheet data available."
        anomaly_text = "\n".join(anomaly_flags) if anomaly_flags else "No anomaly flags detected."
        research_pack = ensure_external_research_context(force=False)
        external_context = research_pack_to_context(research_pack, max_chars=10000)
        prompt = f"""
Prepare concise CFO commentary using the internal financial data and the external research context below.
Company profile: {profile}
Anomaly flags: {anomaly_text}
P&L: {pnl_summary}
KPIs: {kpi_summary}
Balance Sheet: {bs_summary}
External industry / economic context:
{external_context}

Write: Executive Summary, Key Insights, External Context, Risks, Opportunities, Recommended Actions.
Clearly separate company facts from external context. For external claims, mention the source title or URL supplied in the context. Never present external research as audited fact.
"""
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "developer", "content": "You are a concise CFO advisor."}, {"role": "user", "content": prompt}],
            temperature=0.3,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI Commentary failed: {str(e)}"



def dataframe_context(label: str, df: pd.DataFrame | None, max_rows: int = 30, max_chars: int = 4500) -> str:
    """Convert an in-memory dataframe into compact text for AI context."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return f"{label}: Not available."
    try:
        preview = df.head(max_rows).to_string(index=False)
        return f"{label} (first {min(len(df), max_rows)} of {len(df)} rows):\n{preview[:max_chars]}"
    except Exception as exc:
        return f"{label}: Could not render dataframe context ({exc})."


def build_ai_cfo_context() -> str:
    """Build a grounded context pack from the currently uploaded/processed finance data."""
    profile = st.session_state.get("company_profile", {}) or {}
    ar_summary = st.session_state.get("ar_summary")
    ap_summary = st.session_state.get("ap_summary")
    validation_report = st.session_state.get("last_validation_report") or {}

    summary_lines = [
        f"Company Profile: {profile}",
        f"Reporting Structure: {st.session_state.get('reporting_structure')}",
        f"Validation Score: {validation_report.get('score', 'Not available')}",
        f"Validation Critical Items: {len(validation_report.get('critical', []))}",
        f"Validation Warnings: {len(validation_report.get('warnings', []))}",
        f"Validation Recommendations: {len(validation_report.get('recommendations', []))}",
    ]
    if ar_summary is not None:
        summary_lines.append(f"AR Summary: Total={ar_summary.get('total', 0):,.2f}, Overdue={ar_summary.get('overdue', 0):,.2f}, Overdue %={ar_summary.get('overdue_pct', 0):.2f}%")
    if ap_summary is not None:
        summary_lines.append(f"AP Summary: Total={ap_summary.get('total', 0):,.2f}, Overdue={ap_summary.get('overdue', 0):,.2f}, Overdue %={ap_summary.get('overdue_pct', 0):.2f}%")

    context_parts = [
        "\n".join(summary_lines),
        dataframe_context("Consolidated P&L", st.session_state.get("consolidated_pnl")),
        dataframe_context("Consolidated KPIs", st.session_state.get("consolidated_kpis")),
        dataframe_context("Consolidated Balance Sheet", st.session_state.get("consolidated_bs")),
        dataframe_context("Branch KPI Summary", st.session_state.get("branch_summary")),
        dataframe_context("Budget vs Actual Summary", st.session_state.get("budget_summary")),
        dataframe_context("Forecast P&L Comparison", st.session_state.get("forecast_pnl_compare")),
        dataframe_context("Previous Year P&L Comparison", st.session_state.get("previous_year_pnl_compare")),
        dataframe_context("Benchmark Comparison", st.session_state.get("benchmark_compare")),
        dataframe_context("COA Mapping Review", st.session_state.get("coa_mapping_review")),
        dataframe_context("Financial Logic Review", st.session_state.get("financial_logic_review")),
        dataframe_context("AR Ageing Buckets", (ar_summary or {}).get("by_bucket") if ar_summary else None),
        dataframe_context("AP Ageing Buckets", (ap_summary or {}).get("by_bucket") if ap_summary else None),
        dataframe_context("Monthly Actuals", st.session_state.get("monthly_actuals")),
        dataframe_context("Unmapped GL Rows", st.session_state.get("unmapped"), max_rows=20),
    ]
    return "\n\n---\n\n".join(context_parts)[:28000]



def generic_ai_cfo_help_context() -> str:
    """Static product guidance the chatbot can use before any upload."""
    return """
AI CFO Copilot app guidance:
- Mandatory current-period uploads: Current GL Report and COA Mapping.
- GL mandatory columns in Consolidated Only mode: Account code, Debit, Credit. Branch is optional and defaults to Consolidated.
- GL mandatory columns in Branch / Business Unit mode: Account code, Debit, Credit, Branch.
- GL optional but useful columns: Net, Date, Description.
- COA Mapping mandatory columns: Account code, Reporting Group, Reporting Subgroup, Statement.
- COA optional columns: Sign Convention, Display Order, Account Name.
- Forecast P&L upload columns: Reporting Group, Reporting Subgroup, Report Value.
- Forecast Balance Sheet upload columns: Reporting Group, Reporting Subgroup, Balance.
- Previous Year P&L upload columns: Reporting Group, Reporting Subgroup, Report Value.
- Budget upload columns: Month, Branch, Reporting Group, Amount.
- AR/AP Ageing mandatory columns: Party Name, Outstanding Amount. Optional: Document Number, Document Date, Due Date, Branch, Age Bucket.
- Benchmark upload columns: Metric, Benchmark Value.
- Recommended flow: set company profile, choose reporting structure, download templates, replace sample rows, validate and upload, review validation centre, then view dashboard/reports.
- The chatbot can answer generic upload questions before upload. After upload, it can also answer data-specific CFO questions.
- Internet/benchmark capability: the app can use fetched FX rates, World Bank country indicators, starter industry benchmark data, and optionally web search if TAVILY_API_KEY is configured by the deployment owner.
""".strip()


def fetch_tavily_search_context(query: str, max_results: int = 5) -> str:
    """Compatibility wrapper around the source-preserving research service."""
    result = search_web(query, max_results=max_results, search_depth="basic")
    if not result.get("ok"):
        return f"Web search unavailable: {result.get('error') or result.get('answer', '')}"
    lines = [f"Search answer: {result.get('answer', '')}"]
    for item in result.get("results", []):
        lines.append(f"- {item.get('title')}\n  URL: {item.get('url')}\n  Summary: {item.get('content', '')[:700]}")
    return "Web search results:\n" + "\n".join(lines)

def build_external_ai_context(question: str) -> str:
    """Build optional internet/benchmark context from configured APIs and already loaded external data."""
    profile = st.session_state.get("company_profile", {}) or {}
    country = profile.get("Country", "") or "Australia"
    industry = profile.get("Industry", "") or "Other"
    currency = profile.get("Currency", "") or "AUD"
    q = (question or "").lower()

    parts = []
    research_pack = ensure_external_research_context(force=False)
    parts.append(research_pack_to_context(research_pack, max_chars=12000))
    parts.append(dataframe_context("Loaded Country Indicators", st.session_state.get("country_indicators")))
    parts.append(dataframe_context("Loaded External Benchmark Data", st.session_state.get("external_benchmark_df")))
    fx_info = st.session_state.get("fx_rate_info")
    if fx_info:
        parts.append(f"Loaded FX Rate: {fx_info}")

    wants_external = any(word in q for word in [
        "benchmark", "industry", "country", "market", "inflation", "gdp", "forex", "fx", "exchange", "external", "internet", "web", "compare"
    ])

    if wants_external:
        try:
            country_ind = fetch_country_indicators(country)
            parts.append(dataframe_context(f"Fresh World Bank Country Indicators for {country}", country_ind))
        except Exception as exc:
            parts.append(f"World Bank country indicator fetch failed: {exc}")
        try:
            starter_bench = get_builtin_industry_benchmarks(industry, country)
            parts.append(dataframe_context(f"Starter Industry Benchmarks for {industry} / {country}", starter_bench))
        except Exception as exc:
            parts.append(f"Starter benchmark load failed: {exc}")
        try:
            if currency and currency != "AUD":
                parts.append(f"FX context note: profile currency is {currency}. Use the app's FX section for exact conversion rate before financial comparison.")
        except Exception:
            pass
        parts.append(fetch_tavily_search_context(question))

    return "\n\n---\n\n".join([p for p in parts if p])[:18000]


def fallback_chatbot_answer(question: str, has_data: bool) -> str:
    """Useful fallback when OpenAI key is not available."""
    q = (question or "").lower()
    if any(w in q for w in ["branch", "business unit", "division"]):
        return "Branch is optional. Use **Consolidated Only** if the GL has no branch/business unit column; the app will use `Consolidated`. Use **Branch / Business Unit Reporting** only when you want branch-wise P&L/KPIs and your GL has a Branch column."
    if any(w in q for w in ["template", "upload", "column", "format"]):
        return "Use the **Download Sample Templates** section. For GL, the required columns are `Account code`, `Debit`, `Credit`; `Branch` is required only for Branch / Business Unit Reporting. COA needs `Account code`, `Reporting Group`, `Reporting Subgroup`, and `Statement`."
    if any(w in q for w in ["forecast", "3 way", "three way"]):
        return "For now, upload Forecast P&L and Forecast Balance Sheet directly. A driver-based 3-way model should later generate P&L, BS and Cash Flow from assumptions like revenue growth, gross margin, DSO, DPO, inventory days, capex, debt and tax."
    if any(w in q for w in ["benchmark", "industry", "country", "forex", "fx"]):
        return "The app supports uploaded benchmark files, starter industry benchmarks, World Bank country indicators, and FX-rate fetching. For live web search, configure `TAVILY_API_KEY`; otherwise the app uses built-in/API sources only."
    if not has_data:
        return "I can answer generic questions now. For company-specific analysis, upload and validate GL + COA first, then I can review P&L, KPIs, AR/AP, budget, forecast, benchmarks and mapping warnings."
    return "I can help analyse the uploaded financial data. Ask about margin movement, revenue, branch performance, AR/AP risk, budget variance, forecast variance, benchmarks, or mapping issues."


def answer_ai_cfo_question(question: str, mode: str = "Auto") -> str:
    """Chatbot answer supporting generic pre-upload help, uploaded-data analysis, and optional external benchmark context."""
    has_data = st.session_state.get("mapped") is not None
    if OpenAI is None or not _secret("OPENAI_API_KEY"):
        return fallback_chatbot_answer(question, has_data)

    generic_context = generic_ai_cfo_help_context()
    uploaded_context = build_ai_cfo_context() if has_data else "Uploaded finance data: Not available yet. User has not validated and uploaded files."
    external_context = build_external_ai_context(question)
    prior_messages = st.session_state.get("ai_cfo_chat_messages", [])[-10:]
    chat_history = "\n".join([f"{m.get('role', 'user')}: {m.get('content', '')}" for m in prior_messages])[:7000]

    system_prompt = """
You are an AI CFO chatbot inside a finance reporting web app.
You have three jobs:
1. Before upload: answer generic questions about upload templates, columns, validation, benchmarks, forecasts, and how to use the app.
2. After upload: answer data-specific CFO questions using the uploaded financial context.
3. When benchmark/external research is requested: use the external context provided. Do not pretend live internet search is available unless web search context is present.

Rules:
- Do not invent company-specific numbers. Use uploaded-data context only for data-specific analysis.
- If required data is missing, say exactly which upload or field is needed.
- External benchmark data can be broad and may need verification; clearly label it as external/starter/API-based.
- Be practical, CFO-style, and concise. Give recommended actions.
- Do not provide legal, tax, audit, or assurance conclusions.
""".strip()

    user_prompt = f"""
Chat mode selected by user: {mode}

Generic app guidance:
{generic_context}

Uploaded data context:
{uploaded_context}

External/benchmark context:
{external_context}

Recent chat history:
{chat_history}

Current user question:
{question}
"""
    try:
        client = OpenAI(api_key=_secret("OPENAI_API_KEY"))
        model_name = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "developer", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.25,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI CFO Chat failed: {str(e)}"

