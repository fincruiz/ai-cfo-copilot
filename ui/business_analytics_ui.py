from __future__ import annotations

from html import escape
from typing import Any

import altair as alt
import pandas as pd
import streamlit as st

from services.research_service import (
    research_company_environment,
    research_pack_to_context,
    research_sources_dataframe,
    search_web,
    tavily_is_configured,
)


def _df(value: Any) -> pd.DataFrame:
    return value if isinstance(value, pd.DataFrame) else pd.DataFrame()


def _safe_number(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _kpi_map(kpis: pd.DataFrame) -> dict[str, float]:
    if kpis.empty or "KPI" not in kpis.columns or "Value" not in kpis.columns:
        return {}
    return {str(row["KPI"]): _safe_number(row["Value"]) for _, row in kpis.iterrows()}


def _metric_value(mapping: dict[str, float], names: list[str]) -> float:
    lowered = {str(k).lower(): v for k, v in mapping.items()}
    for name in names:
        if name.lower() in lowered:
            return lowered[name.lower()]
    for key, value in lowered.items():
        if any(name.lower() in key for name in names):
            return value
    return 0.0


def _money(value: float, currency: str) -> str:
    symbol = {"AUD": "A$", "USD": "$", "INR": "₹", "GBP": "£", "CAD": "C$", "NZD": "NZ$"}.get(currency, currency + " ")
    magnitude = abs(value)
    if magnitude >= 1_000_000:
        return f"{symbol}{value / 1_000_000:,.2f}M"
    if magnitude >= 1_000:
        return f"{symbol}{value / 1_000:,.1f}K"
    return f"{symbol}{value:,.2f}"


def _monthly_profit_bridge(monthly: pd.DataFrame) -> pd.DataFrame:
    if monthly.empty or not {"Month", "Reporting Group", "Amount"}.issubset(monthly.columns):
        return pd.DataFrame()
    work = monthly.copy()
    group = work["Reporting Group"].astype(str).str.lower()
    work["Category"] = "Other"
    work.loc[group.str.contains("revenue|sales|income", na=False), "Category"] = "Revenue"
    work.loc[group.str.contains("cost of sales|cost of goods|cogs|direct cost", na=False), "Category"] = "COGS"
    work.loc[group.str.contains("operating expense|overhead|opex|administrative", na=False), "Category"] = "Opex"
    pivot = work.groupby(["Month", "Category"], as_index=False)["Amount"].sum().pivot(index="Month", columns="Category", values="Amount").fillna(0)
    for col in ["Revenue", "COGS", "Opex", "Other"]:
        if col not in pivot.columns:
            pivot[col] = 0.0
    pivot["Gross Profit"] = pivot["Revenue"] - pivot["COGS"]
    pivot["Net Profit"] = pivot["Gross Profit"] - pivot["Opex"] + pivot["Other"]
    return pivot.reset_index().sort_values("Month")


def _chart_card(title: str, subtitle: str) -> None:
    st.markdown(
        f'<div class="ba-section-head"><div><b>{escape(title)}</b><span>{escape(subtitle)}</span></div></div>',
        unsafe_allow_html=True,
    )


def apply_business_analytics_style() -> None:
    st.markdown(
        """
        <style>
        .ba-hero{padding:1.25rem 1.35rem;border-radius:24px;background:radial-gradient(circle at 85% 20%,rgba(14,165,233,.20),transparent 28%),linear-gradient(135deg,#07111f,#101a31 55%,#111827);border:1px solid rgba(96,165,250,.22);box-shadow:0 22px 65px rgba(2,6,23,.30);margin:.35rem 0 1rem;animation:baEnter .45s cubic-bezier(.2,.8,.2,1) both}.ba-hero small{color:#67e8f9;font-size:.68rem;font-weight:900;letter-spacing:.12em}.ba-hero h1{color:#fff!important;font-size:2.15rem!important;margin:.25rem 0 .35rem!important;letter-spacing:-.045em}.ba-hero p{color:#cbd5e1;max-width:820px;margin:0}.ba-kpis{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:.72rem;margin-bottom:1rem}.ba-kpi{padding:.9rem;border-radius:17px;background:linear-gradient(180deg,rgba(17,24,39,.98),rgba(15,23,42,.95));border:1px solid rgba(148,163,184,.20);transition:transform .18s ease,border-color .18s ease;animation:baRise .45s both}.ba-kpi:hover{transform:translateY(-4px);border-color:rgba(34,211,238,.46)}.ba-kpi span{display:block;color:#94a3b8;font-size:.7rem;font-weight:850;text-transform:uppercase;letter-spacing:.07em}.ba-kpi b{display:block;color:#fff;font-size:1.25rem;margin-top:.25rem}.ba-kpi em{display:block;color:#67e8f9;font-style:normal;font-size:.68rem;margin-top:.22rem}.ba-section-head{margin:.35rem 0 .45rem}.ba-section-head b{display:block;color:#fff;font-size:1.02rem}.ba-section-head span{display:block;color:#94a3b8;font-size:.75rem;margin-top:.08rem}.research-source{padding:.75rem .8rem;border-radius:14px;background:rgba(15,23,42,.75);border:1px solid rgba(148,163,184,.18);margin:.45rem 0}.research-source a{color:#7dd3fc!important;font-weight:850;text-decoration:none}.research-source p{color:#cbd5e1;font-size:.78rem;margin:.28rem 0 0}.research-warning{padding:.7rem .8rem;border-radius:14px;background:rgba(120,53,15,.32);border:1px solid rgba(251,191,36,.28);color:#fde68a;font-size:.76rem;margin:.4rem 0 .8rem}@keyframes baEnter{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}@keyframes baRise{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}}@media(max-width:950px){.ba-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}@media(max-width:560px){.ba-kpis{grid-template-columns:1fr}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_business_analytics_page() -> None:
    apply_business_analytics_style()
    profile = st.session_state.get("company_profile", {}) or {}
    currency = str(profile.get("Currency") or "AUD")
    monthly = _df(st.session_state.get("monthly_actuals"))
    monthly_branch = _df(st.session_state.get("monthly_branch_actuals"))
    kpis = _df(st.session_state.get("consolidated_kpis"))
    budget_compare = _df(st.session_state.get("budget_compare"))
    benchmark_compare = _df(st.session_state.get("benchmark_compare"))
    executive = _df(st.session_state.get("executive_summary_df"))
    ar_summary = st.session_state.get("ar_summary") or {}
    ap_summary = st.session_state.get("ap_summary") or {}

    st.markdown(
        """
        <div class="ba-hero"><small>BUSINESS ANALYTICS</small><h1>Performance, drivers and external context in one view.</h1><p>Move beyond financial statements with trend analysis, branch comparisons, working-capital signals, budget variance and source-backed market research.</p></div>
        """,
        unsafe_allow_html=True,
    )

    if monthly.empty and kpis.empty:
        st.info("Upload and validate GL + COA to activate business analytics, or open Demo Mode to see the complete page.")

    kpi_values = _kpi_map(kpis)
    revenue = _metric_value(kpi_values, ["Revenue", "Sales"])
    gross_margin = _metric_value(kpi_values, ["Gross Margin %", "Gross Margin"])
    operating_margin = _metric_value(kpi_values, ["Operating Margin %", "Operating Margin"])
    net_profit = _metric_value(kpi_values, ["Net Profit", "Profit After Tax"])
    ar_overdue = _safe_number(ar_summary.get("overdue_pct"))

    st.markdown(
        f"""
        <div class="ba-kpis">
          <div class="ba-kpi"><span>Revenue</span><b>{escape(_money(revenue, currency))}</b><em>Current reporting pack</em></div>
          <div class="ba-kpi"><span>Gross margin</span><b>{gross_margin:,.1f}%</b><em>Profitability quality</em></div>
          <div class="ba-kpi"><span>Operating margin</span><b>{operating_margin:,.1f}%</b><em>Operating efficiency</em></div>
          <div class="ba-kpi"><span>Net profit</span><b>{escape(_money(net_profit, currency))}</b><em>Bottom-line outcome</em></div>
          <div class="ba-kpi"><span>AR overdue</span><b>{ar_overdue:,.1f}%</b><em>Collection exposure</em></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    performance_tab, branch_tab, working_capital_tab, benchmark_tab, research_tab = st.tabs(
        ["Performance", "Branch & Segment", "Working Capital", "Benchmarks", "Market Research"]
    )

    with performance_tab:
        bridge = _monthly_profit_bridge(monthly)
        c1, c2 = st.columns(2)
        with c1:
            _chart_card("Revenue and profit trend", "Monthly revenue, gross profit and net profit")
            if not bridge.empty:
                long = bridge.melt(id_vars="Month", value_vars=["Revenue", "Gross Profit", "Net Profit"], var_name="Metric", value_name="Amount")
                chart = (
                    alt.Chart(long)
                    .mark_line(point=True, strokeWidth=3)
                    .encode(
                        x=alt.X("Month:N", sort=list(bridge["Month"]), title=None),
                        y=alt.Y("Amount:Q", title=f"Amount ({currency})"),
                        color=alt.Color("Metric:N", legend=alt.Legend(orient="bottom")),
                        tooltip=["Month:N", "Metric:N", alt.Tooltip("Amount:Q", format=",.2f")],
                    )
                    .properties(height=320)
                    .interactive()
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.info("Monthly dates are required for trend charts.")
        with c2:
            _chart_card("Cost structure", "Revenue, COGS and operating expense by month")
            if not bridge.empty:
                mix = bridge.melt(id_vars="Month", value_vars=["COGS", "Opex"], var_name="Cost Type", value_name="Amount")
                chart = (
                    alt.Chart(mix)
                    .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                    .encode(
                        x=alt.X("Month:N", sort=list(bridge["Month"]), title=None),
                        y=alt.Y("Amount:Q", title=f"Amount ({currency})"),
                        color=alt.Color("Cost Type:N", legend=alt.Legend(orient="bottom")),
                        tooltip=["Month:N", "Cost Type:N", alt.Tooltip("Amount:Q", format=",.2f")],
                    )
                    .properties(height=320)
                )
                st.altair_chart(chart, use_container_width=True)

        if not budget_compare.empty and {"Reporting Group", "Variance"}.issubset(budget_compare.columns):
            _chart_card("Budget variance", "Positive and negative variance by reporting group")
            variance = budget_compare.groupby("Reporting Group", as_index=False)["Variance"].sum().sort_values("Variance")
            chart = (
                alt.Chart(variance)
                .mark_bar(cornerRadiusEnd=5)
                .encode(
                    y=alt.Y("Reporting Group:N", sort="-x", title=None),
                    x=alt.X("Variance:Q", title=f"Variance ({currency})"),
                    color=alt.condition(alt.datum.Variance >= 0, alt.value("#22c55e"), alt.value("#ef4444")),
                    tooltip=["Reporting Group:N", alt.Tooltip("Variance:Q", format=",.2f")],
                )
                .properties(height=max(220, 42 * len(variance)))
            )
            st.altair_chart(chart, use_container_width=True)

    with branch_tab:
        if monthly_branch.empty:
            st.info("Branch analytics becomes available when the GL contains branch or business-unit data.")
        else:
            _chart_card("Branch revenue trend", "Compare revenue momentum across business units")
            chart = (
                alt.Chart(monthly_branch)
                .mark_line(point=True, strokeWidth=3)
                .encode(
                    x=alt.X("Month:N", title=None),
                    y=alt.Y("Amount:Q", title=f"Revenue ({currency})"),
                    color=alt.Color("Branch:N", legend=alt.Legend(orient="bottom")),
                    tooltip=["Month:N", "Branch:N", alt.Tooltip("Amount:Q", format=",.2f")],
                )
                .properties(height=350)
                .interactive()
            )
            st.altair_chart(chart, use_container_width=True)
            totals = monthly_branch.groupby("Branch", as_index=False)["Amount"].sum().sort_values("Amount", ascending=False)
            _chart_card("Revenue contribution", "Total contribution by branch for the selected period")
            chart2 = (
                alt.Chart(totals)
                .mark_bar(cornerRadiusEnd=6)
                .encode(
                    y=alt.Y("Branch:N", sort="-x", title=None),
                    x=alt.X("Amount:Q", title=f"Revenue ({currency})"),
                    tooltip=["Branch:N", alt.Tooltip("Amount:Q", format=",.2f")],
                )
                .properties(height=max(220, 52 * len(totals)))
            )
            st.altair_chart(chart2, use_container_width=True)

    with working_capital_tab:
        c1, c2 = st.columns(2)
        for column, title, summary in [(c1, "Accounts receivable ageing", ar_summary), (c2, "Accounts payable ageing", ap_summary)]:
            with column:
                _chart_card(title, "Outstanding balance by ageing bucket")
                ageing = _df(summary.get("by_bucket") if summary else None)
                if not ageing.empty:
                    chart = (
                        alt.Chart(ageing)
                        .mark_arc(innerRadius=58, outerRadius=105)
                        .encode(
                            theta=alt.Theta("Outstanding Amount:Q"),
                            color=alt.Color("Age Bucket:N", legend=alt.Legend(orient="bottom")),
                            tooltip=["Age Bucket:N", alt.Tooltip("Outstanding Amount:Q", format=",.2f")],
                        )
                        .properties(height=320)
                    )
                    st.altair_chart(chart, use_container_width=True)
                else:
                    st.info("Upload ageing data to activate this chart.")
        top_ar = _df(ar_summary.get("top_parties") if ar_summary else None)
        if not top_ar.empty:
            _chart_card("Largest customer exposures", "Top accounts by outstanding receivable")
            chart = (
                alt.Chart(top_ar.head(10))
                .mark_bar(cornerRadiusEnd=5)
                .encode(
                    y=alt.Y("Party Name:N", sort="-x", title=None),
                    x=alt.X("Outstanding Amount:Q", title=f"Outstanding ({currency})"),
                    tooltip=["Party Name:N", alt.Tooltip("Outstanding Amount:Q", format=",.2f")],
                )
                .properties(height=340)
            )
            st.altair_chart(chart, use_container_width=True)

    with benchmark_tab:
        st.markdown('<div class="research-warning">Benchmarks must be treated as directional until the source, peer group, accounting definitions and period are verified.</div>', unsafe_allow_html=True)
        if benchmark_compare.empty:
            st.info("Load a benchmark file, starter benchmark set, or refresh market research to compare performance.")
        else:
            display = benchmark_compare.copy()
            if {"Metric", "Current Value", "Benchmark Value"}.issubset(display.columns):
                long = display.melt(id_vars="Metric", value_vars=["Current Value", "Benchmark Value"], var_name="Series", value_name="Value")
                chart = (
                    alt.Chart(long)
                    .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                    .encode(
                        x=alt.X("Metric:N", title=None, axis=alt.Axis(labelAngle=-25)),
                        xOffset="Series:N",
                        y=alt.Y("Value:Q", title="Value"),
                        color=alt.Color("Series:N", legend=alt.Legend(orient="bottom")),
                        tooltip=["Metric:N", "Series:N", alt.Tooltip("Value:Q", format=",.2f")],
                    )
                    .properties(height=360)
                )
                st.altair_chart(chart, use_container_width=True)
            st.dataframe(display, use_container_width=True, hide_index=True)
        if not executive.empty:
            _chart_card("Financial health scorecard", "Management status across profitability, variance and liquidity")
            st.dataframe(executive, use_container_width=True, hide_index=True)

    with research_tab:
        configured = tavily_is_configured()
        if not configured:
            st.warning("Live internet research is disabled. Add TAVILY_API_KEY to Streamlit secrets to activate source-backed industry research.")
        left, right = st.columns([1.1, .9])
        with left:
            st.markdown("#### Industry and country research")
            depth = st.selectbox("Research depth", ["basic", "advanced"], help="Advanced research is more detailed and uses more Tavily credits.")
            if st.button("Refresh company research", type="primary", use_container_width=True, disabled=not configured):
                with st.status("Researching the external environment...", expanded=True) as status:
                    st.write("Searching industry outlook and margin pressures")
                    st.write("Searching financial and working-capital benchmarks")
                    st.write("Refreshing country economic indicators")
                    st.session_state["external_research_pack"] = research_company_environment(profile, search_depth=depth)
                    status.update(label="Research snapshot ready", state="complete", expanded=False)
                st.rerun()
        with right:
            st.markdown("#### Custom web research")
            custom_query = st.text_input("Question", placeholder="Example: Australian manufacturing margin outlook 2026")
            if st.button("Search the web", use_container_width=True, disabled=not configured or not custom_query.strip()):
                with st.spinner("Searching trusted public sources..."):
                    st.session_state["custom_research_result"] = search_web(custom_query.strip(), max_results=7, search_depth="advanced")
                st.rerun()

        pack = st.session_state.get("external_research_pack")
        if pack:
            st.caption(f"Research snapshot: {pack.get('retrieved_at', '')}")
            for section in pack.get("sections", []):
                with st.expander(section.get("label", "Research"), expanded=False):
                    if section.get("answer"):
                        st.write(section["answer"])
                    for source in section.get("results", [])[:5]:
                        st.markdown(
                            f'<div class="research-source"><a href="{escape(source.get("url", ""))}" target="_blank">{escape(source.get("title", "Source"))}</a><p>{escape(source.get("content", "")[:520])}</p></div>',
                            unsafe_allow_html=True,
                        )
            sources = research_sources_dataframe(pack)
            if not sources.empty:
                st.download_button(
                    "Download research source register",
                    sources.to_csv(index=False).encode("utf-8"),
                    file_name="external_research_sources.csv",
                    mime="text/csv",
                )

        custom = st.session_state.get("custom_research_result")
        if custom:
            st.markdown("#### Custom research result")
            if custom.get("answer"):
                st.write(custom["answer"])
            if custom.get("error"):
                st.error(custom["error"])
            for source in custom.get("results", []):
                st.markdown(
                    f'<div class="research-source"><a href="{escape(source.get("url", ""))}" target="_blank">{escape(source.get("title", "Source"))}</a><p>{escape(source.get("content", "")[:520])}</p></div>',
                    unsafe_allow_html=True,
                )

        with st.expander("AI commentary research context", expanded=False):
            st.text_area("Context supplied to AI CFO", research_pack_to_context(pack), height=260, disabled=True)


def render_market_research_page() -> None:
    """Standalone, discoverable market-research workspace."""
    apply_business_analytics_style()
    profile = st.session_state.get("company_profile", {}) or {}
    configured = tavily_is_configured()

    st.markdown(
        """
        <div class="ba-hero"><small>MARKET RESEARCH</small><h1>Current industry intelligence for CFO decisions.</h1><p>Search live public sources, compare external conditions with company performance, and feed source-backed context into AI CFO commentary.</p></div>
        """,
        unsafe_allow_html=True,
    )

    if not configured:
        st.warning("Live internet research is disabled. Add TAVILY_API_KEY to .streamlit/secrets.toml or Streamlit Cloud Secrets.")
    else:
        st.success("Tavily is connected. Research results can be used by the AI CFO.")

    company = str(profile.get("Company Name") or "Current company")
    industry = str(profile.get("Industry") or "the selected industry")
    country = str(profile.get("Country") or "the selected country")
    st.caption(f"Research workspace: {company} · {industry} · {country}")

    scan_tab, custom_tab, sources_tab, ai_tab = st.tabs(
        ["Executive Research Scan", "Custom Search", "Source Register", "AI Context"]
    )

    with scan_tab:
        c1, c2 = st.columns([1.15, .85])
        with c1:
            st.markdown("#### Refresh the external environment")
            depth = st.selectbox(
                "Research depth",
                ["basic", "advanced"],
                index=1,
                key="standalone_research_depth",
                help="Advanced search provides richer results and uses more Tavily credits.",
            )
            st.write("The scan covers industry outlook, margin pressure, working-capital benchmarks, economic conditions, regulation and market risks.")
            if st.button(
                "Run executive research scan",
                type="primary",
                use_container_width=True,
                disabled=not configured,
                key="standalone_refresh_company_research",
            ):
                with st.status("Building the external intelligence pack...", expanded=True) as status:
                    st.write("Searching industry outlook and demand conditions")
                    st.write("Searching margins, costs and working-capital benchmarks")
                    st.write("Searching economic, regulatory and supply-chain developments")
                    st.session_state["external_research_pack"] = research_company_environment(profile, search_depth=depth)
                    status.update(label="External intelligence pack ready", state="complete", expanded=False)
                st.rerun()
        with c2:
            st.markdown("#### Suggested CFO questions")
            st.markdown(
                """
                - How is this industry performing in the selected country?  
                - What margin benchmarks should management compare against?  
                - What external risks could affect the next forecast?  
                - Are labour, freight, commodity or interest-rate pressures changing?  
                - What regulatory changes should be discussed in the board pack?
                """
            )

        pack = st.session_state.get("external_research_pack")
        if pack:
            st.caption(f"Last refreshed: {pack.get('retrieved_at', '')}")
            for section in pack.get("sections", []):
                with st.expander(section.get("label", "Research"), expanded=True):
                    if section.get("answer"):
                        st.write(section["answer"])
                    for source in section.get("results", [])[:5]:
                        st.markdown(
                            f'<div class="research-source"><a href="{escape(source.get("url", ""))}" target="_blank">{escape(source.get("title", "Source"))}</a><p>{escape(source.get("content", "")[:520])}</p></div>',
                            unsafe_allow_html=True,
                        )
        else:
            st.info("Run the executive research scan to create a source-backed external intelligence pack.")

    with custom_tab:
        st.markdown("#### Ask a targeted market question")
        custom_query = st.text_input(
            "Research question",
            placeholder="Example: Australian manufacturing gross-margin outlook and freight cost pressure in 2026",
            key="standalone_custom_research_query",
        )
        if st.button(
            "Search trusted public sources",
            use_container_width=True,
            disabled=not configured or not custom_query.strip(),
            key="standalone_custom_research_button",
        ):
            with st.spinner("Searching current public sources..."):
                st.session_state["custom_research_result"] = search_web(
                    custom_query.strip(), max_results=8, search_depth="advanced"
                )
            st.rerun()

        custom = st.session_state.get("custom_research_result")
        if custom:
            if custom.get("answer"):
                st.markdown("#### Research answer")
                st.write(custom["answer"])
            if custom.get("error"):
                st.error(custom["error"])
            for source in custom.get("results", []):
                st.markdown(
                    f'<div class="research-source"><a href="{escape(source.get("url", ""))}" target="_blank">{escape(source.get("title", "Source"))}</a><p>{escape(source.get("content", "")[:520])}</p></div>',
                    unsafe_allow_html=True,
                )

    with sources_tab:
        pack = st.session_state.get("external_research_pack")
        sources = research_sources_dataframe(pack) if pack else pd.DataFrame()
        if sources.empty:
            st.info("No source register is available yet. Run a research scan first.")
        else:
            st.dataframe(sources, use_container_width=True, hide_index=True)
            st.download_button(
                "Download source register",
                sources.to_csv(index=False).encode("utf-8"),
                file_name="external_research_sources.csv",
                mime="text/csv",
                use_container_width=True,
            )

    with ai_tab:
        pack = st.session_state.get("external_research_pack")
        context = research_pack_to_context(pack)
        st.markdown("#### Context supplied to AI CFO")
        st.caption("This research is combined with uploaded financial data when the AI CFO prepares commentary and comparisons.")
        st.text_area("Research context", context, height=360, disabled=True, label_visibility="collapsed")
