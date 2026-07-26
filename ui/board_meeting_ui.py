from __future__ import annotations
from html import escape
import streamlit as st
from services.ratio_engine import calculate_management_ratios, format_ratio

SLIDES = [
    ("Company Snapshot", "◈"),
    ("Financial Performance", "↗"),
    ("Working Capital", "◎"),
    ("People & Productivity", "♙"),
    ("Risks & Outlook", "⚠"),
    ("Board Decisions", "✓"),
]


def _metric_value(df, name: str, currency: str = "AUD") -> str:
    if df is None or getattr(df, "empty", True):
        return "Not available"
    row = df[df["Ratio"] == name]
    if row.empty:
        return "Not available"
    r = row.iloc[0]
    return format_ratio(r["Value"], r["Unit"], currency)


def render_board_meeting_page() -> None:
    st.session_state.setdefault("board_meeting_slide", 0)
    idx = max(0, min(len(SLIDES) - 1, int(st.session_state["board_meeting_slide"])))
    profile = st.session_state.get("company_profile", {}) or {}
    currency = profile.get("Currency", "AUD")
    ratios = calculate_management_ratios(st.session_state, st.session_state.get("board_people_data", {}))
    title, icon = SLIDES[idx]
    company = escape(str(profile.get("Company Name") or "Company"))
    period = escape(str(profile.get("Report Period") or profile.get("Financial Year") or "Current period"))

    st.markdown(
        f'<div class="v6-meeting-hero"><div class="v6-meeting-kicker">BOARD MEETING MODE · {idx + 1}/{len(SLIDES)}</div><div class="v6-meeting-title">{icon} {escape(title)}</div><div class="v6-meeting-company">{company} · {period}</div></div>',
        unsafe_allow_html=True,
    )
    st.progress((idx + 1) / len(SLIDES))

    if idx == 0:
        c1, c2, c3 = st.columns(3)
        c1.metric("Current ratio", _metric_value(ratios, "Current Ratio", currency))
        c2.metric("Cash conversion cycle", _metric_value(ratios, "Cash Conversion Cycle", currency))
        c3.metric("Revenue / employee", _metric_value(ratios, "Revenue per Employee", currency))
        st.info("Use this mode during the meeting. Move through the pack with the controls below and drill into the full pages from the sidebar.")
    elif idx == 1:
        st.markdown("### Management performance view")
        pnl = st.session_state.get("consolidated_pnl")
        if pnl is not None:
            st.dataframe(pnl, use_container_width=True, hide_index=True)
        else:
            st.warning("Load financial data to populate this slide.")
    elif idx == 2:
        c1, c2, c3 = st.columns(3)
        c1.metric("DSO", _metric_value(ratios, "DSO", currency))
        c2.metric("DPO", _metric_value(ratios, "DPO", currency))
        c3.metric("DIO", _metric_value(ratios, "DIO", currency))
        st.markdown(st.session_state.get("board_ai_narrative") or st.session_state.get("ai_commentary") or "Generate the AI board narrative to show the working-capital discussion.")
    elif idx == 3:
        people = st.session_state.get("board_people_data", {}) or {}
        cols = st.columns(4)
        for col, (key, value) in zip(cols, list(people.items())[:4]):
            col.metric(key, value)
        branch = st.session_state.get("branch_people_data")
        if branch is not None and not getattr(branch, "empty", True):
            st.dataframe(branch, use_container_width=True, hide_index=True)
    elif idx == 4:
        inputs = st.session_state.get("board_report_inputs", {}) or {}
        st.markdown("### Principal risks")
        st.write(inputs.get("Top risks") or "Not yet entered.")
        st.markdown("### Management outlook")
        st.write(inputs.get("Management outlook") or "Not yet entered.")
    else:
        inputs = st.session_state.get("board_report_inputs", {}) or {}
        st.markdown("### Decisions and approvals requested")
        st.success(inputs.get("Decisions required") or "No decisions have been entered yet.")
        st.markdown("### Strategic priorities")
        st.write(inputs.get("Strategic priorities") or "Not yet entered.")

    left, middle, right = st.columns([1, 2, 1])
    if left.button("← Previous", disabled=idx == 0, use_container_width=True):
        st.session_state["board_meeting_slide"] = idx - 1
        st.rerun()
    middle.markdown(f"<div style='text-align:center;color:#b8c6db;padding:.65rem'>{escape(title)}</div>", unsafe_allow_html=True)
    if right.button("Next →", disabled=idx == len(SLIDES) - 1, use_container_width=True, type="primary"):
        st.session_state["board_meeting_slide"] = idx + 1
        st.rerun()
