from __future__ import annotations

import json
import pandas as pd
import streamlit as st
from PIL import Image
import altair as alt

from services.ratio_engine import calculate_management_ratios, format_ratio
from services.board_report_service import create_board_report_docx, create_board_report_html


def _init_board_state() -> None:
    defaults = {
        "board_people_data": {
            "Total employees": 86, "Technical staff": 52, "Management staff": 9,
            "Apprentice staff": 8, "Sales staff": 7, "Administration staff": 10,
        },
        "board_report_inputs": {
            "Management outlook": "Revenue momentum remains positive, with management focused on margin recovery and cash conversion.",
            "Strategic priorities": "1. Improve collections and reduce overdue debt.\n2. Recover freight cost increases.\n3. Refresh the rolling forecast.\n4. Build technical capacity and apprentice retention.",
            "Top risks": "Customer concentration, labour availability, freight inflation, overdue receivables and forecast execution.",
            "Decisions required": "Approve the revised collections escalation framework and the next-quarter workforce plan.",
            "People commentary": "Technical capacity remains the key constraint. Apprentice conversion and retention are strategic priorities.",
            "Governance commentary": "No material compliance breaches were reported for the period.",
        },
        "branch_people_data": pd.DataFrame(columns=["Branch", "Employees", "Technical", "Management", "Apprentices"]),
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def _metric_cards(ratios: pd.DataFrame, names: list[str], currency: str) -> None:
    cols = st.columns(len(names))
    for col, name in zip(cols, names):
        row = ratios[ratios["Ratio"] == name]
        with col:
            if row.empty:
                st.metric(name, "N/A")
            else:
                item = row.iloc[0]
                st.metric(name, format_ratio(item["Value"], item["Unit"], currency), item["Status"])


def render_ratio_page() -> None:
    _init_board_state()
    profile = st.session_state.get("company_profile", {}) or {}
    currency = profile.get("Currency", "AUD")
    ratios = calculate_management_ratios(st.session_state, st.session_state["board_people_data"])
    st.session_state["management_ratios"] = ratios

    st.markdown("## Ratio & Productivity Centre")
    st.caption("Liquidity, working capital, leverage, profitability, returns and workforce productivity in one management view.")
    _metric_cards(ratios, ["Current Ratio", "Quick Ratio", "DSO", "DIO", "Debt to Equity"], currency)
    _metric_cards(ratios, ["Gross Margin", "Net Profit Margin", "Return on Assets", "Asset Turnover", "Revenue per Employee"], currency)

    st.markdown("### Ratio health map")
    chart_df = ratios.dropna(subset=["Value"]).copy()
    chart_df["Display"] = chart_df.apply(lambda r: format_ratio(r["Value"], r["Unit"], currency), axis=1)
    chart = alt.Chart(chart_df).mark_circle(size=380, opacity=.88).encode(
        x=alt.X("Category:N", title=None, sort=None),
        y=alt.Y("Ratio:N", title=None, sort=None),
        color=alt.Color("Tone:N", scale=alt.Scale(domain=["good", "warning", "bad", "neutral"], range=["#2dd4bf", "#fbbf24", "#fb7185", "#8aa4c8"]), legend=None),
        tooltip=["Ratio", "Category", "Display", "Status", "Interpretation"],
    ).properties(height=max(420, len(chart_df) * 28))
    st.altair_chart(chart, use_container_width=True)

    display = ratios.copy()
    display["Result"] = display.apply(lambda r: format_ratio(r["Value"], r["Unit"], currency), axis=1)
    st.dataframe(display[["Category", "Ratio", "Result", "Status", "Interpretation"]], use_container_width=True, hide_index=True)
    st.info("Ratios are decision-support indicators, not audit conclusions. Review mapping, sign conventions, period length and average-balance requirements before external use.")


def _generate_ai_board_narrative(ratios: pd.DataFrame, inputs: dict) -> str:
    profile = st.session_state.get("company_profile", {}) or {}
    currency = profile.get("Currency", "AUD")
    ratio_context = "\n".join(f"- {r['Ratio']}: {format_ratio(r['Value'], r['Unit'], currency)}; {r['Status']}" for _, r in ratios.iterrows())
    base = (
        f"{profile.get('Company Name','The company')} is reporting for {profile.get('Report Period', profile.get('Financial Year','the current period'))}. "
        f"Management's outlook is: {inputs.get('Management outlook','Not provided')}\n\n"
        f"Key ratio assessment:\n{ratio_context}\n\n"
        f"Strategic priorities: {inputs.get('Strategic priorities','Not provided')}\n"
        f"Principal risks: {inputs.get('Top risks','Not provided')}\n"
        "The board should focus discussion on cash conversion, margin quality, delivery capacity, forecast reliability and the specific approvals requested by management."
    )
    try:
        from openai import OpenAI
        key = st.secrets.get("OPENAI_API_KEY", "")
        if not key:
            return base
        client = OpenAI(api_key=key)
        prompt = f"""You are an experienced CFO preparing a board report. Write a concise, evidence-led executive narrative with sections: Performance, Cash & Working Capital, Balance Sheet & Risk, People & Capacity, Outlook, Board Attention. Never invent facts. Clearly label unavailable information.\n\nCompany inputs:\n{json.dumps(inputs, indent=2)}\n\nRatios:\n{ratio_context}\n\nExisting AI commentary:\n{st.session_state.get('ai_commentary','')}"""
        response = client.chat.completions.create(model="gpt-4o-mini", messages=[{"role":"user","content":prompt}], temperature=.2)
        return response.choices[0].message.content.strip()
    except Exception:
        return base


def render_board_report_page() -> None:
    _init_board_state()
    profile = st.session_state.get("company_profile", {}) or {}
    currency = profile.get("Currency", "AUD")
    st.markdown("## Board & Management Reporting Studio")
    st.caption("Collect non-financial context, calculate management ratios, generate an AI-supported board narrative and export a board-ready report.")

    tabs = st.tabs(["1. Company context", "2. Workforce & capacity", "3. Ratios", "4. Board narrative", "5. Export pack"])
    with tabs[0]:
        inputs = st.session_state["board_report_inputs"]
        st.markdown("### Company branding")
        st.caption("The logo saved during company creation is used automatically. You can replace it here for this workspace.")
        logo_file = st.file_uploader("Upload company logo for the board pack", type=["png", "jpg", "jpeg"], help="Use a clear PNG or JPG. The logo is placed on the Word and HTML cover pages.", key="board_company_logo_upload")
        if logo_file is not None:
            st.session_state["company_logo_bytes"] = logo_file.getvalue()
            st.session_state["company_logo_name"] = logo_file.name
            st.success("Company logo saved for this workspace and will be embedded in the report.")
        logo_bytes = st.session_state.get("company_logo_bytes")
        if logo_bytes:
            c_logo, c_info = st.columns([0.22, 0.78], vertical_alignment="center")
            with c_logo:
                st.image(logo_bytes, width=150)
            with c_info:
                st.markdown("**Branding ready**")
                st.caption(st.session_state.get("company_logo_name") or "Uploaded company logo")
                if st.button("Remove logo", key="remove_board_logo"):
                    st.session_state["company_logo_bytes"] = None
                    st.session_state["company_logo_name"] = None
                    st.rerun()
        st.divider()
        with st.form("board_context_form"):
            outlook = st.text_area("Management outlook", inputs.get("Management outlook", ""), height=110)
            priorities = st.text_area("Strategic priorities", inputs.get("Strategic priorities", ""), height=130)
            risks = st.text_area("Top risks and mitigations", inputs.get("Top risks", ""), height=130)
            decisions = st.text_area("Board decisions or approvals required", inputs.get("Decisions required", ""), height=100)
            governance = st.text_area("Governance, compliance, safety or legal matters", inputs.get("Governance commentary", ""), height=100)
            if st.form_submit_button("Save board context", type="primary", use_container_width=True):
                st.session_state["board_report_inputs"].update({"Management outlook": outlook, "Strategic priorities": priorities, "Top risks": risks, "Decisions required": decisions, "Governance commentary": governance})
                st.success("Board context saved.")

    with tabs[1]:
        people = st.session_state["board_people_data"]
        with st.form("people_form"):
            c1, c2, c3 = st.columns(3)
            total = c1.number_input("Total employees", min_value=0, value=int(people.get("Total employees", 0)))
            technical = c2.number_input("Technical staff", min_value=0, value=int(people.get("Technical staff", 0)))
            management = c3.number_input("Management staff", min_value=0, value=int(people.get("Management staff", 0)))
            c4, c5, c6 = st.columns(3)
            apprentice = c4.number_input("Apprentice staff", min_value=0, value=int(people.get("Apprentice staff", 0)))
            sales = c5.number_input("Sales staff", min_value=0, value=int(people.get("Sales staff", 0)))
            admin = c6.number_input("Administration staff", min_value=0, value=int(people.get("Administration staff", 0)))
            commentary = st.text_area("People, capacity, turnover, training and succession commentary", st.session_state["board_report_inputs"].get("People commentary", ""), height=120)
            if st.form_submit_button("Save workforce profile", type="primary", use_container_width=True):
                st.session_state["board_people_data"] = {"Total employees":total,"Technical staff":technical,"Management staff":management,"Apprentice staff":apprentice,"Sales staff":sales,"Administration staff":admin}
                st.session_state["board_report_inputs"]["People commentary"] = commentary
                st.success("Workforce profile saved and productivity ratios refreshed.")
        st.markdown("### Branch workforce allocation")
        branches = st.session_state.get("detected_branches") or []
        branch_df = st.session_state.get("branch_people_data")
        if (branch_df is None or branch_df.empty) and branches:
            branch_df = pd.DataFrame({"Branch":branches,"Employees":[0]*len(branches),"Technical":[0]*len(branches),"Management":[0]*len(branches),"Apprentices":[0]*len(branches)})
        st.session_state["branch_people_data"] = st.data_editor(branch_df, use_container_width=True, num_rows="dynamic", hide_index=True, key="branch_people_editor")

    ratios = calculate_management_ratios(st.session_state, st.session_state["board_people_data"])
    st.session_state["management_ratios"] = ratios
    with tabs[2]:
        _metric_cards(ratios, ["Current Ratio", "DSO", "DIO", "Cash Conversion Cycle", "Revenue per Employee"], currency)
        display = ratios.copy(); display["Result"] = display.apply(lambda r: format_ratio(r["Value"], r["Unit"], currency), axis=1)
        st.dataframe(display[["Category","Ratio","Result","Status","Interpretation"]], use_container_width=True, hide_index=True)

    with tabs[3]:
        if st.button("Generate / refresh AI board narrative", type="primary", use_container_width=True):
            with st.spinner("Reviewing financial performance, ratios, people inputs and board priorities..."):
                st.session_state["board_ai_narrative"] = _generate_ai_board_narrative(ratios, st.session_state["board_report_inputs"])
        narrative = st.session_state.get("board_ai_narrative") or st.session_state.get("ai_commentary") or "Generate the board narrative after completing the context and workforce sections."
        edited = st.text_area("Board executive narrative", narrative, height=430)
        if st.button("Save edited narrative", use_container_width=True):
            st.session_state["board_ai_narrative"] = edited
            st.success("Narrative saved.")

    with tabs[4]:
        completeness = sum([
            bool(st.session_state.get("consolidated_pnl") is not None),
            bool(st.session_state.get("consolidated_bs") is not None),
            bool(st.session_state.get("board_people_data", {}).get("Total employees")),
            bool(st.session_state.get("board_report_inputs", {}).get("Strategic priorities")),
            bool(st.session_state.get("board_ai_narrative") or st.session_state.get("ai_commentary")),
        ])
        st.progress(completeness / 5, text=f"Board pack readiness: {completeness}/5 core areas complete")
        html_bytes = create_board_report_html(st.session_state, ratios, st.session_state["board_report_inputs"])
        company_slug = str(profile.get("Company Name", "company")).lower().replace(" ", "_").replace("—", "-")
        try:
            docx_bytes = create_board_report_docx(st.session_state, ratios, st.session_state["board_report_inputs"])
            docx_error = None
        except ModuleNotFoundError:
            docx_bytes = None
            docx_error = "Word export needs the python-docx package. Close the app and run: uv pip install python-docx"
        except Exception as exc:
            docx_bytes = None
            docx_error = f"Word export could not be prepared: {exc}"
        c1, c2 = st.columns(2)
        with c1:
            if docx_bytes:
                st.download_button("Download branded board report (Word)", docx_bytes, f"{company_slug}_board_report.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, type="primary")
            else:
                st.button("Word export unavailable", disabled=True, use_container_width=True)
        with c2:
            st.download_button("Download branded board report (HTML)", html_bytes, f"{company_slug}_board_report.html", "text/html", use_container_width=True)
        if docx_error:
            st.warning(docx_error)
        st.warning("The pack is management-prepared and AI-assisted. Validate figures, assumptions, narrative and sensitive disclosures before board circulation.")
