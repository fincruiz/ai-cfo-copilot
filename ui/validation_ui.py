import pandas as pd
import streamlit as st
from core.excel_templates import dataframe_to_excel_bytes

def show_required_columns(title, required_cols, optional_cols=None):
    st.markdown(f"**{title}**")
    req_df = pd.DataFrame({"Column": required_cols, "Required": ["Yes"] * len(required_cols)})
    if optional_cols:
        opt_df = pd.DataFrame({"Column": optional_cols, "Required": ["Optional"] * len(optional_cols)})
        display_df = pd.concat([req_df, opt_df], ignore_index=True)
    else:
        display_df = req_df
    st.dataframe(display_df, use_container_width=True, hide_index=True)


def calculate_validation_score(critical_count: int, warning_count: int, recommendation_count: int) -> int:
    score = 100 - (critical_count * 35) - (warning_count * 8) - (recommendation_count * 3)
    return max(0, min(100, score))


def render_validation_centre(critical_items=None, warning_items=None, recommendation_items=None, info_items=None, previews=None, block_processing=False):
    """Show upload validation results in a popup-style Validation Centre."""
    critical_items = critical_items or []
    warning_items = warning_items or []
    recommendation_items = recommendation_items or []
    info_items = info_items or []
    previews = previews or {}

    score = calculate_validation_score(len(critical_items), len(warning_items), len(recommendation_items))

    def _content():
        st.markdown("### Data Validation Centre")
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Readiness Score", f"{score}/100")
        s2.metric("Critical Errors", len(critical_items))
        s3.metric("Warnings", len(warning_items))
        s4.metric("Recommendations", len(recommendation_items))

        if not critical_items and not warning_items and not recommendation_items:
            st.success("No validation errors and no recommendations. Data is ready to generate reports.")
        elif critical_items:
            st.error("Critical errors found. Please fix these before reports can be generated.")
        else:
            st.warning("Data can be processed, but review the warnings/recommendations below.")

        if critical_items:
            st.markdown("#### Critical Errors")
            st.dataframe(pd.DataFrame(critical_items), use_container_width=True, hide_index=True)
        if warning_items:
            st.markdown("#### Warnings")
            st.dataframe(pd.DataFrame(warning_items), use_container_width=True, hide_index=True)
        if recommendation_items:
            st.markdown("#### Recommendations")
            st.dataframe(pd.DataFrame(recommendation_items), use_container_width=True, hide_index=True)
        if info_items:
            st.markdown("#### Information")
            st.dataframe(pd.DataFrame(info_items), use_container_width=True, hide_index=True)

        issue_frames = []
        if critical_items:
            issue_frames.append(pd.DataFrame(critical_items).assign(Severity="Critical"))
        if warning_items:
            issue_frames.append(pd.DataFrame(warning_items).assign(Severity="Warning"))
        if recommendation_items:
            issue_frames.append(pd.DataFrame(recommendation_items).assign(Severity="Recommendation"))
        if info_items:
            issue_frames.append(pd.DataFrame(info_items).assign(Severity="Info"))
        if issue_frames:
            issue_df = pd.concat(issue_frames, ignore_index=True)
            st.download_button(
                "Download Validation Review",
                data=dataframe_to_excel_bytes({"Validation Review": issue_df}),
                file_name="validation_review.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_validation_review_popup",
            )

        if previews:
            st.markdown("#### File Previews")
            for name, df in previews.items():
                with st.expander(f"Preview: {name}"):
                    st.dataframe(df.head(5), use_container_width=True)

        if block_processing:
            st.caption("Reports are blocked until critical errors are fixed.")
        else:
            st.caption("You can proceed. Recommendations do not change your mapping automatically.")

    if hasattr(st, "dialog"):
        @st.dialog("Validation Centre")
        def _dialog():
            _content()
        _dialog()
    else:
        with st.expander("Validation Centre", expanded=True):
            _content()

