import pandas as pd
from core.excel_templates import dataframe_to_excel_bytes

def create_excel_pack(consolidated_pnl, consolidated_bs, consolidated_kpis, branch_summary, branch_outputs, unmapped, executive_summary=None, monthly_actuals=None, monthly_branch_actuals=None, ar_df=None, ap_df=None, budget_compare=None, forecast_compare=None, py_compare=None, benchmark_compare=None, forecast_bs=None, fx_rate_info=None, country_indicators=None, external_benchmark_df=None, consolidated_pnl_detail=None, consolidated_bs_detail=None, coa_mapping_review=None, financial_logic_review=None):
    df_dict = {"Executive Summary": executive_summary if executive_summary is not None else pd.DataFrame(), "Consolidated P&L": consolidated_pnl}
    if consolidated_pnl_detail is not None and not consolidated_pnl_detail.empty:
        df_dict["P&L Detail by GL"] = consolidated_pnl_detail
    if consolidated_bs is not None and not consolidated_bs.empty:
        df_dict["Consolidated BS"] = consolidated_bs
    if consolidated_bs_detail is not None and not consolidated_bs_detail.empty:
        df_dict["BS Detail by GL"] = consolidated_bs_detail
    if forecast_bs is not None and not forecast_bs.empty:
        df_dict["Forecast BS"] = forecast_bs
    if consolidated_kpis is not None:
        df_dict["Consolidated KPIs"] = consolidated_kpis
    if branch_summary is not None and not branch_summary.empty:
        df_dict["Branch Summary KPIs"] = branch_summary
    if monthly_actuals is not None and not monthly_actuals.empty:
        df_dict["Monthly Trends"] = monthly_actuals
    if monthly_branch_actuals is not None and not monthly_branch_actuals.empty:
        df_dict["Branch Monthly Trends"] = monthly_branch_actuals
    if branch_outputs:
        for branch, reports in branch_outputs.items():
            df_dict[f"{str(branch)[:20]} P&L"] = reports.get("pnl", pd.DataFrame())
            if reports.get("pnl_detail") is not None and not reports.get("pnl_detail").empty:
                df_dict[f"{str(branch)[:18]} GL Detail"] = reports.get("pnl_detail")
            if reports.get("kpis") is not None:
                df_dict[f"{str(branch)[:20]} KPIs"] = reports["kpis"]
    if unmapped is not None and not unmapped.empty:
        df_dict["Unmapped Accounts"] = unmapped
    if ar_df is not None and not ar_df.empty:
        df_dict["AR Ageing"] = ar_df
    if ap_df is not None and not ap_df.empty:
        df_dict["AP Ageing"] = ap_df
    if budget_compare is not None and not budget_compare.empty:
        df_dict["Budget vs Actual"] = budget_compare
    if forecast_compare is not None and not forecast_compare.empty:
        df_dict["Actual vs Forecast"] = forecast_compare
    if py_compare is not None and not py_compare.empty:
        df_dict["Actual vs PY"] = py_compare
    if benchmark_compare is not None and not benchmark_compare.empty:
        df_dict["Benchmark Comparison"] = benchmark_compare
    if coa_mapping_review is not None and not coa_mapping_review.empty:
        df_dict["COA Mapping Review"] = coa_mapping_review
    if financial_logic_review is not None and not financial_logic_review.empty:
        df_dict["Financial Logic Review"] = financial_logic_review
    if external_benchmark_df is not None and not external_benchmark_df.empty:
        df_dict["Benchmark Source"] = external_benchmark_df
    if country_indicators is not None and not country_indicators.empty:
        df_dict["Country Indicators"] = country_indicators
    if fx_rate_info is not None:
        df_dict["FX Rate"] = pd.DataFrame([fx_rate_info])
    return dataframe_to_excel_bytes(df_dict)

