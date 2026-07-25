from pathlib import Path
import pandas as pd
from core.common import slugify_company_name
from core.excel_templates import dataframe_to_excel_bytes

HISTORY_ROOT = Path("history")
HISTORY_ROOT.mkdir(exist_ok=True)

def save_run_to_history(company_profile, consolidated_pnl, consolidated_bs, consolidated_kpis, branch_summary):
    company_name = company_profile.get("Company Name", "").strip()
    if not company_name:
        return
    company_slug = slugify_company_name(company_name)
    financial_year = company_profile.get("Financial Year", "unknown_year").strip().replace(" ", "_") or "unknown_year"
    reporting_period = company_profile.get("Report Period", company_profile.get("Reporting Period", "unknown_period")).strip().replace(" ", "_") or "unknown_period"
    run_folder = HISTORY_ROOT / company_slug / f"{financial_year}_{reporting_period}"
    run_folder.mkdir(parents=True, exist_ok=True)
    consolidated_pnl.to_excel(run_folder / "consolidated_pnl.xlsx", index=False)
    if consolidated_bs is not None and not consolidated_bs.empty:
        consolidated_bs.to_excel(run_folder / "consolidated_bs.xlsx", index=False)
    if consolidated_kpis is not None:
        consolidated_kpis.to_excel(run_folder / "consolidated_kpis.xlsx", index=False)
    if branch_summary is not None and not branch_summary.empty:
        branch_summary.to_excel(run_folder / "branch_summary.xlsx", index=False)


def list_saved_company_runs(company_name: str):
    company_folder = HISTORY_ROOT / slugify_company_name(company_name)
    if not company_folder.exists():
        return []
    return sorted([item.name for item in company_folder.iterdir() if item.is_dir()], reverse=True)


def restore_run_from_history(company_name: str, run_name: str):
    run_folder = HISTORY_ROOT / slugify_company_name(company_name) / run_name
    restored = {}
    if (run_folder / "consolidated_pnl.xlsx").exists():
        restored["prior_pnl"] = pd.read_excel(run_folder / "consolidated_pnl.xlsx")
    if (run_folder / "consolidated_bs.xlsx").exists():
        restored["prior_bs"] = pd.read_excel(run_folder / "consolidated_bs.xlsx")
    if (run_folder / "consolidated_kpis.xlsx").exists():
        restored["prior_kpis"] = pd.read_excel(run_folder / "consolidated_kpis.xlsx")
    return restored

