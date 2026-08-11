import pandas as pd
from core.normalizers import standardize_key_columns
from core.common import validate_required_columns
from modules.reporting import apply_sign_convention_to_gl, validate_coa_mapping_integrity, resolve_coa_duplicate_rows

def prepare_data(gl_file, mapping_file, kpi_file=None, latest_bs_file=None, allow_duplicate_coa_cleanup: bool = False, reporting_structure: str = "Consolidated Only"):
    gl = pd.read_excel(gl_file)
    coa = pd.read_excel(mapping_file)
    kpi_master = pd.read_excel(kpi_file) if kpi_file is not None else None
    latest_bs = pd.read_excel(latest_bs_file) if latest_bs_file is not None else None
    gl, coa, kpi_master, latest_bs = standardize_key_columns(gl, coa, kpi_master, latest_bs)
    branch_required = reporting_structure == "Branch / Business Unit Reporting"
    gl_required_cols = ["Account code", "Debit", "Credit"] + (["Branch"] if branch_required else [])
    validate_required_columns(gl, gl_required_cols, "Current GL Report")
    validate_required_columns(coa, ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], "COA Mapping")
    if kpi_master is not None:
        validate_required_columns(kpi_master, ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"], "KPI Master")
    if "Branch" not in gl.columns:
        gl["Branch"] = "Consolidated"
    gl["Account code"] = gl["Account code"].astype(str).str.strip()
    coa["Account code"] = coa["Account code"].astype(str).str.strip()
    validate_coa_mapping_integrity(coa, allow_duplicate_cleanup=allow_duplicate_coa_cleanup)
    if allow_duplicate_coa_cleanup:
        coa = resolve_coa_duplicate_rows(coa, keep="first")
    gl["Branch"] = gl["Branch"].astype(str).str.strip().replace({"": "Consolidated", "nan": "Consolidated"})
    gl["Debit"] = pd.to_numeric(gl["Debit"], errors="coerce").fillna(0)
    gl["Credit"] = pd.to_numeric(gl["Credit"], errors="coerce").fillna(0)
    if "Net" not in gl.columns:
        gl["Net"] = gl["Debit"] - gl["Credit"]
    else:
        gl["Net"] = pd.to_numeric(gl["Net"], errors="coerce").fillna(gl["Debit"] - gl["Credit"])
    if "Date" in gl.columns:
        gl["Date"] = pd.to_datetime(gl["Date"], errors="coerce")
    data = gl.merge(coa, on="Account code", how="left", validate="many_to_one")
    unmapped = data[data["Reporting Group"].isna()].copy()
    mapped = data[data["Reporting Group"].notna()].copy()
    if "Sign Convention" not in mapped.columns:
        mapped["Sign Convention"] = "positive"
    mapped["Report Value"] = mapped.apply(apply_sign_convention_to_gl, axis=1)
    pnl_mapped = mapped[mapped["Statement"].astype(str).str.strip().str.lower() == "income statement"].copy()
    bs_mapped = mapped[mapped["Statement"].astype(str).str.strip().str.lower() == "balance sheet"].copy()
    return gl, coa, kpi_master, latest_bs, mapped, pnl_mapped, bs_mapped, unmapped

