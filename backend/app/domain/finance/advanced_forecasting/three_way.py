from __future__ import annotations

from dataclasses import dataclass
from typing import Dict
import numpy as np
import pandas as pd

from .models import ForecastConfig


@dataclass
class ThreeWayForecastResult:
    profit_and_loss: pd.DataFrame
    balance_sheet: pd.DataFrame
    cash_flow: pd.DataFrame
    ratios: pd.DataFrame
    schedules: Dict[str, pd.DataFrame]
    checks: pd.DataFrame
    forecast_basis: pd.DataFrame


class ThreeWayForecastEngine:
    """Converts an operating forecast into integrated financial statements."""

    REQUIRED_COLUMNS = [
        "Period", "Revenue", "COGS", "Payroll", "Other Opex"
    ]

    def __init__(self, operating_forecast: pd.DataFrame, config: ForecastConfig):
        self.forecast = operating_forecast.copy()
        self.config = config
        self.config.validate()
        self.config.opening_balance_sheet.normalized()
        self._validate_forecast()

    def _validate_forecast(self) -> None:
        missing = [c for c in self.REQUIRED_COLUMNS if c not in self.forecast.columns]
        if missing:
            raise ValueError(f"Operating forecast missing columns: {missing}")
        self.forecast["Period"] = pd.to_datetime(self.forecast["Period"])
        self.forecast = self.forecast.sort_values("Period").reset_index(drop=True)
        for col in self.REQUIRED_COLUMNS[1:]:
            self.forecast[col] = pd.to_numeric(
                self.forecast[col], errors="coerce"
            ).fillna(0.0)

    def run(self) -> ThreeWayForecastResult:
        cfg = self.config
        d = cfg.drivers
        ob = cfg.opening_balance_sheet

        cash = ob.cash
        ar = ob.accounts_receivable
        inventory = ob.inventory
        oca = ob.other_current_assets
        gross_ppe = ob.gross_ppe
        accum_dep = ob.accumulated_depreciation
        ona = ob.other_non_current_assets

        ap = ob.accounts_payable
        accrued = ob.accrued_expenses
        ocl = ob.other_current_liabilities
        debt_current = ob.debt_current
        debt_non_current = ob.debt_non_current
        oncl = ob.other_non_current_liabilities

        share_capital = ob.share_capital
        retained_earnings = float(ob.retained_earnings or 0.0)

        capex_vintages = []
        rows = []
        wc_schedule = []
        debt_schedule = []
        ppe_schedule = []

        for _, source in self.forecast.iterrows():
            period = pd.Timestamp(source["Period"])
            days = period.days_in_month

            revenue = float(source["Revenue"])
            cogs = float(source["COGS"])
            payroll = float(source["Payroll"])
            other_opex = float(source["Other Opex"])
            gross_profit = revenue - cogs

            capex = max(0.0, revenue * d.capex_pct_revenue)
            capex_vintages.append({
                "remaining": d.useful_life_months,
                "monthly_dep": capex / max(d.useful_life_months, 1),
            })
            depreciation = 0.0
            for vintage in capex_vintages:
                if vintage["remaining"] > 0:
                    depreciation += vintage["monthly_dep"]
                    vintage["remaining"] -= 1

            opening_debt = debt_current + debt_non_current
            interest = opening_debt * d.annual_interest_rate / 12

            ebitda = gross_profit - payroll - other_opex
            ebit = ebitda - depreciation
            pbt = ebit - interest
            tax = max(0.0, pbt * d.tax_rate)
            net_income = pbt - tax
            dividend = max(0.0, net_income * d.dividend_pct_net_income)

            target_ar = revenue * d.dso_days / days
            target_inventory = cogs * d.inventory_days / days
            target_oca = revenue * d.other_current_assets_pct_revenue
            target_ap = cogs * d.dpo_days / days
            target_accrued = (
                payroll + other_opex
            ) * d.accrued_expenses_pct_opex
            target_ocl = revenue * d.other_current_liabilities_pct_revenue

            delta_ar = target_ar - ar
            delta_inventory = target_inventory - inventory
            delta_oca = target_oca - oca
            delta_ap = target_ap - ap
            delta_accrued = target_accrued - accrued
            delta_ocl = target_ocl - ocl

            cfo = (
                net_income + depreciation
                - delta_ar - delta_inventory - delta_oca
                + delta_ap + delta_accrued + delta_ocl
            )

            scheduled_repayment = min(
                d.scheduled_debt_repayment,
                opening_debt,
            )
            cash_before_funding = (
                cash + cfo - capex - scheduled_repayment - dividend
            )

            debt_draw = 0.0
            cash_sweep = 0.0
            if cash_before_funding < d.minimum_cash:
                funding_need = d.minimum_cash - cash_before_funding
                capacity = max(
                    0.0,
                    d.revolver_limit
                    - (opening_debt - scheduled_repayment),
                )
                debt_draw = min(funding_need, capacity)
                cash_before_funding += debt_draw
            elif cash_before_funding > d.minimum_cash:
                surplus = cash_before_funding - d.minimum_cash
                outstanding_after_scheduled = max(
                    0.0, opening_debt - scheduled_repayment
                )
                cash_sweep = min(surplus, outstanding_after_scheduled)
                cash_before_funding -= cash_sweep

            closing_debt = max(
                0.0,
                opening_debt - scheduled_repayment - cash_sweep + debt_draw,
            )
            debt_current = min(
                closing_debt,
                d.scheduled_debt_repayment * 12,
            )
            debt_non_current = max(0.0, closing_debt - debt_current)

            cfi = -capex
            cff = debt_draw - scheduled_repayment - cash_sweep - dividend
            net_cash_change = cfo + cfi + cff
            closing_cash = cash + net_cash_change

            gross_ppe += capex
            accum_dep -= depreciation
            retained_earnings += net_income - dividend

            cash = closing_cash
            ar = target_ar
            inventory = target_inventory
            oca = target_oca
            ap = target_ap
            accrued = target_accrued
            ocl = target_ocl

            total_assets = (
                cash + ar + inventory + oca
                + gross_ppe + accum_dep + ona
            )
            total_liabilities = (
                ap + accrued + ocl + debt_current
                + debt_non_current + oncl
            )
            total_equity = share_capital + retained_earnings
            balance_check = total_assets - total_liabilities - total_equity

            rows.append({
                "Period": period,
                "Revenue": revenue,
                "COGS": cogs,
                "Gross Profit": gross_profit,
                "Payroll": payroll,
                "Other Opex": other_opex,
                "EBITDA": ebitda,
                "Depreciation": depreciation,
                "EBIT": ebit,
                "Interest Expense": interest,
                "Profit Before Tax": pbt,
                "Tax Expense": tax,
                "Net Income": net_income,

                "Cash From Operations": cfo,
                "Capital Expenditure": capex,
                "Cash From Investing": cfi,
                "Debt Draw": debt_draw,
                "Scheduled Debt Repayment": scheduled_repayment,
                "Cash Sweep": cash_sweep,
                "Dividends": dividend,
                "Cash From Financing": cff,
                "Net Change in Cash": net_cash_change,

                "Cash": cash,
                "Accounts Receivable": ar,
                "Inventory": inventory,
                "Other Current Assets": oca,
                "Gross PPE": gross_ppe,
                "Accumulated Depreciation": accum_dep,
                "Net PPE": gross_ppe + accum_dep,
                "Other Non-current Assets": ona,
                "Accounts Payable": ap,
                "Accrued Expenses": accrued,
                "Other Current Liabilities": ocl,
                "Current Debt": debt_current,
                "Non-current Debt": debt_non_current,
                "Other Non-current Liabilities": oncl,
                "Share Capital": share_capital,
                "Retained Earnings": retained_earnings,
                "Total Assets": total_assets,
                "Total Liabilities": total_liabilities,
                "Total Equity": total_equity,
                "Balance Check": balance_check,
            })

            wc_schedule.append({
                "Period": period,
                "DSO": d.dso_days,
                "DPO": d.dpo_days,
                "Inventory Days": d.inventory_days,
                "Accounts Receivable": ar,
                "Inventory": inventory,
                "Accounts Payable": ap,
                "Net Working Capital": ar + inventory + oca - ap - accrued - ocl,
                "Change in AR": delta_ar,
                "Change in Inventory": delta_inventory,
                "Change in AP": delta_ap,
            })

            debt_schedule.append({
                "Period": period,
                "Opening Debt": opening_debt,
                "Interest": interest,
                "Draw": debt_draw,
                "Scheduled Repayment": scheduled_repayment,
                "Cash Sweep": cash_sweep,
                "Closing Debt": closing_debt,
            })

            ppe_schedule.append({
                "Period": period,
                "Opening Gross PPE": gross_ppe - capex,
                "Capex": capex,
                "Closing Gross PPE": gross_ppe,
                "Depreciation": depreciation,
                "Accumulated Depreciation": accum_dep,
                "Net PPE": gross_ppe + accum_dep,
            })

        df = pd.DataFrame(rows).set_index("Period")

        pl = df[[
            "Revenue", "COGS", "Gross Profit", "Payroll",
            "Other Opex", "EBITDA", "Depreciation", "EBIT",
            "Interest Expense", "Profit Before Tax",
            "Tax Expense", "Net Income",
        ]].copy()

        bs = df[[
            "Cash", "Accounts Receivable", "Inventory",
            "Other Current Assets", "Gross PPE",
            "Accumulated Depreciation", "Net PPE",
            "Other Non-current Assets", "Total Assets",
            "Accounts Payable", "Accrued Expenses",
            "Other Current Liabilities", "Current Debt",
            "Non-current Debt", "Other Non-current Liabilities",
            "Total Liabilities", "Share Capital",
            "Retained Earnings", "Total Equity", "Balance Check",
        ]].copy()

        cf = df[[
            "Net Income", "Depreciation", "Cash From Operations",
            "Capital Expenditure", "Cash From Investing",
            "Debt Draw", "Scheduled Debt Repayment", "Cash Sweep",
            "Dividends", "Cash From Financing",
            "Net Change in Cash", "Cash",
        ]].copy()

        ratios = pd.DataFrame(index=df.index)
        ratios["Gross Margin"] = np.where(
            df["Revenue"] != 0,
            df["Gross Profit"] / df["Revenue"],
            np.nan,
        )
        ratios["EBITDA Margin"] = np.where(
            df["Revenue"] != 0,
            df["EBITDA"] / df["Revenue"],
            np.nan,
        )
        ratios["Net Margin"] = np.where(
            df["Revenue"] != 0,
            df["Net Income"] / df["Revenue"],
            np.nan,
        )
        current_assets = (
            df["Cash"] + df["Accounts Receivable"]
            + df["Inventory"] + df["Other Current Assets"]
        )
        current_liabilities = (
            df["Accounts Payable"] + df["Accrued Expenses"]
            + df["Other Current Liabilities"] + df["Current Debt"]
        )
        ratios["Current Ratio"] = np.where(
            current_liabilities != 0,
            current_assets / current_liabilities,
            np.nan,
        )
        ratios["Quick Ratio"] = np.where(
            current_liabilities != 0,
            (df["Cash"] + df["Accounts Receivable"])
            / current_liabilities,
            np.nan,
        )
        total_debt = df["Current Debt"] + df["Non-current Debt"]
        ratios["Debt / Equity"] = np.where(
            df["Total Equity"] != 0,
            total_debt / df["Total Equity"],
            np.nan,
        )
        ratios["Interest Cover"] = np.where(
            df["Interest Expense"] != 0,
            df["EBIT"] / df["Interest Expense"],
            np.nan,
        )
        ratios["DSO"] = d.dso_days
        ratios["DPO"] = d.dpo_days
        ratios["Inventory Days"] = d.inventory_days
        ratios["Cash Conversion Cycle"] = (
            d.dso_days + d.inventory_days - d.dpo_days
        )

        checks = pd.DataFrame(index=df.index)
        checks["Balance Check"] = df["Balance Check"]
        checks["Balanced"] = checks["Balance Check"].abs() < 0.01
        checks["Cash Below Minimum"] = df["Cash"] < d.minimum_cash - 0.01
        checks["Debt Above Limit"] = total_debt > d.revolver_limit + 0.01

        return ThreeWayForecastResult(
            profit_and_loss=pl,
            balance_sheet=bs,
            cash_flow=cf,
            ratios=ratios,
            schedules={
                "Working Capital": pd.DataFrame(wc_schedule).set_index("Period"),
                "Debt": pd.DataFrame(debt_schedule).set_index("Period"),
                "PPE": pd.DataFrame(ppe_schedule).set_index("Period"),
            },
            checks=checks,
            forecast_basis=self.forecast.set_index("Period"),
        )
