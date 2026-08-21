"use client";

import { useEffect, useMemo, useState } from "react";
import { useSearchParams } from "next/navigation";
import {
  AlertTriangle,
  CheckCircle2,
  Database,
  Loader2,
  Play,
  RefreshCw,
  Sparkles,
  X,
} from "lucide-react";

import { ModuleResetButton } from "@/components/module-reset-button";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { formatMoney } from "@/lib/finance-format";
import { advancedFinanceService } from "@/services/advanced-finance-service";
import type { ForecastRun, PlanningBaseline } from "@/types/advanced-forecasting";

const fallbackDrivers = {
  gross_margin: 0.42,
  payroll_pct_revenue: 0.22,
  other_opex_pct_revenue: 0.14,
  annual_interest_rate: 0.075,
  tax_rate: 0.30,
  dso_days: 42,
  dpo_days: 35,
  inventory_days: 45,
  capex_pct_revenue: 0.04,
  useful_life_months: 60,
  scheduled_debt_repayment: 10000,
  minimum_cash: 100000,
  revolver_limit: 500000,
  dividend_pct_net_income: 0,
};

const emptyOpeningBalanceSheet = {
  cash: 0,
  accounts_receivable: 0,
  inventory: 0,
  other_current_assets: 0,
  gross_ppe: 0,
  accumulated_depreciation: 0,
  other_non_current_assets: 0,
  accounts_payable: 0,
  accrued_expenses: 0,
  other_current_liabilities: 0,
  debt_current: 0,
  debt_non_current: 0,
  other_non_current_liabilities: 0,
  share_capital: 0,
  retained_earnings: null as number | null,
};

const initialForm = {
  run_name: "24 Month Management Forecast",
  forecast_start: "",
  forecast_months: 24,
  trend_weight: 0.45,
  budget_weight: 0.35,
  run_rate_weight: 0.20,
  seasonality_enabled: true,
  drivers: fallbackDrivers,
  opening_balance_sheet: emptyOpeningBalanceSheet,
};

export default function ThreeWay() {
  const params = useSearchParams();
  const [form, setForm] = useState<any>(initialForm);
  const [result, setResult] = useState<ForecastRun | null>(null);
  const [loading, setLoading] = useState(false);
  const [baselineLoading, setBaselineLoading] = useState(true);
  const [baseline, setBaseline] = useState<PlanningBaseline | null>(null);
  const [baselineError, setBaselineError] = useState("");
  const [aiBanner, setAiBanner] = useState(false);

  async function loadBaseline() {
    setBaselineLoading(true);
    setBaselineError("");
    try {
      const actual = await advancedFinanceService.planningBaseline();
      setBaseline(actual);
      setForm((current: any) => ({
        ...current,
        forecast_start: actual.suggested_forecast_start,
        drivers: {
          ...current.drivers,
          ...(actual.suggested_drivers || {}),
        },
        opening_balance_sheet: {
          ...emptyOpeningBalanceSheet,
          ...actual.opening_balance_sheet,
        },
      }));
    } catch {
      setBaseline(null);
      setBaselineError(
        "FinCruiz could not build a forecast baseline from mapped actuals. Load and map finance data before running a live-company three-way forecast.",
      );
    } finally {
      setBaselineLoading(false);
    }
  }

  useEffect(() => {
    void loadBaseline();
  }, []);

  useEffect(() => {
    if (params.get("from_ai") === "1") {
      const scenario = params.get("scenario") || "Management decision";
      const headcount = Number(params.get("headcount_change") || 0);
      setForm((current: any) => ({
        ...current,
        run_name: `AI scenario — ${scenario}${headcount ? ` (+${headcount} people)` : ""}`,
      }));
      setAiBanner(true);
    }
  }, [params]);

  const canRun = Boolean(baseline && form.forecast_start && !baselineLoading);

  async function run() {
    if (!canRun) return;
    setLoading(true);
    try {
      setResult(await advancedFinanceService.runForecast(form));
    } finally {
      setLoading(false);
    }
  }

  function updateDriver(key: string, value: number) {
    setForm((current: any) => ({
      ...current,
      drivers: { ...current.drivers, [key]: value },
    }));
  }

  function updateOpening(key: string, value: number | null) {
    setForm((current: any) => ({
      ...current,
      opening_balance_sheet: { ...current.opening_balance_sheet, [key]: value },
    }));
  }

  const baselineLabel = useMemo(() => {
    if (!baseline) return "No actual-data baseline loaded";
    return `${baseline.history_months} months of mapped actuals · through ${baseline.period_end}`;
  }, [baseline]);

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      {aiBanner ? (
        <div className="flex items-start justify-between gap-4 rounded-2xl border border-violet-200 bg-gradient-to-r from-violet-50 to-indigo-50 p-4 dark:border-violet-500/20 dark:from-violet-950/30 dark:to-indigo-950/20">
          <div className="flex gap-3">
            <div className="mt-0.5 flex size-9 items-center justify-center rounded-xl bg-violet-600 text-white"><Sparkles className="size-4" /></div>
            <div>
              <p className="font-semibold">Decision handed off from Ask FinCruiz</p>
              <p className="mt-1 text-sm text-muted-foreground">The scenario name is carried into the model, but financial assumptions remain visible for review. FinCruiz never silently runs a decision model.</p>
            </div>
          </div>
          <button onClick={() => setAiBanner(false)} aria-label="Dismiss" className="rounded-lg p-1 hover:bg-white/70"><X className="size-4" /></button>
        </div>
      ) : null}

      <div>
        <p className="text-sm text-muted-foreground">Forecasting</p>
        <div className="flex items-center justify-between gap-4">
          <h1 className="text-3xl font-semibold">Integrated Three-Way Forecast</h1>
          <ModuleResetButton scope="forecasts" label="Reset forecasts" description="Remove saved forecast and scenario runs only. Actual finance data remains." />
        </div>
        <p className="mt-2 text-muted-foreground">Linked P&L, Balance Sheet and Cash Flow built from your mapped actuals, working-capital assumptions and management scenarios.</p>
      </div>

      <Card className={baselineError ? "border-amber-300" : "border-emerald-200 dark:border-emerald-500/20"}>
        <CardContent className="flex flex-col gap-3 p-5 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex items-start gap-3">
            <div className={`flex size-10 shrink-0 items-center justify-center rounded-xl ${baselineError ? "bg-amber-100 text-amber-700" : "bg-emerald-100 text-emerald-700"}`}>
              {baselineLoading ? <Loader2 className="size-4 animate-spin" /> : baselineError ? <AlertTriangle className="size-4" /> : <Database className="size-4" />}
            </div>
            <div>
              <p className="font-semibold">Actual-data forecast baseline</p>
              <p className="mt-1 text-sm text-muted-foreground">{baselineLoading ? "Reading mapped finance data…" : baselineError || baselineLabel}</p>
              {baseline ? <p className="mt-1 text-xs text-muted-foreground">Forecast starts {baseline.suggested_forecast_start}. Opening balances and recent operating ratios have been prefilled from company data and remain editable.</p> : null}
            </div>
          </div>
          <Button type="button" variant="outline" onClick={() => void loadBaseline()} disabled={baselineLoading}><RefreshCw className={`size-4 ${baselineLoading ? "animate-spin" : ""}`} />Reload actuals</Button>
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle>Forecast configuration</CardTitle>
          <CardDescription>Review the assumptions before running. Percentages are shown as human-readable percentages; the API receives decimal rates.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-6">
          <div className="grid gap-4 md:grid-cols-4">
            <Field label="Run name"><Input value={form.run_name} onChange={(e) => setForm({ ...form, run_name: e.target.value })} /></Field>
            <Field label="Forecast start"><Input type="date" value={form.forecast_start} onChange={(e) => setForm({ ...form, forecast_start: e.target.value })} /></Field>
            <Field label="Forecast months"><Input type="number" min={1} max={60} value={form.forecast_months} onChange={(e) => setForm({ ...form, forecast_months: Number(e.target.value) })} /></Field>
            <div className="flex items-end"><Button className="w-full" onClick={() => void run()} disabled={loading || !canRun}>{loading ? <Loader2 className="size-4 animate-spin" /> : <Play className="size-4" />}Run forecast</Button></div>
          </div>

          <AssumptionSection title="Operating model" description="These assumptions drive gross profit and EBITDA.">
            <PercentField label="Gross margin" value={form.drivers.gross_margin} onChange={(v) => updateDriver("gross_margin", v)} />
            <PercentField label="Payroll % of revenue" value={form.drivers.payroll_pct_revenue} onChange={(v) => updateDriver("payroll_pct_revenue", v)} />
            <PercentField label="Other opex % of revenue" value={form.drivers.other_opex_pct_revenue} onChange={(v) => updateDriver("other_opex_pct_revenue", v)} />
            <PercentField label="Tax rate" value={form.drivers.tax_rate} onChange={(v) => updateDriver("tax_rate", v)} />
          </AssumptionSection>

          <AssumptionSection title="Working capital" description="Collection, supplier and inventory assumptions flow directly into cash.">
            <NumberField label="DSO (days)" value={form.drivers.dso_days} onChange={(v) => updateDriver("dso_days", v)} />
            <NumberField label="DPO (days)" value={form.drivers.dpo_days} onChange={(v) => updateDriver("dpo_days", v)} />
            <NumberField label="Inventory days" value={form.drivers.inventory_days} onChange={(v) => updateDriver("inventory_days", v)} />
            <NumberField label="Minimum cash" value={form.drivers.minimum_cash} onChange={(v) => updateDriver("minimum_cash", v)} />
          </AssumptionSection>

          <AssumptionSection title="Capital & funding" description="Capex, financing and depreciation assumptions are explicit rather than hidden in the model.">
            <PercentField label="Capex % of revenue" value={form.drivers.capex_pct_revenue} onChange={(v) => updateDriver("capex_pct_revenue", v)} />
            <PercentField label="Interest rate" value={form.drivers.annual_interest_rate} onChange={(v) => updateDriver("annual_interest_rate", v)} />
            <PercentField label="Dividend % of net income" value={form.drivers.dividend_pct_net_income} onChange={(v) => updateDriver("dividend_pct_net_income", v)} />
            <NumberField label="Useful life (months)" value={form.drivers.useful_life_months} onChange={(v) => updateDriver("useful_life_months", Math.max(1, Math.round(v)))} />
            <NumberField label="Monthly debt repayment" value={form.drivers.scheduled_debt_repayment} onChange={(v) => updateDriver("scheduled_debt_repayment", v)} />
            <NumberField label="Revolver limit" value={form.drivers.revolver_limit} onChange={(v) => updateDriver("revolver_limit", v)} />
          </AssumptionSection>

          <details className="rounded-2xl border bg-muted/20 p-4">
            <summary className="cursor-pointer font-semibold">Opening Balance Sheet — review source balances</summary>
            <p className="mt-2 text-sm text-muted-foreground">These values are populated from mapped actuals. Edit only when you intentionally want the forecast opening position to differ from source data.</p>
            <div className="mt-4 grid gap-4 md:grid-cols-4">
              {Object.entries(
                form.opening_balance_sheet as Record<string, number | null>,
              ).map(([key, value]) => (
                <Field key={key} label={humanize(key)}>
                  <Input
                    type="number"
                    step="0.01"
                    value={value === null ? "" : value}
                    onChange={(e) =>
                      updateOpening(
                        key,
                        e.target.value === "" ? null : Number(e.target.value),
                      )
                    }
                  />
                </Field>
              ))}
            </div>
          </details>
        </CardContent>
      </Card>

      {result ? (
        <>
          <div className="grid gap-4 md:grid-cols-5">
            {[["Revenue", "forecast_revenue"], ["EBITDA", "forecast_ebitda"], ["Net income", "forecast_net_income"], ["Closing cash", "closing_cash"], ["Closing debt", "closing_debt"]].map(([label, key]) => (
              <Card key={key}><CardHeader><CardDescription>{label}</CardDescription><CardTitle>{formatMoney(result.summary[key] as number)}</CardTitle></CardHeader></Card>
            ))}
          </div>
          <Card>
            <CardHeader><CardTitle className="flex items-center gap-2"><CheckCircle2 className="size-5 text-emerald-600" />Integrity checks</CardTitle></CardHeader>
            <CardContent><p>Balanced across all periods: <b>{String(result.summary.balanced)}</b></p><p>Minimum cash: <b>{formatMoney(result.summary.minimum_cash as number)}</b></p></CardContent>
          </Card>
          <Statement title="Profit & Loss" rows={result.profit_and_loss} />
          <Statement title="Balance Sheet" rows={result.balance_sheet} />
          <Statement title="Cash Flow" rows={result.cash_flow} />
        </>
      ) : null}
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return <label className="space-y-1.5 text-sm"><span className="font-medium">{label}</span>{children}</label>;
}

function NumberField({ label, value, onChange }: { label: string; value: number; onChange: (value: number) => void }) {
  return <Field label={label}><Input type="number" step="0.01" value={value} onChange={(e) => onChange(Number(e.target.value))} /></Field>;
}

function PercentField({ label, value, onChange }: { label: string; value: number; onChange: (value: number) => void }) {
  return <Field label={label}><div className="relative"><Input className="pr-8" type="number" step="0.1" value={Number((value * 100).toFixed(2))} onChange={(e) => onChange(Number(e.target.value) / 100)} /><span className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-sm text-muted-foreground">%</span></div></Field>;
}

function AssumptionSection({ title, description, children }: { title: string; description: string; children: React.ReactNode }) {
  return <section><div className="mb-3"><h3 className="font-semibold">{title}</h3><p className="text-xs text-muted-foreground">{description}</p></div><div className="grid gap-4 md:grid-cols-4">{children}</div></section>;
}

function humanize(value: string) {
  return value.split("_").map((part) => part.charAt(0).toUpperCase() + part.slice(1)).join(" ");
}

function Statement({ title, rows }: { title: string; rows: Array<Record<string, unknown>> }) {
  const keys = rows.length ? Object.keys(rows[0]) : [];
  return <Card><CardHeader><CardTitle>{title}</CardTitle></CardHeader><CardContent><div className="overflow-x-auto"><table className="w-full min-w-[1100px] text-xs"><thead><tr>{keys.map((key) => <th className="p-2 text-left" key={key}>{key}</th>)}</tr></thead><tbody>{rows.map((row, index) => <tr className="border-t" key={index}>{keys.map((key) => <td className="p-2" key={key}>{typeof row[key] === "number" ? Number(row[key]).toLocaleString() : String(row[key] ?? "")}</td>)}</tr>)}</tbody></table></div></CardContent></Card>;
}
