"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";
import { ModuleResetButton } from "@/components/module-reset-button";
import { InsightChart } from "@/components/insight-chart";
import { Loader2, TrendingUp } from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { getApiErrorMessage } from "@/lib/api";
import { formatMoney } from "@/lib/finance-format";
import { financeService } from "@/services/finance-service";
import type { Branch, ForecastResult } from "@/types/finance";
import type { AIVisualization } from "@/types/analytics";

export default function ForecastingPage() {
  const [branches, setBranches] = useState<Branch[]>([]);
  const [branchId, setBranchId] = useState("");
  const [group, setGroup] = useState("Revenue");
  const [method, setMethod] = useState("run_rate");
  const [months, setMonths] = useState(12);
  const [result, setResult] = useState<ForecastResult | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    financeService.getBranches().then(setBranches).catch(() => setBranches([]));
  }, []);

  const chart = useMemo<AIVisualization | null>(() => {
    if (!result) return null;
    return {
      type: "line", title: `${result.reporting_group} — forecast range`,
      subtitle: `${result.method.replace("_", " ")} · based on ${result.history_periods} historical periods · ${result.confidence} confidence`,
      labels: result.points.map(p => String(p.period).slice(0,7)),
      series: [
        { name: "Base forecast", data: result.points.map(p => Number(p.base||0)) },
        { name: "Downside", data: result.points.map(p => Number(p.downside||0)) },
        { name: "Upside", data: result.points.map(p => Number(p.upside||0)) },
      ], value_format: "currency",
    };
  }, [result]);

  async function generate(event: FormEvent) {
    event.preventDefault();
    setIsLoading(true);
    setError("");
    try {
      setResult(await financeService.createForecast({
        reporting_group: group,
        future_months: months,
        method,
        branch_id: branchId || null,
        downside_factor: 0.9,
        upside_factor: 1.1,
        recent_months: 3,
      }));
    } catch (forecastError) {
      setError(getApiErrorMessage(forecastError));
    } finally {
      setIsLoading(false);
    }
  }

  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div>
        <p className="text-sm font-medium text-muted-foreground">Planning</p>
        <div className="flex items-center justify-between gap-4"><h1 className="mt-1 text-3xl font-semibold tracking-tight">Forecasting</h1><ModuleResetButton scope="forecasts" label="Reset forecasts" description="This removes saved forecast and scenario runs. Your actual financial data remains." /></div>
        <p className="mt-2 text-muted-foreground">
          Run-rate and linear-trend forecasts using saved monthly actuals.
        </p>
      </div>

      {error ? <Alert variant="destructive"><AlertDescription>{error}</AlertDescription></Alert> : null}
      {result?.warning ? <Alert><AlertDescription>{result.warning}</AlertDescription></Alert> : null}

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2"><TrendingUp className="size-5" />Forecast assumptions</CardTitle>
          <CardDescription>Use at least 12 complete months for stronger confidence.</CardDescription>
        </CardHeader>
        <CardContent>
          <form className="grid gap-3 md:grid-cols-4" onSubmit={generate}>
            <select className="h-10 rounded-md border bg-background px-3" value={branchId} onChange={(event) => setBranchId(event.target.value)}>
              <option value="">Consolidated company</option>
              {branches.filter((branch) => branch.is_active).map((branch) => (
                <option key={branch.id} value={branch.id}>{branch.branch_code} — {branch.branch_name}</option>
              ))}
            </select>
            <select className="h-10 rounded-md border bg-background px-3" value={group} onChange={(event) => setGroup(event.target.value)}>
              {["Revenue", "Cost of Sales", "Operating Expenses", "Depreciation", "Finance Costs"].map((value) => (
                <option key={value} value={value}>{value}</option>
              ))}
            </select>
            <select className="h-10 rounded-md border bg-background px-3" value={method} onChange={(event) => setMethod(event.target.value)}>
              <option value="run_rate">Run rate</option>
              <option value="trend">Linear trend</option>
            </select>
            <div className="flex gap-2">
              <Input type="number" min={1} max={60} value={months} onChange={(event) => setMonths(Number(event.target.value))} />
              <Button disabled={isLoading}>
                {isLoading ? <Loader2 className="size-4 animate-spin" /> : <TrendingUp className="size-4" />}
                Generate
              </Button>
            </div>
          </form>
        </CardContent>
      </Card>

      {result ? (
        <>
        {chart ? <InsightChart visualization={chart} /> : null}
        <Card>
          <CardHeader>
            <CardTitle>{result.reporting_group} forecast</CardTitle>
            <CardDescription>
              {result.method.replace("_", " ")} · {result.history_periods} historical periods · {result.confidence} confidence
            </CardDescription>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
              <table className="w-full min-w-[700px] text-sm">
                <thead><tr className="border-b text-left text-muted-foreground"><th className="p-3">Period</th><th className="p-3">Downside</th><th className="p-3">Base</th><th className="p-3">Upside</th></tr></thead>
                <tbody>
                  {result.points.map((point) => (
                    <tr key={point.period} className="border-b">
                      <td className="p-3">{point.period}</td>
                      <td className="p-3">{formatMoney(point.downside)}</td>
                      <td className="p-3 font-semibold">{formatMoney(point.base)}</td>
                      <td className="p-3">{formatMoney(point.upside)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </CardContent>
        </Card>
        </>
      ) : null}
    </div>
  );
}
