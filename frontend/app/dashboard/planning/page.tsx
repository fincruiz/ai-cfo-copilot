"use client";
import { useEffect, useState } from "react";
import { BarChart3, FileUp, Loader2, RefreshCw } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { formatMoney, toNumber } from "@/lib/finance-format";
import { planningService } from "@/services/planning-service";
import type { VarianceLine } from "@/types/planning";

export default function PlanningPage() {
  const [budget, setBudget] = useState<File | null>(null);
  const [forecast, setForecast] = useState<File | null>(null);
  const [version, setVersion] = useState("FY27 Working");
  const [rows, setRows] = useState<VarianceLine[]>([]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");

  async function load() {
    try {
      setRows(await planningService.getVariance());
    } catch (loadError) {
      console.error("Unable to load planning variance", loadError);
      setRows([]);
      setMessage(
        "Planning data could not be loaded. Confirm the planning migration has been run, then refresh.",
      );
    }
  }
  useEffect(() => { void load(); }, []);

  async function upload(kind: "budget" | "forecast") {
    const file = kind === "budget" ? budget : forecast;
    if (!file) return;
    setLoading(true);
    try {
      const result = kind === "budget"
        ? await planningService.uploadBudget(file, version)
        : await planningService.uploadForecast(file, version);
      setMessage(`${result.inserted_rows} ${kind} lines saved.`);
      await load();
    } finally { setLoading(false); }
  }

  function template() {
    const csv = "period,reporting_group,reporting_subgroup,account_code,branch,amount,notes\n2027-01,Revenue,Sales,4000,MEL,150000,Base plan\n2027-01,Operating Expenses,Rent,6100,MEL,30000,Lease\n";
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    const a = document.createElement("a"); a.href=url; a.download="budget_forecast_template.csv"; a.click(); URL.revokeObjectURL(url);
  }

  return <div className="mx-auto max-w-7xl space-y-7">
    <div className="animate-rise">
      <p className="text-sm font-medium text-muted-foreground">Planning & intelligence</p>
      <h1 className="mt-1 text-3xl font-semibold">Budgets, Forecasts & Variance</h1>
      <p className="mt-2 text-muted-foreground">Import existing plans now. The next phase enables native driver-based planning inside FinCruiz.</p>
    </div>
    {message ? <Alert><AlertDescription>{message}</AlertDescription></Alert> : null}
    <div className="grid gap-5 lg:grid-cols-2">
      {(["budget","forecast"] as const).map(kind => <Card key={kind} className="animate-card-in">
        <CardHeader><CardTitle className="capitalize">{kind} import</CardTitle><CardDescription>Monthly plan lines by reporting group, account and branch.</CardDescription></CardHeader>
        <CardContent className="space-y-4">
          <Input value={version} onChange={e=>setVersion(e.target.value)} placeholder="Version name" />
          <Input type="file" accept=".csv" onChange={e => kind==="budget" ? setBudget(e.target.files?.[0]??null) : setForecast(e.target.files?.[0]??null)} />
          <div className="flex gap-2">
            <Button onClick={()=>void upload(kind)} disabled={loading || !(kind==="budget"?budget:forecast)}><FileUp className="size-4"/>{loading?<Loader2 className="size-4 animate-spin"/>:"Upload"}</Button>
            <Button variant="outline" onClick={template}>Template</Button>
          </div>
        </CardContent>
      </Card>)}
    </div>
    <Card>
      <CardHeader className="flex-row items-center justify-between"><div><CardTitle>Actual vs Budget vs Forecast</CardTitle><CardDescription>Variance comparison from saved monthly actuals and plan uploads.</CardDescription></div><Button variant="outline" onClick={()=>void load()}><RefreshCw className="size-4"/>Refresh</Button></CardHeader>
      <CardContent>
        {!rows.length ? <div className="flex min-h-48 items-center justify-center text-muted-foreground">Upload a budget or forecast to activate comparison.</div> :
        <div className="overflow-x-auto"><table className="w-full min-w-[900px] text-sm"><thead><tr className="border-b text-left text-muted-foreground">{["Period","Group","Actual","Budget","Budget Var","Forecast","Forecast Var"].map(h=><th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{rows.map((r,i)=><tr key={`${r.period}-${r.reporting_group}-${i}`} className="border-b"><td className="p-3">{r.period}</td><td className="p-3">{r.reporting_group}</td><td className="p-3">{formatMoney(r.actual)}</td><td className="p-3">{formatMoney(r.budget)}</td><td className={toNumber(r.budget_variance)<0?"p-3 text-red-600":"p-3 text-emerald-600"}>{formatMoney(r.budget_variance)}</td><td className="p-3">{formatMoney(r.forecast)}</td><td className={toNumber(r.forecast_variance)<0?"p-3 text-red-600":"p-3 text-emerald-600"}>{formatMoney(r.forecast_variance)}</td></tr>)}</tbody></table></div>}
      </CardContent>
    </Card>
  </div>;
}
