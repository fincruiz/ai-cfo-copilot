"use client";

import Link from "next/link";
import { useEffect, useMemo, useState } from "react";
import { ModuleResetButton } from "@/components/module-reset-button";
import { ArrowRight, CheckCircle2, Loader2, RefreshCw, Save, WandSparkles } from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { getApiErrorMessage } from "@/lib/api";
import { humanize } from "@/lib/finance-format";
import { financeService } from "@/services/finance-service";
import type { AccountMappingInput, MappingSuggestion } from "@/types/finance";

const statements = ["income_statement", "balance_sheet"];
const groups = ["Revenue", "Cost of Sales", "Operating Expenses", "Depreciation", "Other Income", "Other Expenses", "Finance Costs", "Tax", "Current Assets", "Non Current Assets", "Current Liabilities", "Non Current Liabilities", "Equity"];

export default function MappingPage() {
  const [rows, setRows] = useState<MappingSuggestion[]>([]);
  const [savedCount, setSavedCount] = useState(0);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [search, setSearch] = useState("");
  const [welcomeFlow, setWelcomeFlow] = useState(false);

  async function load() {
    setIsLoading(true);
    setError("");
    try {
      const [suggestions, mappings] = await Promise.all([
        financeService.getMappingSuggestions(),
        financeService.getMappings(),
      ]);
      setSavedCount(mappings.length);
      setRows(suggestions);
    } catch (loadError) {
      setError(getApiErrorMessage(loadError));
    } finally {
      setIsLoading(false);
    }
  }

  useEffect(() => { setWelcomeFlow(new URLSearchParams(window.location.search).get("welcome") === "1"); void load(); }, []);

  const filteredRows = useMemo(() => {
    const term = search.trim().toLowerCase();
    if (!term) return rows;
    return rows.filter((row) => `${row.source_account_code} ${row.source_account_name ?? ""} ${row.reporting_group}`.toLowerCase().includes(term));
  }, [rows, search]);

  function update(index: number, field: keyof MappingSuggestion, value: string) {
    setRows((current) => current.map((row, rowIndex) => rowIndex === index ? { ...row, [field]: value } : row));
  }

  async function save() {
    setIsSaving(true);
    setError("");
    setSuccess("");
    try {
      const items: AccountMappingInput[] = rows.map((row, index) => ({
        source_account_code: row.source_account_code,
        source_account_name: row.source_account_name,
        statement: row.statement,
        reporting_group: row.reporting_group,
        reporting_subgroup: row.reporting_subgroup,
        sign_convention: row.sign_convention,
        display_order: (index + 1) * 10,
        is_confirmed: true,
      }));
      const count = await financeService.saveMappings(items);
      setSuccess(`${count} account mappings saved. Reports are ready to refresh.`);
      setSavedCount((value) => value + count);
      setRows([]);
    } catch (saveError) {
      setError(getApiErrorMessage(saveError));
    } finally {
      setIsSaving(false);
    }
  }

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
        <div>
          <p className="text-sm font-medium text-muted-foreground">Finance intelligence</p>
          <div className="flex items-center justify-between gap-4"><h1 className="mt-1 text-3xl font-semibold tracking-tight">Account mapping</h1><ModuleResetButton scope="account_mappings" label="Reset mappings" description="This removes only saved account mappings. Your General Ledger remains loaded." /></div>
          <p className="mt-2 text-muted-foreground">Review AI suggestions before they drive your financial reports.</p>
        </div>
        <div className="flex gap-2">
          <Button variant="outline" onClick={() => void load()} disabled={isLoading}><RefreshCw className="size-4" />Refresh</Button>
          <Button onClick={() => void save()} disabled={!rows.length || isSaving}>{isSaving ? <Loader2 className="size-4 animate-spin" /> : <Save className="size-4" />}Save all mappings</Button>
        </div>
      </div>

      {error ? <Alert variant="destructive"><AlertDescription>{error}</AlertDescription></Alert> : null}
      {success ? <Alert><CheckCircle2 className="size-4" /><AlertDescription>{success}</AlertDescription></Alert> : null}
      {welcomeFlow && (success || (!isLoading && rows.length === 0 && savedCount > 0)) ? <div className="flex justify-end"><Link href="/dashboard/getting-started" className="inline-flex items-center gap-2 rounded-xl bg-primary px-4 py-2 text-sm font-semibold text-primary-foreground">Continue guided setup<ArrowRight className="size-4"/></Link></div> : null}

      <div className="grid gap-4 sm:grid-cols-3">
        <Card><CardHeader className="pb-2"><CardDescription>Saved mappings</CardDescription><CardTitle className="text-3xl">{savedCount}</CardTitle></CardHeader></Card>
        <Card><CardHeader className="pb-2"><CardDescription>Suggestions awaiting approval</CardDescription><CardTitle className="text-3xl">{rows.length}</CardTitle></CardHeader></Card>
        <Card><CardHeader className="pb-2"><CardDescription>Average confidence</CardDescription><CardTitle className="text-3xl">{rows.length ? `${Math.round(rows.reduce((sum, row) => sum + row.confidence, 0) / rows.length * 100)}%` : "—"}</CardTitle></CardHeader></Card>
      </div>

      <Card>
        <CardHeader>
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
            <div><CardTitle className="flex items-center gap-2"><WandSparkles className="size-5" />Mapping suggestions</CardTitle><CardDescription>Edit any classification before saving.</CardDescription></div>
            <Input className="sm:max-w-xs" placeholder="Search accounts..." value={search} onChange={(event) => setSearch(event.target.value)} />
          </div>
        </CardHeader>
        <CardContent>
          {isLoading ? <div className="flex min-h-52 items-center justify-center gap-2 text-muted-foreground"><Loader2 className="size-5 animate-spin" />Loading suggestions...</div> : !rows.length ? (
            <div className="flex min-h-52 flex-col items-center justify-center text-center"><CheckCircle2 className="mb-3 size-9 text-emerald-600" /><p className="font-medium">No unmapped accounts</p><p className="mt-1 text-sm text-muted-foreground">All detected accounts have mappings.</p></div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full min-w-[1100px] text-sm">
                <thead><tr className="border-b text-left text-muted-foreground">{["Account", "Statement", "Reporting group", "Subgroup", "Sign", "Confidence", "Reason"].map((heading) => <th key={heading} className="px-3 py-3 font-medium">{heading}</th>)}</tr></thead>
                <tbody>
                  {filteredRows.map((row) => {
                    const sourceIndex = rows.findIndex((item) => item.source_account_code === row.source_account_code);
                    return (
                      <tr key={row.source_account_code} className="border-b align-top last:border-0">
                        <td className="px-3 py-3"><p className="font-medium">{row.source_account_code}</p><p className="text-muted-foreground">{row.source_account_name}</p></td>
                        <td className="px-3 py-3"><select className="h-9 rounded-md border bg-background px-2" value={row.statement} onChange={(event) => update(sourceIndex, "statement", event.target.value)}>{statements.map((item) => <option key={item} value={item}>{humanize(item)}</option>)}</select></td>
                        <td className="px-3 py-3"><select className="h-9 rounded-md border bg-background px-2" value={row.reporting_group} onChange={(event) => update(sourceIndex, "reporting_group", event.target.value)}>{groups.map((item) => <option key={item} value={item}>{item}</option>)}</select></td>
                        <td className="px-3 py-3"><Input className="min-w-48" value={row.reporting_subgroup ?? ""} onChange={(event) => update(sourceIndex, "reporting_subgroup", event.target.value)} /></td>
                        <td className="px-3 py-3"><select className="h-9 rounded-md border bg-background px-2" value={row.sign_convention} onChange={(event) => update(sourceIndex, "sign_convention", event.target.value)}><option value="debit">Debit</option><option value="credit">Credit</option><option value="positive">Positive</option></select></td>
                        <td className="px-3 py-3"><span className="rounded-full bg-muted px-2 py-1 font-medium">{Math.round(row.confidence * 100)}%</span></td>
                        <td className="px-3 py-3 text-muted-foreground">{row.reason}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}
