"use client";

import { useEffect, useState } from "react";
import { ArrowRight, Loader2, RefreshCw, Search } from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { getApiErrorMessage } from "@/lib/api";
import { formatMoney, formatPercent, toNumber } from "@/lib/finance-format";
import { financeService } from "@/services/finance-service";
import type {
  BalanceSheet,
  Branch,
  BranchComparison,
  LedgerTransaction,
  MonthlyActual,
  ProfitAndLoss,
  ReportContext,
  TrialBalance,
} from "@/types/finance";

const tabs = [
  "trial-balance",
  "profit-and-loss",
  "balance-sheet",
  "monthly-actuals",
  "branch-comparison",
  "transactions",
] as const;
type Tab = typeof tabs[number];
type Filters = { startDate: string; endDate: string; branchId: string; accountCode: string };

function tabLabel(tab: Tab) {
  return ({
    "trial-balance": "Trial balance",
    "profit-and-loss": "Profit & loss",
    "balance-sheet": "Balance sheet",
    "monthly-actuals": "Monthly actuals",
    "branch-comparison": "Branch comparison",
    transactions: "Ledger transactions",
  } as Record<Tab, string>)[tab];
}

function friendlyDate(value?: string | null) {
  if (!value) return "—";
  return new Date(`${value.slice(0, 10)}T00:00:00`).toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" });
}

export default function ReportsPage() {
  const [tab, setTab] = useState<Tab>("trial-balance");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [branchId, setBranchId] = useState("");
  const [accountCode, setAccountCode] = useState("");
  const [branches, setBranches] = useState<Branch[]>([]);
  const [reportContext, setReportContext] = useState<ReportContext | null>(null);
  const [trialBalance, setTrialBalance] = useState<TrialBalance | null>(null);
  const [pnl, setPnl] = useState<ProfitAndLoss | null>(null);
  const [balanceSheet, setBalanceSheet] = useState<BalanceSheet | null>(null);
  const [monthly, setMonthly] = useState<MonthlyActual[]>([]);
  const [comparison, setComparison] = useState<BranchComparison[]>([]);
  const [transactions, setTransactions] = useState<LedgerTransaction[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState("");

  async function load(filters: Filters = { startDate, endDate, branchId, accountCode }) {
    setIsLoading(true);
    setError("");
    try {
      const common = { startDate: filters.startDate, endDate: filters.endDate, branchId: filters.branchId || undefined };
      const [branchRows, context, tb, profitLoss, bs, monthlyRows, comparisonRows, transactionRows] = await Promise.all([
        financeService.getBranches(),
        financeService.getReportContext({ branchId: filters.branchId || undefined }),
        financeService.getTrialBalance(common),
        financeService.getProfitAndLoss(common),
        financeService.getBalanceSheet({ endDate: filters.endDate, branchId: filters.branchId || undefined }),
        financeService.getMonthlyActuals(common),
        financeService.getBranchComparison({ startDate: filters.startDate, endDate: filters.endDate }),
        financeService.getLedgerTransactions({ ...common, accountCode: filters.accountCode || undefined, limit: 500 }),
      ]);
      setBranches(branchRows);
      setReportContext(context);
      setTrialBalance(tb);
      setPnl(profitLoss);
      setBalanceSheet(bs);
      setMonthly(monthlyRows);
      setComparison(comparisonRows);
      setTransactions(transactionRows);
    } catch (loadError) {
      setError(getApiErrorMessage(loadError));
    } finally {
      setIsLoading(false);
    }
  }

  useEffect(() => {
    const search = new URLSearchParams(window.location.search);
    const requestedTab = search.get("tab") as Tab | null;
    const initialTab = requestedTab && tabs.includes(requestedTab) ? requestedTab : "trial-balance";
    const filters = {
      startDate: search.get("start_date") ?? "",
      endDate: search.get("end_date") ?? "",
      branchId: search.get("branch_id") ?? "",
      accountCode: search.get("account_code") ?? "",
    };
    setTab(initialTab);
    setStartDate(filters.startDate);
    setEndDate(filters.endDate);
    setBranchId(filters.branchId);
    setAccountCode(filters.accountCode);
    void load(filters);
    // Initial URL state only.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function syncUrl(nextTab: Tab, nextAccountCode = accountCode) {
    const search = new URLSearchParams(window.location.search);
    search.set("tab", nextTab);
    if (nextAccountCode) search.set("account_code", nextAccountCode); else search.delete("account_code");
    if (startDate) search.set("start_date", startDate); else search.delete("start_date");
    if (endDate) search.set("end_date", endDate); else search.delete("end_date");
    if (branchId) search.set("branch_id", branchId); else search.delete("branch_id");
    window.history.replaceState(null, "", `${window.location.pathname}?${search.toString()}`);
  }

  function chooseTab(next: Tab) {
    setTab(next);
    syncUrl(next);
  }

  async function openAccountTransactions(code: string) {
    setAccountCode(code);
    setTab("transactions");
    syncUrl("transactions", code);
    await load({ startDate, endDate, branchId, accountCode: code });
  }

  const current = tab === "trial-balance" ? trialBalance : tab === "profit-and-loss" ? pnl : tab === "balance-sheet" ? balanceSheet : null;
  const lines = current?.lines ?? [];

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      <div className="flex flex-col gap-4 xl:flex-row xl:items-end xl:justify-between">
        <div>
          <p className="text-sm font-medium text-muted-foreground">Financial reporting</p>
          <h1 className="mt-1 text-3xl font-semibold tracking-tight">Reports</h1>
          <p className="mt-2 text-muted-foreground">Consolidated and branch-level reports from the active validated ledger.</p>
          <p className="mt-2 text-xs font-medium text-muted-foreground">
            Reporting period {friendlyDate(reportContext?.period_start)} – {friendlyDate(reportContext?.period_end)} · Data as of {friendlyDate(reportContext?.data_as_of)} · {(reportContext?.transaction_count ?? 0).toLocaleString()} transactions
          </p>
        </div>

        <div className="flex flex-wrap items-end gap-3">
          <div>
            <p className="mb-1 text-xs text-muted-foreground">View</p>
            <select className="h-10 min-w-52 rounded-md border bg-background px-3 text-sm" value={branchId} onChange={(event) => setBranchId(event.target.value)}>
              <option value="">Consolidated company</option>
              {branches.filter((branch) => branch.is_active).map((branch) => <option key={branch.id} value={branch.id}>{branch.branch_code} — {branch.branch_name}</option>)}
            </select>
          </div>
          <div><p className="mb-1 text-xs text-muted-foreground">Start date</p><Input type="date" value={startDate} onChange={(event) => setStartDate(event.target.value)} /></div>
          <div><p className="mb-1 text-xs text-muted-foreground">End date</p><Input type="date" value={endDate} onChange={(event) => setEndDate(event.target.value)} /></div>
          <Button onClick={() => { syncUrl(tab); void load(); }} disabled={isLoading}><RefreshCw className="size-4" />Refresh</Button>
        </div>
      </div>

      {error ? <Alert variant="destructive"><AlertDescription>{error}</AlertDescription></Alert> : null}

      <div className="flex flex-wrap gap-2">
        {tabs.map((item) => <Button key={item} variant={tab === item ? "default" : "outline"} onClick={() => chooseTab(item)}>{tabLabel(item)}</Button>)}
      </div>

      {tab === "trial-balance" && trialBalance ? (
        <div className="grid gap-4 sm:grid-cols-3"><Metric label="Total debit" value={trialBalance.total_debit} /><Metric label="Total credit" value={trialBalance.total_credit} /><Metric label="Difference" value={trialBalance.difference} good={Math.abs(toNumber(trialBalance.difference)) < 0.01} /></div>
      ) : null}

      {tab === "profit-and-loss" && pnl ? (
        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4"><Metric label="Revenue" value={pnl.revenue} /><Metric label="Gross profit" value={pnl.gross_profit} /><Metric label="Operating profit" value={pnl.operating_profit} /><Metric label="Net profit" value={pnl.net_profit} /></div>
      ) : null}

      {tab === "balance-sheet" && balanceSheet ? (
        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4"><Metric label="Total assets" value={balanceSheet.total_assets} /><Metric label="Total liabilities" value={balanceSheet.total_liabilities} /><Metric label="Equity" value={balanceSheet.equity} /><Metric label="Balance difference" value={balanceSheet.balance_difference} good={Math.abs(toNumber(balanceSheet.balance_difference)) < 0.01} /></div>
      ) : null}

      {tab === "monthly-actuals" ? (
        <DataTable title="Monthly actuals" description="Standardized monthly P&L values used by forecasting." headings={["Month", "Revenue", "Gross profit", "Operating expenses", "EBIT", "Net profit"]} rows={monthly.map((row) => [row.month, formatMoney(row.revenue), formatMoney(row.gross_profit), formatMoney(row.operating_expenses), formatMoney(row.ebit), formatMoney(row.net_profit)])} isLoading={isLoading} />
      ) : tab === "branch-comparison" ? (
        <DataTable title="Branch comparison" description="Side-by-side operating performance for branches with activity." headings={["Branch", "Revenue", "Gross profit", "EBIT", "Net profit", "Net margin"]} rows={comparison.map((row) => [`${row.branch_code} — ${row.branch_name}`, formatMoney(row.revenue), formatMoney(row.gross_profit), formatMoney(row.ebit), formatMoney(row.net_profit), row.net_margin_percent == null ? "—" : formatPercent(row.net_margin_percent)])} isLoading={isLoading} />
      ) : tab === "transactions" ? (
        <Card>
          <CardHeader>
            <div className="flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
              <div><CardTitle>Ledger transactions</CardTitle><CardDescription>Trace report evidence to active validated source ledger lines.</CardDescription></div>
              <div className="flex gap-2"><Input value={accountCode} onChange={(event) => setAccountCode(event.target.value)} placeholder="Account code (optional)" className="w-52" /><Button variant="outline" onClick={() => { syncUrl("transactions", accountCode); void load(); }}><Search className="size-4" />Filter</Button></div>
            </div>
          </CardHeader>
          <CardContent>
            {isLoading ? <Loading /> : !transactions.length ? <Empty text="No ledger lines match these filters." /> : (
              <div className="overflow-x-auto"><table className="w-full min-w-[980px] text-sm"><thead><tr className="border-b text-left text-muted-foreground">{["Date", "Account", "Description", "Document", "Debit", "Credit", "Source ref"].map((heading) => <th key={heading} className="px-3 py-3 font-medium">{heading}</th>)}</tr></thead><tbody>{transactions.map((row) => <tr key={row.id} className="border-b"><td className="px-3 py-3">{row.transaction_date}</td><td className="px-3 py-3"><p className="font-mono text-xs">{row.source_account_code}</p><p className="text-xs text-muted-foreground">{row.source_account_name ?? "—"}</p></td><td className="max-w-80 px-3 py-3">{row.description ?? "—"}</td><td className="px-3 py-3">{row.document_number ?? "—"}</td><td className="px-3 py-3 text-right tabular-nums">{formatMoney(row.debit)}</td><td className="px-3 py-3 text-right tabular-nums">{formatMoney(row.credit)}</td><td className="px-3 py-3 font-mono text-xs">{row.external_reference ?? "—"}</td></tr>)}</tbody></table></div>
            )}
          </CardContent>
        </Card>
      ) : (
        <Card>
          <CardHeader>
            <div className="flex items-start justify-between gap-3">
              <div><CardTitle>{tabLabel(tab)}</CardTitle><CardDescription>{branchId ? "Selected branch report." : "Consolidated company report."}</CardDescription></div>
              {tab !== "trial-balance" ? <Button variant="outline" onClick={() => chooseTab("transactions")}>View ledger evidence<ArrowRight className="size-4" /></Button> : null}
            </div>
          </CardHeader>
          <CardContent>
            {isLoading ? <Loading /> : !lines.length ? <Empty text="No report lines found for the selected filters." /> : (
              <div className="overflow-x-auto"><table className="w-full min-w-[680px] text-sm"><thead><tr className="border-b text-left text-muted-foreground"><th className="px-3 py-3 font-medium">Code</th><th className="px-3 py-3 font-medium">Account / line</th><th className="px-3 py-3 text-right font-medium">Amount</th><th className="px-3 py-3 text-right font-medium">Evidence</th></tr></thead><tbody>{lines.map((line, index) => <tr key={`${line.code}-${index}`} className={line.is_total ? "border-b bg-muted/30 font-semibold" : "border-b"}><td className="px-3 py-3 font-mono text-xs">{line.code}</td><td className="px-3 py-3">{line.label}</td><td className="px-3 py-3 text-right tabular-nums">{formatMoney(line.amount)}</td><td className="px-3 py-3 text-right">{tab === "trial-balance" && !line.is_total ? <button type="button" onClick={() => void openAccountTransactions(line.code)} className="text-xs font-semibold text-indigo-600 hover:underline dark:text-indigo-300">Transactions</button> : "—"}</td></tr>)}</tbody></table></div>
            )}
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function Metric({ label, value, good }: { label: string; value: string | number; good?: boolean }) {
  return <Card><CardHeader className="pb-2"><CardDescription>{label}</CardDescription><CardTitle className={good === false ? "text-destructive" : good ? "text-emerald-600" : ""}>{formatMoney(value)}</CardTitle></CardHeader></Card>;
}

function DataTable({ title, description, headings, rows, isLoading }: { title: string; description: string; headings: string[]; rows: string[][]; isLoading: boolean }) {
  return <Card><CardHeader><CardTitle>{title}</CardTitle><CardDescription>{description}</CardDescription></CardHeader><CardContent>{isLoading ? <Loading /> : rows.length === 0 ? <Empty text="No records found." /> : <div className="overflow-x-auto"><table className="w-full min-w-[760px] text-sm"><thead><tr className="border-b text-left text-muted-foreground">{headings.map((heading) => <th key={heading} className="px-3 py-3 font-medium">{heading}</th>)}</tr></thead><tbody>{rows.map((row, rowIndex) => <tr key={`${title}-${rowIndex}`} className="border-b">{row.map((value, columnIndex) => <td key={`${rowIndex}-${columnIndex}`} className="px-3 py-3">{value}</td>)}</tr>)}</tbody></table></div>}</CardContent></Card>;
}

function Loading() { return <div className="flex min-h-52 items-center justify-center gap-2 text-muted-foreground"><Loader2 className="size-5 animate-spin" />Loading…</div>; }
function Empty({ text }: { text: string }) { return <div className="flex min-h-52 items-center justify-center text-center text-muted-foreground">{text}</div>; }
