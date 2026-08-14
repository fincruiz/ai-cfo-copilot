"use client";

import { useEffect, useState } from "react";
import { Loader2, RefreshCw } from "lucide-react";

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
  MonthlyActual,
  ProfitAndLoss,
  TrialBalance,
} from "@/types/finance";

const tabs = [
  "trial-balance",
  "profit-and-loss",
  "balance-sheet",
  "monthly-actuals",
  "branch-comparison",
] as const;
type Tab = typeof tabs[number];

export default function ReportsPage() {
  const [tab, setTab] = useState<Tab>("trial-balance");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [branchId, setBranchId] = useState("");
  const [branches, setBranches] = useState<Branch[]>([]);
  const [trialBalance, setTrialBalance] = useState<TrialBalance | null>(null);
  const [pnl, setPnl] = useState<ProfitAndLoss | null>(null);
  const [balanceSheet, setBalanceSheet] = useState<BalanceSheet | null>(null);
  const [monthly, setMonthly] = useState<MonthlyActual[]>([]);
  const [comparison, setComparison] = useState<BranchComparison[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState("");

  async function load() {
    setIsLoading(true);
    setError("");
    try {
      const params = {
        startDate,
        endDate,
        branchId: branchId || undefined,
      };
      const [branchRows, tb, profitLoss, bs, monthlyRows, comparisonRows] =
        await Promise.all([
          financeService.getBranches(),
          financeService.getTrialBalance(params),
          financeService.getProfitAndLoss(params),
          financeService.getBalanceSheet({
            endDate,
            branchId: branchId || undefined,
          }),
          financeService.getMonthlyActuals(params),
          financeService.getBranchComparison({ startDate, endDate }),
        ]);
      setBranches(branchRows);
      setTrialBalance(tb);
      setPnl(profitLoss);
      setBalanceSheet(bs);
      setMonthly(monthlyRows);
      setComparison(comparisonRows);
    } catch (loadError) {
      setError(getApiErrorMessage(loadError));
    } finally {
      setIsLoading(false);
    }
  }

  useEffect(() => { void load(); }, []);

  const current =
    tab === "trial-balance"
      ? trialBalance
      : tab === "profit-and-loss"
        ? pnl
        : tab === "balance-sheet"
          ? balanceSheet
          : null;
  const lines = current?.lines ?? [];

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      <div className="flex flex-col gap-4 xl:flex-row xl:items-end xl:justify-between">
        <div>
          <p className="text-sm font-medium text-muted-foreground">Financial reporting</p>
          <h1 className="mt-1 text-3xl font-semibold tracking-tight">Reports</h1>
          <p className="mt-2 text-muted-foreground">
            Consolidated and branch-level reports from the active ledger dataset.
          </p>
        </div>

        <div className="flex flex-wrap items-end gap-3">
          <div>
            <p className="mb-1 text-xs text-muted-foreground">View</p>
            <select
              className="h-10 min-w-52 rounded-md border bg-background px-3 text-sm"
              value={branchId}
              onChange={(event) => setBranchId(event.target.value)}
            >
              <option value="">Consolidated company</option>
              {branches.filter((branch) => branch.is_active).map((branch) => (
                <option key={branch.id} value={branch.id}>
                  {branch.branch_code} — {branch.branch_name}
                </option>
              ))}
            </select>
          </div>
          <div>
            <p className="mb-1 text-xs text-muted-foreground">Start date</p>
            <Input type="date" value={startDate} onChange={(event) => setStartDate(event.target.value)} />
          </div>
          <div>
            <p className="mb-1 text-xs text-muted-foreground">End date</p>
            <Input type="date" value={endDate} onChange={(event) => setEndDate(event.target.value)} />
          </div>
          <Button onClick={() => void load()} disabled={isLoading}>
            <RefreshCw className="size-4" />Refresh
          </Button>
        </div>
      </div>

      {error ? <Alert variant="destructive"><AlertDescription>{error}</AlertDescription></Alert> : null}

      <div className="flex flex-wrap gap-2">
        {tabs.map((item) => (
          <Button
            key={item}
            variant={tab === item ? "default" : "outline"}
            onClick={() => setTab(item)}
          >
            {{
              "trial-balance": "Trial balance",
              "profit-and-loss": "Profit & loss",
              "balance-sheet": "Balance sheet",
              "monthly-actuals": "Monthly actuals",
              "branch-comparison": "Branch comparison",
            }[item]}
          </Button>
        ))}
      </div>

      {tab === "trial-balance" && trialBalance ? (
        <div className="grid gap-4 sm:grid-cols-3">
          <Metric label="Total debit" value={trialBalance.total_debit} />
          <Metric label="Total credit" value={trialBalance.total_credit} />
          <Metric
            label="Difference"
            value={trialBalance.difference}
            good={Math.abs(toNumber(trialBalance.difference)) < 0.01}
          />
        </div>
      ) : null}

      {tab === "profit-and-loss" && pnl ? (
        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          <Metric label="Revenue" value={pnl.revenue} />
          <Metric label="Gross profit" value={pnl.gross_profit} />
          <Metric label="Operating profit" value={pnl.operating_profit} />
          <Metric label="Net profit" value={pnl.net_profit} />
        </div>
      ) : null}

      {tab === "balance-sheet" && balanceSheet ? (
        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          <Metric label="Total assets" value={balanceSheet.total_assets} />
          <Metric label="Total liabilities" value={balanceSheet.total_liabilities} />
          <Metric label="Equity" value={balanceSheet.equity} />
          <Metric
            label="Balance difference"
            value={balanceSheet.balance_difference}
            good={Math.abs(toNumber(balanceSheet.balance_difference)) < 0.01}
          />
        </div>
      ) : null}

      {tab === "monthly-actuals" ? (
        <DataTable
          title="Monthly actuals"
          description="Standardized monthly P&L values used by forecasting."
          headings={["Month", "Revenue", "Gross profit", "Operating expenses", "EBIT", "Net profit"]}
          rows={monthly.map((row) => [
            row.month,
            formatMoney(row.revenue),
            formatMoney(row.gross_profit),
            formatMoney(row.operating_expenses),
            formatMoney(row.ebit),
            formatMoney(row.net_profit),
          ])}
          isLoading={isLoading}
        />
      ) : tab === "branch-comparison" ? (
        <DataTable
          title="Branch comparison"
          description="Side-by-side operating performance for branches with activity."
          headings={["Branch", "Revenue", "Gross profit", "EBIT", "Net profit", "Net margin"]}
          rows={comparison.map((row) => [
            `${row.branch_code} — ${row.branch_name}`,
            formatMoney(row.revenue),
            formatMoney(row.gross_profit),
            formatMoney(row.ebit),
            formatMoney(row.net_profit),
            row.net_margin_percent == null ? "—" : formatPercent(row.net_margin_percent),
          ])}
          isLoading={isLoading}
        />
      ) : (
        <Card>
          <CardHeader>
            <CardTitle>
              {tab === "trial-balance"
                ? "Trial balance"
                : tab === "profit-and-loss"
                  ? "Profit and loss"
                  : "Balance sheet"}
            </CardTitle>
            <CardDescription>
              {branchId ? "Selected branch report." : "Consolidated company report."}
            </CardDescription>
          </CardHeader>
          <CardContent>
            {isLoading ? (
              <div className="flex min-h-64 items-center justify-center gap-2 text-muted-foreground">
                <Loader2 className="size-5 animate-spin" />Generating reports...
              </div>
            ) : !lines.length ? (
              <div className="flex min-h-64 items-center justify-center text-center text-muted-foreground">
                No report lines found for the selected filters.
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full min-w-[680px] text-sm">
                  <thead>
                    <tr className="border-b text-left text-muted-foreground">
                      <th className="px-3 py-3 font-medium">Code</th>
                      <th className="px-3 py-3 font-medium">Account / line</th>
                      <th className="px-3 py-3 text-right font-medium">Amount</th>
                    </tr>
                  </thead>
                  <tbody>
                    {lines.map((line, index) => (
                      <tr
                        key={`${line.code}-${index}`}
                        className={line.is_total ? "border-b bg-muted/30 font-semibold" : "border-b"}
                      >
                        <td className="px-3 py-3 font-mono text-xs">{line.code}</td>
                        <td className="px-3 py-3">{line.label}</td>
                        <td className="px-3 py-3 text-right tabular-nums">{formatMoney(line.amount)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function Metric({
  label,
  value,
  good,
}: {
  label: string;
  value: string | number;
  good?: boolean;
}) {
  return (
    <Card>
      <CardHeader className="pb-2">
        <CardDescription>{label}</CardDescription>
        <CardTitle className={good === false ? "text-destructive" : good ? "text-emerald-600" : ""}>
          {formatMoney(value)}
        </CardTitle>
      </CardHeader>
    </Card>
  );
}

function DataTable({
  title,
  description,
  headings,
  rows,
  isLoading,
}: {
  title: string;
  description: string;
  headings: string[];
  rows: string[][];
  isLoading: boolean;
}) {
  return (
    <Card>
      <CardHeader>
        <CardTitle>{title}</CardTitle>
        <CardDescription>{description}</CardDescription>
      </CardHeader>
      <CardContent>
        {isLoading ? (
          <div className="flex min-h-52 items-center justify-center gap-2 text-muted-foreground">
            <Loader2 className="size-5 animate-spin" />Loading...
          </div>
        ) : rows.length === 0 ? (
          <div className="flex min-h-52 items-center justify-center text-muted-foreground">
            No records found.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full min-w-[760px] text-sm">
              <thead>
                <tr className="border-b text-left text-muted-foreground">
                  {headings.map((heading) => (
                    <th key={heading} className="px-3 py-3 font-medium">{heading}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, rowIndex) => (
                  <tr key={`${title}-${rowIndex}`} className="border-b">
                    {row.map((value, columnIndex) => (
                      <td key={`${rowIndex}-${columnIndex}`} className="px-3 py-3">{value}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </CardContent>
    </Card>
  );
}
