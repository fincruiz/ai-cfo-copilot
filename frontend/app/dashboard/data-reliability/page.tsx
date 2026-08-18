"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import {
  AlertTriangle,
  CheckCircle2,
  FileCheck2,
  Loader2,
  RefreshCw,
  ShieldAlert,
  ShieldCheck,
} from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { getApiErrorMessage } from "@/lib/api";
import { financeService } from "@/services/finance-service";
import type { FinanceReliability } from "@/types/finance";

const categoryLabels: Record<string, string> = {
  ingestion: "Ingestion & source data",
  reconciliation: "Financial reconciliation",
  mapping: "Account mapping",
  traceability: "Source traceability",
  branches: "Branches & consolidation",
  periods: "Periods & recency",
  finance: "Finance",
};

function statusStyle(status: string) {
  if (status === "pass" || status === "ready") {
    return "border-emerald-200 bg-emerald-50 text-emerald-900";
  }
  if (status === "warning" || status === "attention") {
    return "border-amber-200 bg-amber-50 text-amber-900";
  }
  return "border-red-200 bg-red-50 text-red-900";
}

function StatusIcon({ status }: { status: string }) {
  if (status === "pass" || status === "ready") {
    return <CheckCircle2 className="size-5 text-emerald-600" />;
  }
  if (status === "warning" || status === "attention") {
    return <AlertTriangle className="size-5 text-amber-600" />;
  }
  return <ShieldAlert className="size-5 text-red-600" />;
}

export default function DataReliabilityPage() {
  const [result, setResult] = useState<FinanceReliability | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  async function load() {
    setLoading(true);
    setError("");
    try {
      setResult(await financeService.getFinanceReliability());
    } catch (loadError) {
      setError(getApiErrorMessage(loadError));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    void load();
  }, []);

  const grouped = useMemo(() => {
    const value: Record<string, FinanceReliability["checks"]> = {};
    for (const check of result?.checks ?? []) {
      (value[check.category] ??= []).push(check);
    }
    return value;
  }, [result]);

  if (loading && !result) {
    return (
      <div className="flex min-h-[55vh] items-center justify-center gap-3 text-muted-foreground">
        <Loader2 className="size-5 animate-spin" />
        Certifying the active finance dataset…
      </div>
    );
  }

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
        <div>
          <p className="text-sm font-medium text-muted-foreground">
            Data & finance reliability
          </p>
          <h1 className="mt-1 text-3xl font-semibold tracking-tight">
            Finance Reliability Certification
          </h1>
          <p className="mt-2 max-w-3xl text-muted-foreground">
            A deterministic preflight over the active ledger, mappings,
            reconciliations, branches, source lineage, period coverage and
            ingestion state.
          </p>
        </div>

        <Button onClick={() => void load()} disabled={loading}>
          <RefreshCw className={loading ? "size-4 animate-spin" : "size-4"} />
          Re-run certification
        </Button>
      </div>

      {error ? (
        <Alert variant="destructive">
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      ) : null}

      {result ? (
        <>
          <Card className={statusStyle(result.status)}>
            <CardContent className="p-6">
              <div className="flex flex-col gap-5 lg:flex-row lg:items-center lg:justify-between">
                <div className="flex items-start gap-4">
                  <div className="mt-1">
                    <StatusIcon status={result.status} />
                  </div>
                  <div>
                    <p className="text-xs font-bold uppercase tracking-[.16em]">
                      {result.status === "ready"
                        ? "Ready"
                        : result.status === "attention"
                          ? "Attention required"
                          : "Blocked"}
                    </p>
                    <p className="mt-1 text-2xl font-semibold">
                      {result.status === "ready"
                        ? "Active finance data passed launch reliability checks."
                        : result.status === "attention"
                          ? "The dataset is usable, but review the flagged items."
                          : "Resolve blocking finance issues before relying on management outputs."}
                    </p>
                    <p className="mt-2 text-sm opacity-80">
                      Certified {new Date(result.certified_at).toLocaleString()}
                    </p>
                  </div>
                </div>

                <div className="grid grid-cols-4 gap-2 text-center">
                  <Metric label="Score" value={`${result.score}%`} />
                  <Metric label="Pass" value={String(result.pass_count)} />
                  <Metric label="Warnings" value={String(result.warning_count)} />
                  <Metric label="Fails" value={String(result.fail_count)} />
                </div>
              </div>
            </CardContent>
          </Card>

          <div className="grid gap-4 sm:grid-cols-3">
            <Card>
              <CardHeader className="pb-2">
                <CardDescription>Financial assurance</CardDescription>
                <CardTitle>
                  Grade {result.assurance_grade} · {result.assurance_score}%
                </CardTitle>
              </CardHeader>
            </Card>
            <Card>
              <CardHeader className="pb-2">
                <CardDescription>Active ledger coverage</CardDescription>
                <CardTitle className="text-lg">
                  {result.first_transaction_date ?? "—"} →{" "}
                  {result.last_transaction_date ?? "—"}
                </CardTitle>
              </CardHeader>
            </Card>
            <Card>
              <CardHeader className="pb-2">
                <CardDescription>Active upload</CardDescription>
                <CardTitle className="truncate text-base font-mono">
                  {result.active_upload_id ?? "No active upload"}
                </CardTitle>
              </CardHeader>
            </Card>
          </div>

          {Object.entries(grouped).map(([category, checks]) => (
            <Card key={category}>
              <CardHeader>
                <CardTitle>{categoryLabels[category] ?? category}</CardTitle>
                <CardDescription>
                  {checks.filter((item) => item.status === "pass").length} of{" "}
                  {checks.length} checks passed.
                </CardDescription>
              </CardHeader>
              <CardContent className="grid gap-3">
                {checks.map((check) => (
                  <div
                    key={check.key}
                    className="rounded-2xl border bg-background p-4"
                  >
                    <div className="flex items-start gap-3">
                      <StatusIcon status={check.status} />
                      <div className="min-w-0 flex-1">
                        <div className="flex flex-wrap items-center gap-2">
                          <p className="font-semibold">{check.label}</p>
                          {check.blocking && check.status === "fail" ? (
                            <span className="rounded-full bg-red-100 px-2 py-0.5 text-[10px] font-bold uppercase text-red-700">
                              Blocking
                            </span>
                          ) : null}
                        </div>
                        <p className="mt-1 text-sm text-muted-foreground">
                          {check.detail}
                        </p>
                        {check.action ? (
                          <p className="mt-2 text-sm font-medium">
                            Next action: {check.action}
                          </p>
                        ) : null}
                      </div>
                    </div>
                  </div>
                ))}
              </CardContent>
            </Card>
          ))}

          <div className="grid gap-4 lg:grid-cols-3">
            <Link
              href="/dashboard/uploads"
              className="rounded-2xl border bg-background p-5 transition hover:-translate-y-0.5 hover:shadow-md"
            >
              <FileCheck2 className="size-5 text-primary" />
              <p className="mt-3 font-semibold">Review source uploads</p>
              <p className="mt-1 text-sm text-muted-foreground">
                Validate or replace the active General Ledger dataset.
              </p>
            </Link>

            <Link
              href="/dashboard/mapping"
              className="rounded-2xl border bg-background p-5 transition hover:-translate-y-0.5 hover:shadow-md"
            >
              <ShieldCheck className="size-5 text-primary" />
              <p className="mt-3 font-semibold">Review account mapping</p>
              <p className="mt-1 text-sm text-muted-foreground">
                Resolve unmapped or incorrectly classified accounts.
              </p>
            </Link>

            <Link
              href="/dashboard/reports"
              className="rounded-2xl border bg-background p-5 transition hover:-translate-y-0.5 hover:shadow-md"
            >
              <ShieldCheck className="size-5 text-primary" />
              <p className="mt-3 font-semibold">Open financial reports</p>
              <p className="mt-1 text-sm text-muted-foreground">
                Review the certified trial balance, P&L and balance sheet.
              </p>
            </Link>
          </div>

          <p className="text-xs leading-5 text-muted-foreground">
            {result.caveat}
          </p>
        </>
      ) : null}
    </div>
  );
}

function Metric({ label, value }: { label: string; value: string }) {
  return (
    <div className="min-w-20 rounded-xl border border-current/10 bg-white/40 px-3 py-3">
      <p className="text-[10px] font-bold uppercase tracking-wide opacity-70">
        {label}
      </p>
      <p className="mt-1 text-xl font-semibold">{value}</p>
    </div>
  );
}
