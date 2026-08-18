"use client";

import { useEffect, useState } from "react";
import {
  Activity,
  AlertTriangle,
  CheckCircle2,
  CircleAlert,
  Copy,
  Database,
  HardDrive,
  LifeBuoy,
  Loader2,
  RefreshCcw,
  ShieldCheck,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  systemHealthService,
  type OperationalReadiness,
  type Readiness,
} from "@/services/system-health-service";
import {
  workspaceService,
  type WorkspaceStatus,
} from "@/services/workspace-service";

function tone(status: string) {
  if (status === "healthy") return "border-emerald-200 bg-emerald-50";
  if (status === "degraded") return "border-amber-200 bg-amber-50";
  return "border-red-200 bg-red-50";
}

export default function SupportPage() {
  const [health, setHealth] = useState<Readiness | null>(null);
  const [operations, setOperations] = useState<OperationalReadiness | null>(null);
  const [workspace, setWorkspace] = useState<WorkspaceStatus | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  async function refresh() {
    setLoading(true);
    setError("");

    try {
      const [platform, operational, workspaceStatus] = await Promise.all([
        systemHealthService.readiness(),
        systemHealthService.operations(),
        workspaceService.getStatus(),
      ]);
      setHealth(platform);
      setOperations(operational);
      setWorkspace(workspaceStatus);
    } catch {
      setError(
        "FinCruiz could not complete all diagnostics. Retry once; if the problem continues, share the diagnostic summary with support.",
      );
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    void refresh();
  }, []);

  function copySummary() {
    const summary = [
      "FinCruiz diagnostics",
      `Platform: ${health?.status ?? "unknown"}`,
      `API: ${health?.checks.api.status ?? "unknown"}`,
      `Database: ${health?.checks.database?.status ?? "unknown"} (${health?.checks.database?.latency_ms ?? 0} ms)`,
      `Operations: ${operations?.status ?? "unknown"} (${operations?.score ?? 0}%)`,
      `Ingestion: ${operations?.ingestion_open_jobs ?? 0} open / ${operations?.ingestion_stale_jobs ?? 0} stale / ${operations?.ingestion_recent_failures ?? 0} recent failures`,
      `Active GL datasets: ${operations?.active_gl_datasets ?? 0}`,
      `Workspace transactions: ${workspace?.transaction_count ?? 0}`,
      `Mappings: ${workspace?.mapping_count ?? 0}`,
      `Version: ${health?.version ?? "unknown"}`,
    ].join("\n");

    void navigator.clipboard.writeText(summary);
  }

  const combined =
    health?.status === "unhealthy" || operations?.status === "unhealthy"
      ? "unhealthy"
      : health?.status === "degraded" || operations?.status === "degraded"
        ? "degraded"
        : "healthy";

  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <p className="text-sm text-muted-foreground">
            Performance & failure recovery
          </p>
          <h1 className="text-3xl font-semibold">Support & diagnostics</h1>
          <p className="mt-2 max-w-3xl text-sm text-muted-foreground">
            Verify platform response, database latency, ingestion health,
            staging storage and launch indexes without exposing ledger values,
            credentials or AI prompts.
          </p>
        </div>

        <div className="flex gap-2">
          <Button variant="outline" onClick={copySummary} disabled={!health}>
            <Copy className="size-4" />
            Copy diagnostic summary
          </Button>
          <Button variant="outline" onClick={() => void refresh()} disabled={loading}>
            <RefreshCcw className={`size-4 ${loading ? "animate-spin" : ""}`} />
            Refresh checks
          </Button>
        </div>
      </div>

      {error ? (
        <div className="rounded-2xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-950">
          {error}
        </div>
      ) : null}

      <Card className={tone(combined)}>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            {loading ? (
              <Loader2 className="size-5 animate-spin" />
            ) : combined === "healthy" ? (
              <CheckCircle2 className="size-5 text-emerald-600" />
            ) : combined === "degraded" ? (
              <AlertTriangle className="size-5 text-amber-600" />
            ) : (
              <CircleAlert className="size-5 text-red-600" />
            )}
            Operational readiness
          </CardTitle>
          <CardDescription>
            {combined === "healthy"
              ? "Core services and workspace operations are responding normally."
              : combined === "degraded"
                ? "The workspace can operate, but one or more launch checks need attention."
                : "A launch-critical operational issue requires investigation."}
          </CardDescription>
        </CardHeader>
        <CardContent>
          <p className="text-3xl font-semibold">
            {operations?.score ?? "—"}
            {operations ? "%" : ""}
          </p>
        </CardContent>
      </Card>

      <div className="grid gap-4 md:grid-cols-4">
        <Card>
          <CardHeader>
            <CardDescription>API</CardDescription>
            <CardTitle className="flex items-center gap-2">
              <Activity className="size-5" />
              {health?.checks.api.status ?? "Checking…"}
            </CardTitle>
          </CardHeader>
        </Card>

        <Card>
          <CardHeader>
            <CardDescription>Database</CardDescription>
            <CardTitle className="flex items-center gap-2">
              <Database className="size-5" />
              {health?.checks.database?.status ?? "Checking…"}
            </CardTitle>
          </CardHeader>
          <CardContent className="text-sm text-muted-foreground">
            {health?.checks.database
              ? `${health.checks.database.latency_ms.toFixed(2)} ms`
              : "Waiting for diagnostic"}
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardDescription>Ingestion queue</CardDescription>
            <CardTitle>{operations?.ingestion_open_jobs ?? "—"} open</CardTitle>
          </CardHeader>
          <CardContent className="text-sm text-muted-foreground">
            {operations
              ? `${operations.ingestion_stale_jobs} stale · ${operations.ingestion_recent_failures} recent failures`
              : "Waiting for diagnostic"}
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardDescription>Workspace data</CardDescription>
            <CardTitle>
              {workspace ? workspace.transaction_count.toLocaleString() : "—"} transactions
            </CardTitle>
          </CardHeader>
          <CardContent className="text-sm text-muted-foreground">
            {workspace?.mapping_count ?? 0} mappings ·{" "}
            {workspace?.upload_count ?? 0} uploads
          </CardContent>
        </Card>
      </div>

      {operations ? (
        <Card>
          <CardHeader>
            <CardTitle>Operational checks</CardTitle>
            <CardDescription>
              These diagnostics are read-only. FinCruiz does not automatically
              retry a stale import because doing so blindly could duplicate work.
            </CardDescription>
          </CardHeader>
          <CardContent className="grid gap-3">
            {operations.checks.map((check) => (
              <div
                key={check.key}
                className="rounded-2xl border bg-background p-4"
              >
                <div className="flex items-start gap-3">
                  {check.key === "staging_storage" ? (
                    <HardDrive className="mt-0.5 size-5 text-muted-foreground" />
                  ) : check.status === "healthy" ? (
                    <CheckCircle2 className="mt-0.5 size-5 text-emerald-600" />
                  ) : (
                    <AlertTriangle className="mt-0.5 size-5 text-amber-600" />
                  )}
                  <div>
                    <p className="font-semibold">{check.label}</p>
                    <p className="mt-1 text-sm text-muted-foreground">
                      {check.detail}
                    </p>
                    {check.action ? (
                      <p className="mt-2 text-sm font-medium">
                        Action: {check.action}
                      </p>
                    ) : null}
                  </div>
                </div>
              </div>
            ))}
          </CardContent>
        </Card>
      ) : null}

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <ShieldCheck className="size-5" />
            Recovery principles
          </CardTitle>
        </CardHeader>
        <CardContent className="grid gap-3 sm:grid-cols-3">
          <div className="rounded-xl bg-muted/40 p-4">
            <LifeBuoy className="size-4" />
            <p className="mt-2 font-medium">Preserve the last good dataset</p>
            <p className="mt-1 text-xs text-muted-foreground">
              Failed or incomplete imports must never silently replace the
              previously validated active ledger.
            </p>
          </div>
          <div className="rounded-xl bg-muted/40 p-4">
            <Database className="size-4" />
            <p className="mt-2 font-medium">Measure before migrating</p>
            <p className="mt-1 text-xs text-muted-foreground">
              Database latency and load-test results should drive scaling
              decisions rather than assumptions about PostgreSQL/Supabase.
            </p>
          </div>
          <div className="rounded-xl bg-muted/40 p-4">
            <Activity className="size-4" />
            <p className="mt-2 font-medium">Trace slow requests</p>
            <p className="mt-1 text-xs text-muted-foreground">
              Every API response carries a request ID and server timing; slow
              requests are promoted to warning-level logs.
            </p>
          </div>
        </CardContent>
      </Card>

      <p className="text-xs text-muted-foreground">
        Version {health?.version ?? "—"} · {health?.environment ?? "—"}
      </p>
    </div>
  );
}
