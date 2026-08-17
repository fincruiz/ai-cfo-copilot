"use client";
import { useEffect, useState } from "react";
import { Activity, CheckCircle2, CircleAlert, Copy, Database, LifeBuoy, RefreshCcw, ShieldCheck } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { systemHealthService, type Readiness } from "@/services/system-health-service";
import { workspaceService, type WorkspaceStatus } from "@/services/workspace-service";

export default function SupportPage(){
 const [health,setHealth]=useState<Readiness|null>(null),[workspace,setWorkspace]=useState<WorkspaceStatus|null>(null),[loading,setLoading]=useState(true),[error,setError]=useState("");
 async function refresh(){setLoading(true);setError("");try{const [h,w]=await Promise.all([systemHealthService.readiness(),workspaceService.getStatus()]);setHealth(h);setWorkspace(w)}catch{setError("FinCruiz could not complete all diagnostics. Retry once; if the problem continues, share the diagnostic summary with support.")}finally{setLoading(false)}}
 useEffect(()=>{void refresh()},[]);
 const healthy=health?.status==="healthy";
 function copySummary(){const summary=`FinCruiz diagnostics\nPlatform: ${health?.status??"unknown"}\nAPI: ${health?.checks.api.status??"unknown"}\nDatabase: ${health?.checks.database?.status??"unknown"} (${health?.checks.database?.latency_ms??0} ms)\nWorkspace transactions: ${workspace?.transaction_count??0}\nMappings: ${workspace?.mapping_count??0}\nVersion: ${health?.version??"unknown"}`;void navigator.clipboard.writeText(summary)}
 return <div className="mx-auto max-w-5xl space-y-6">
  <div className="flex flex-wrap items-end justify-between gap-3"><div><p className="text-sm text-muted-foreground">Launch reliability</p><h1 className="text-3xl font-semibold">Support & diagnostics</h1><p className="mt-2 max-w-2xl text-sm text-muted-foreground">Check workspace and platform health before raising a support request. No financial values or AI prompts are shown here.</p></div><div className="flex gap-2"><Button variant="outline" onClick={copySummary} disabled={!health}><Copy className="size-4"/>Copy diagnostic summary</Button><Button variant="outline" onClick={refresh} disabled={loading}><RefreshCcw className={`size-4 ${loading?"animate-spin":""}`}/>Refresh checks</Button></div></div>
  {error?<div className="rounded-2xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-950">{error}</div>:null}
  <div className="grid gap-4 md:grid-cols-3">
   <Card><CardHeader><CardDescription>API</CardDescription><CardTitle className="flex items-center gap-2"><Activity className="size-5"/>{health?.checks.api.status??"Checking…"}</CardTitle></CardHeader></Card>
   <Card><CardHeader><CardDescription>Database</CardDescription><CardTitle className="flex items-center gap-2"><Database className="size-5"/>{health?.checks.database?.status??"Checking…"}</CardTitle></CardHeader><CardContent className="text-sm text-muted-foreground">{health?.checks.database?`${health.checks.database.latency_ms.toFixed(2)} ms health query`:"Waiting for diagnostic"}</CardContent></Card>
   <Card><CardHeader><CardDescription>Workspace data</CardDescription><CardTitle>{workspace?workspace.transaction_count.toLocaleString():"—"} transactions</CardTitle></CardHeader><CardContent className="text-sm text-muted-foreground">{workspace?.mapping_count??0} mappings · {workspace?.upload_count??0} uploads</CardContent></Card>
  </div>
  <Card className={healthy?"border-emerald-200":"border-amber-200"}><CardHeader><CardTitle className="flex items-center gap-2">{healthy?<CheckCircle2 className="size-5 text-emerald-600"/>:<CircleAlert className="size-5 text-amber-600"/>}Platform readiness</CardTitle><CardDescription>{healthy?"Core services are responding normally.":"One or more core checks need attention."}</CardDescription></CardHeader><CardContent className="grid gap-3 sm:grid-cols-3"><div className="rounded-xl bg-muted/40 p-4"><ShieldCheck className="size-4"/><p className="mt-2 font-medium">Privacy-safe diagnostics</p><p className="mt-1 text-xs text-muted-foreground">No ledger amounts or question text.</p></div><div className="rounded-xl bg-muted/40 p-4"><LifeBuoy className="size-4"/><p className="mt-2 font-medium">Before contacting support</p><p className="mt-1 text-xs text-muted-foreground">Refresh once and note which check is degraded.</p></div><div className="rounded-xl bg-muted/40 p-4"><Activity className="size-4"/><p className="mt-2 font-medium">Version</p><p className="mt-1 text-xs text-muted-foreground">{health?.version??"—"} · {health?.environment??"—"}</p></div></CardContent></Card>
 </div>
}
