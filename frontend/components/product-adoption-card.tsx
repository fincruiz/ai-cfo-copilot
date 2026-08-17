"use client";

import { useEffect, useMemo, useState } from "react";
import { Activity, BarChart3, Eye, Loader2, ShieldCheck } from "lucide-react";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { companyService } from "@/services/company-service";
import { usageService } from "@/services/usage-service";

type Row = { event_name: string; count: number; users: number };

function readable(event: string) {
  return event.replaceAll("_", " ").replace(/\b\w/g, (x) => x.toUpperCase());
}

export function ProductAdoptionCard() {
  const [rows, setRows] = useState<Row[] | null>(null);
  const [allowed, setAllowed] = useState(false);

  useEffect(() => {
    companyService.getAccess().then(async (access) => {
      if (!access.can_manage_members) return setRows([]);
      setAllowed(true);
      try { setRows(await usageService.summary(30)); } catch { setRows([]); }
    }).catch(() => setRows([]));
  }, []);

  const total = useMemo(() => (rows || []).reduce((sum, row) => sum + row.count, 0), [rows]);
  if (rows === null) return <Card><CardContent className="flex min-h-28 items-center justify-center"><Loader2 className="size-5 animate-spin" /></CardContent></Card>;
  if (!allowed) return null;

  return <Card>
    <CardHeader><div className="flex items-start gap-3"><div className="rounded-xl bg-emerald-500/10 p-2.5 text-emerald-700"><Activity className="size-5" /></div><div><CardTitle>Product adoption · last 30 days</CardTitle><CardDescription className="mt-1">Privacy-safe behavioural telemetry only. FinCruiz does not put financial values, ERP payloads or AI question text into this usage summary.</CardDescription></div></div></CardHeader>
    <CardContent>
      <div className="mb-4 grid gap-3 sm:grid-cols-2"><div className="rounded-xl border bg-muted/20 p-4"><div className="flex items-center gap-2 text-xs text-muted-foreground"><Eye className="size-3.5" />Recorded feature interactions</div><p className="mt-2 text-2xl font-semibold">{total.toLocaleString()}</p></div><div className="rounded-xl border bg-muted/20 p-4"><div className="flex items-center gap-2 text-xs text-muted-foreground"><ShieldCheck className="size-3.5" />Privacy boundary</div><p className="mt-2 text-sm font-semibold">Behaviour metadata only</p></div></div>
      {rows.length ? <div className="space-y-2">{rows.slice(0, 8).map((row) => { const width = Math.max(4, (row.count / Math.max(1, rows[0]?.count || 1)) * 100); return <div key={row.event_name} className="grid grid-cols-[minmax(0,1fr)_64px] items-center gap-3"><div><div className="flex justify-between gap-3 text-xs"><span className="truncate">{readable(row.event_name)}</span><span className="text-muted-foreground">{row.users} user(s)</span></div><div className="mt-1.5 h-2 overflow-hidden rounded-full bg-muted"><div className="h-full rounded-full bg-gradient-to-r from-emerald-500 to-sky-500" style={{ width: `${width}%` }} /></div></div><div className="text-right text-sm font-semibold">{row.count}</div></div>; })}</div> : <div className="rounded-xl border border-dashed p-5 text-sm text-muted-foreground"><BarChart3 className="mb-2 size-5" />Usage data will appear as beta customers start navigating the workspace.</div>}
    </CardContent>
  </Card>;
}
