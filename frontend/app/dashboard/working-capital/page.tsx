"use client";
import { useEffect, useState } from "react";
import { Loader2, Users, WalletCards } from "lucide-react";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { analyticsService } from "@/services/analytics-service";
import { formatMoney, toNumber } from "@/lib/finance-format";
import type { WorkingCapitalSummary } from "@/types/analytics";
import { ModuleResetButton } from "@/components/module-reset-button";

export default function WorkingCapitalPage() {
  const [ar,setAr]=useState<WorkingCapitalSummary|null>(null);
  const [ap,setAp]=useState<WorkingCapitalSummary|null>(null);
  const [loading,setLoading]=useState(true);
  useEffect(()=>{Promise.all([analyticsService.getWorkingCapital("AR"),analyticsService.getWorkingCapital("AP")]).then(([a,p])=>{setAr(a);setAp(p)}).finally(()=>setLoading(false))},[]);
  if(loading)return <div className="flex min-h-[500px] items-center justify-center"><Loader2 className="size-5 animate-spin"/></div>;
  return <div className="mx-auto max-w-7xl space-y-7">
    <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between"><div><p className="text-sm font-medium text-muted-foreground">Working capital</p><h1 className="mt-1 text-3xl font-semibold">Customer Collections & Vendor Payments</h1><p className="mt-2 text-muted-foreground">Invoice ageing, exposure concentration and payment-cycle analysis.</p></div><div className="flex gap-2"><ModuleResetButton scope="ar_ageing" label="Reset AR" description="Remove only Accounts Receivable ageing data."/><ModuleResetButton scope="ap_ageing" label="Reset AP" description="Remove only Accounts Payable ageing data."/></div></div>
    <div className="grid gap-5 lg:grid-cols-2"><Panel title="Accounts receivable" icon={Users} data={ar}/><Panel title="Accounts payable" icon={WalletCards} data={ap}/></div>
  </div>
}
function Panel({title,icon:Icon,data}:{title:string;icon:typeof Users;data:WorkingCapitalSummary|null}) {
  return <Card><CardHeader><CardTitle className="flex items-center gap-2"><Icon className="size-5"/>{title}</CardTitle><CardDescription>{data?`${data.document_count} invoices across ${data.party_count} parties`:"No ageing uploaded"}</CardDescription></CardHeader><CardContent>
    {!data?<div className="flex min-h-48 items-center justify-center text-muted-foreground">Upload the ageing report in Import Centre.</div>:<div className="space-y-5">
      <div className="grid grid-cols-3 gap-3">{[["Outstanding",formatMoney(data.total_outstanding)],["Overdue",formatMoney(data.overdue_amount)],["Overdue %",`${toNumber(data.overdue_percent).toFixed(1)}%`]].map(([l,v])=><div key={l} className="rounded-xl border p-3"><p className="text-xs text-muted-foreground">{l}</p><p className="mt-1 font-semibold">{v}</p></div>)}</div>
      <div className="space-y-2">{data.buckets.map(b=><div key={b.bucket} className="flex justify-between rounded-xl bg-muted/40 p-3 text-sm"><span>{b.bucket}</span><b>{formatMoney(b.amount)}</b></div>)}</div>
      <div>{data.top_parties.slice(0,8).map(p=><div key={p.party_name} className="flex justify-between border-b py-3 text-sm"><span>{p.party_name}</span><b>{formatMoney(p.outstanding_amount)}</b></div>)}</div>
    </div>}
  </CardContent></Card>
}
