"use client";

import Link from "next/link";
import { useEffect, useMemo, useState } from "react";
import { BarChart3, BrainCircuit, Loader2, RefreshCw, Sparkles } from "lucide-react";
import { InsightChart } from "@/components/insight-chart";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { analyticsService } from "@/services/analytics-service";
import { marketService } from "@/services/market-service";
import { getApiErrorMessage } from "@/lib/api";
import type { AIVisualization, AnalyticsOverview } from "@/types/analytics";

export default function VisualBIPage(){
  const [data,setData]=useState<AnalyticsOverview|null>(null); const [currency,setCurrency]=useState("AUD"); const [loading,setLoading]=useState(true); const [error,setError]=useState("");
  async function load(){setLoading(true);setError("");try{const [overview,market]=await Promise.all([analyticsService.getOverview(),marketService.current().catch(()=>null)]);setData(overview);if(market?.currency_code)setCurrency(market.currency_code)}catch(e){setError(getApiErrorMessage(e))}finally{setLoading(false)}}
  useEffect(()=>{void load()},[]);
  const charts=useMemo(()=>buildCharts(data,currency),[data,currency]);
  if(loading)return <div className="flex min-h-[500px] items-center justify-center gap-2 text-muted-foreground"><Loader2 className="size-5 animate-spin"/>Building visual intelligence...</div>;
  return <div className="mx-auto max-w-7xl space-y-6 pb-12">
    <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between"><div><p className="text-sm text-muted-foreground">Finance & performance</p><h1 className="mt-1 text-3xl font-semibold">Visual BI</h1><p className="mt-2 max-w-3xl text-muted-foreground">Ready-to-use graphs from the data already loaded into FinCruiz. You do not need to build a BI report first.</p></div><div className="flex gap-2"><Button variant="outline" onClick={()=>void load()}><RefreshCw className="size-4"/>Refresh</Button><Link href="/dashboard"><Button><BrainCircuit className="size-4"/>Ask FinCruiz</Button></Link></div></div>
    {error?<Alert variant="destructive"><AlertDescription>{error}</AlertDescription></Alert>:null}
    <Card className="border-indigo-200 bg-gradient-to-r from-indigo-50/70 via-background to-sky-50/60 dark:border-indigo-900 dark:from-indigo-950/20 dark:to-sky-950/10"><CardContent className="flex flex-col gap-4 p-5 md:flex-row md:items-center md:justify-between"><div><p className="font-semibold">Two ways to use BI</p><p className="mt-1 text-sm text-muted-foreground">Open these standard views, or ask a question such as “show revenue trend”, “compare branches” or “show receivables ageing”. FinCruiz chooses a suitable chart using grounded company data.</p></div><div className="flex shrink-0 items-center gap-2 rounded-2xl border bg-background px-4 py-3 text-sm"><Sparkles className="size-4 text-indigo-600"/>Graphs update from current data</div></CardContent></Card>
    {!charts.length?<Card><CardContent className="flex min-h-64 items-center justify-center text-center text-sm text-muted-foreground">Load and map finance data to activate Visual BI.</CardContent></Card>:<div className="grid gap-5 xl:grid-cols-2">{charts.map((chart)=><InsightChart key={chart.title} visualization={chart}/>)}</div>}
  </div>
}

function buildCharts(data:AnalyticsOverview|null,currency:string):AIVisualization[]{if(!data)return[];const out:AIVisualization[]=[];const monthly=(data.monthly_actuals??[]).slice(-12);if(monthly.length){out.push({type:"line",title:"Revenue & net profit trend",subtitle:"Last 12 available reporting months",labels:monthly.map(r=>String(r.month).slice(0,7)),series:[{name:"Revenue",data:monthly.map(r=>Number(r.revenue||0))},{name:"Net profit",data:monthly.map(r=>Number(r.net_profit||0))}],value_format:"currency",currency});out.push({type:"area",title:"Gross margin trend",subtitle:"Revenue retained after cost of sales",labels:monthly.map(r=>String(r.month).slice(0,7)),series:[{name:"Gross margin %",data:monthly.map(r=>{const rev=Number(r.revenue||0);const gp=Number(r.gross_profit||0);return rev?gp/rev*100:0})}],value_format:"percent"});}
  if(data.branch_comparison?.length){const rows=data.branch_comparison.slice(0,12);out.push({type:"bar",title:"Branch performance",subtitle:"Revenue and net profit by branch",labels:rows.map(r=>String(r.branch_name||r.branch_code||"Branch")),series:[{name:"Revenue",data:rows.map(r=>Number(r.revenue||0))},{name:"Net profit",data:rows.map(r=>Number(r.net_profit||0))}],value_format:"currency",currency});}
  if(data.ar_summary?.buckets?.length)out.push({type:"donut",title:"Receivables ageing",subtitle:"Where customer cash is currently tied up",labels:data.ar_summary.buckets.map(b=>b.bucket),series:[{name:"Outstanding",data:data.ar_summary.buckets.map(b=>Number(b.amount||0))}],value_format:"currency",currency});
  if(data.ap_summary?.buckets?.length)out.push({type:"stacked_bar",title:"Payables ageing",subtitle:"Supplier obligations by ageing bucket",labels:data.ap_summary.buckets.map(b=>b.bucket),series:[{name:"Payables",data:data.ap_summary.buckets.map(b=>Number(b.amount||0))}],value_format:"currency",currency});return out;}
