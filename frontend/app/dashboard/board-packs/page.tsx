"use client";
import { ModuleResetButton } from "@/components/module-reset-button";
import Link from "next/link";
import { FileBarChart, Presentation, Sparkles } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
export default function BoardpacksPage(){
 return <div className="mx-auto max-w-6xl space-y-7"><div className="animate-rise"><p className="text-sm text-muted-foreground">Board & exports</p><div className="flex items-center justify-between gap-4"><h1 className="mt-1 text-3xl font-semibold">Board Packs</h1><ModuleResetButton scope="board_packs" label="Clear board packs" description="This removes generated board-pack files and saved pack records. Your source finance data remains." /></div><p className="mt-2 text-muted-foreground">Assemble financial statements, KPIs, analytics, forecasts and commentary.</p></div>
 <div className="grid gap-5 md:grid-cols-3">{[
 ["Executive summary","Revenue, profitability, cash, working capital and key risks."],
 ["Performance visuals","Monthly trends, branch comparison and variance charts."],
 ["Management actions","AI-assisted commentary, priorities and follow-up decisions."]
 ].map(([t,d],i)=><Card key={t} className="animate-card-in" style={{animationDelay:`${i*90}ms`}}><CardHeader><Sparkles className="size-5 text-indigo-600"/><CardTitle>{t}</CardTitle><CardDescription>{d}</CardDescription></CardHeader></Card>)}</div>
 <Card><CardHeader><CardTitle>Legacy migration status</CardTitle><CardDescription>The finance data and analytics foundation is now connected. The next export build will convert these sections into saved board-pack versions and downloadable PPTX/PDF files.</CardDescription></CardHeader><CardContent className="flex gap-3"><Link href="/dashboard/analytics"><Button><FileBarChart className="size-4"/>Review analytics</Button></Link><Link href="/dashboard/planning"><Button variant="outline"><Presentation className="size-4"/>Review planning</Button></Link></CardContent></Card></div>
}
