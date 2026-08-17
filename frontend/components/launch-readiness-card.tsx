"use client";
import Link from "next/link";
import { ArrowRight, CheckCircle2, Circle } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import type { LaunchReadiness } from "@/services/workspace-service";

export function LaunchReadinessCard({ readiness }: { readiness: LaunchReadiness }) {
  if (readiness.score >= 100) return null;
  return <Card className="overflow-hidden border-primary/20 bg-gradient-to-r from-primary/[.06] via-background to-sky-500/[.04]"><CardHeader><div className="flex flex-wrap items-start justify-between gap-3"><div><CardTitle>Get FinCruiz management-ready</CardTitle><CardDescription>{readiness.completed_steps} of {readiness.total_steps} essentials complete. FinCruiz uses this setup to keep reports, BI and AI grounded.</CardDescription></div><div className="rounded-full border bg-background px-3 py-1 text-sm font-semibold">{readiness.score}% ready</div></div><div className="mt-3 h-2 overflow-hidden rounded-full bg-muted"><div className="h-full rounded-full bg-primary transition-all" style={{width:`${readiness.score}%`}}/></div></CardHeader><CardContent><div className="grid gap-2 md:grid-cols-4">{readiness.checks.map(check=><Link key={check.key} href={check.path} className="rounded-xl border bg-background/80 p-3 transition hover:-translate-y-0.5 hover:shadow-sm"><div className="flex items-center gap-2 text-sm font-medium">{check.ready?<CheckCircle2 className="size-4 text-emerald-600"/>:<Circle className="size-4 text-muted-foreground"/>}{check.label}</div><p className="mt-2 text-xs leading-5 text-muted-foreground">{check.detail}</p></Link>)}</div><div className="mt-4 flex justify-end"><Button asChild><Link href={readiness.next_path}>{readiness.next_label}<ArrowRight className="size-4"/></Link></Button></div></CardContent></Card>;
}
