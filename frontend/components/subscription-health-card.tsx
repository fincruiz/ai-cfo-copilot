"use client";

import { useEffect, useState } from "react";
import { CheckCircle2, CircleAlert, Clock3, CreditCard, Loader2, ShieldCheck } from "lucide-react";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { companyService } from "@/services/company-service";
import { subscriptionService, type BetaReadiness, type SubscriptionStatus } from "@/services/subscription-service";

function entitlementLabel(key: string) {
  return key.replaceAll("_", " ").replace(/\b\w/g, (letter) => letter.toUpperCase());
}

export function SubscriptionHealthCard() {
  const [subscription, setSubscription] = useState<SubscriptionStatus | null>(null);
  const [readiness, setReadiness] = useState<BetaReadiness | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let active = true;
    Promise.all([subscriptionService.status(), companyService.getAccess()])
      .then(async ([status, access]) => {
        if (!active) return;
        setSubscription(status);
        if (access.can_manage_members) {
          try { setReadiness(await subscriptionService.betaReadiness()); } catch { /* admin helper is non-blocking */ }
        }
      })
      .finally(() => active && setLoading(false));
    return () => { active = false; };
  }, []);

  if (loading) return <Card><CardContent className="flex min-h-32 items-center justify-center"><Loader2 className="size-5 animate-spin" /></CardContent></Card>;
  if (!subscription) return null;

  const entitlementEntries = Object.entries(subscription.entitlements).filter(([, value]) => value === true || typeof value === "number").slice(0, 7);

  return (
    <Card className="overflow-hidden border-indigo-200/80">
      <div className="h-1.5 bg-gradient-to-r from-indigo-600 via-violet-500 to-sky-500" />
      <CardHeader>
        <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
          <div className="flex items-start gap-3">
            <div className="rounded-xl bg-indigo-500/10 p-2.5 text-indigo-700"><CreditCard className="size-5" /></div>
            <div><CardTitle>Plan & beta health</CardTitle><CardDescription className="mt-1">Application entitlements are separated from the future billing provider, so pricing can evolve without changing finance logic.</CardDescription></div>
          </div>
          <div className="rounded-xl border bg-background px-4 py-2 text-right">
            <p className="text-xs uppercase tracking-[.14em] text-muted-foreground">Current plan</p>
            <p className="mt-1 font-semibold capitalize">{subscription.plan} · {subscription.status.replaceAll("_", " ")}</p>
            {subscription.days_remaining != null ? <p className="mt-1 flex items-center justify-end gap-1 text-xs text-muted-foreground"><Clock3 className="size-3" />{subscription.days_remaining} trial day(s) remaining</p> : null}
          </div>
        </div>
      </CardHeader>
      <CardContent className="space-y-5">
        <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-4">
          {entitlementEntries.map(([key, value]) => <div key={key} className="rounded-xl border bg-muted/20 p-3"><p className="text-[11px] text-muted-foreground">{entitlementLabel(key)}</p><p className="mt-1 font-semibold">{value === true ? "Included" : value === -1 ? "Unlimited" : Number(value).toLocaleString()}</p></div>)}
        </div>
        {readiness ? <div className="rounded-2xl border bg-muted/20 p-4">
          <div className="flex items-center justify-between gap-3"><div><p className="text-sm font-semibold">Customer workspace readiness</p><p className="text-xs text-muted-foreground">A practical pre-demo / paid-beta check for this workspace.</p></div><div className={`rounded-full px-3 py-1 text-sm font-semibold ${readiness.status === "ready" ? "bg-emerald-100 text-emerald-800" : readiness.status === "blocked" ? "bg-red-100 text-red-800" : "bg-amber-100 text-amber-800"}`}>{readiness.score}%</div></div>
          <div className="mt-4 grid gap-2 md:grid-cols-2">{readiness.checks.map((check) => <div key={check.key} className="flex gap-2 rounded-xl bg-background p-3">{check.status === "ready" ? <CheckCircle2 className="mt-0.5 size-4 shrink-0 text-emerald-600" /> : check.status === "blocked" ? <CircleAlert className="mt-0.5 size-4 shrink-0 text-red-600" /> : <ShieldCheck className="mt-0.5 size-4 shrink-0 text-amber-600" />}<div><p className="text-sm font-medium">{check.label}</p><p className="mt-0.5 text-xs leading-5 text-muted-foreground">{check.detail}</p></div></div>)}</div>
        </div> : null}
      </CardContent>
    </Card>
  );
}
