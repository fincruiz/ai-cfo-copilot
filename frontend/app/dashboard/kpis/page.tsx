"use client";

import { useEffect, useMemo, useState } from "react";
import { Loader2, RefreshCw } from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { getApiErrorMessage } from "@/lib/api";
import { formatNumber, humanize } from "@/lib/finance-format";
import { financeService } from "@/services/finance-service";
import type { Ratio } from "@/types/finance";

export default function KpisPage() {
  const [ratios, setRatios] = useState<Ratio[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState("");

  async function load() {
    setIsLoading(true);
    setError("");
    try { setRatios(await financeService.getKpis()); }
    catch (loadError) { setError(getApiErrorMessage(loadError)); }
    finally { setIsLoading(false); }
  }

  useEffect(() => { void load(); }, []);

  const groups = useMemo(() => ratios.reduce<Record<string, Ratio[]>>((result, ratio) => {
    (result[ratio.category] ??= []).push(ratio);
    return result;
  }, {}), [ratios]);

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      <div className="flex items-end justify-between gap-4">
        <div><p className="text-sm font-medium text-muted-foreground">Performance</p><h1 className="mt-1 text-3xl font-semibold tracking-tight">Financial KPIs</h1><p className="mt-2 text-muted-foreground">Ratios and indicators calculated from current reports.</p></div>
        <Button onClick={() => void load()} disabled={isLoading}><RefreshCw className="size-4" />Refresh</Button>
      </div>
      {error ? <Alert variant="destructive"><AlertDescription>{error}</AlertDescription></Alert> : null}
      {isLoading ? <div className="flex min-h-64 items-center justify-center gap-2 text-muted-foreground"><Loader2 className="size-5 animate-spin" />Calculating KPIs...</div> : !ratios.length ? <Card><CardContent className="flex min-h-64 items-center justify-center text-muted-foreground">Save account mappings to calculate KPIs.</CardContent></Card> : Object.entries(groups).map(([category, items]) => (
        <section key={category} className="space-y-3">
          <h2 className="text-lg font-semibold">{humanize(category)}</h2>
          <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">
            {items.map((ratio) => <Card key={`${category}-${ratio.name}`}><CardHeader className="pb-3"><div className="flex items-start justify-between gap-3"><div><CardDescription>{ratio.name}</CardDescription><CardTitle className="mt-2 text-3xl">{ratio.value === null || ratio.value === undefined ? "—" : `${formatNumber(ratio.value)}${ratio.unit === "%" ? "%" : ratio.unit ? ` ${ratio.unit}` : ""}`}</CardTitle></div><span className="rounded-full bg-muted px-2 py-1 text-xs font-medium">{ratio.status}</span></div></CardHeader><CardContent><p className="text-sm leading-6 text-muted-foreground">{ratio.interpretation}</p></CardContent></Card>)}
          </div>
        </section>
      ))}
    </div>
  );
}
