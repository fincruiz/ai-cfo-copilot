"use client";

import Link from "next/link";
import { Search, X, ArrowRight, Sparkles } from "lucide-react";
import { useMemo, useState } from "react";

export interface Capability {
  group: string;
  label: string;
  description: string;
  href: string;
  keywords?: string;
}

export function FeatureExplorer({ capabilities, onClose }: { capabilities: Capability[]; onClose: () => void }) {
  const [query, setQuery] = useState("");
  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return capabilities;
    return capabilities.filter((item) => `${item.label} ${item.description} ${item.group} ${item.keywords ?? ""}`.toLowerCase().includes(q));
  }, [capabilities, query]);
  const groups = Array.from(new Set(results.map((item) => item.group)));

  return (
    <div className="fixed inset-0 z-[120] flex items-start justify-center bg-slate-950/45 p-4 pt-[7vh] backdrop-blur-sm" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="flex max-h-[82vh] w-full max-w-4xl flex-col overflow-hidden rounded-[28px] border bg-background shadow-2xl">
        <div className="border-b p-5 sm:p-6">
          <div className="flex items-center gap-3">
            <div className="flex size-10 items-center justify-center rounded-2xl bg-primary/10"><Sparkles className="size-5 text-primary" /></div>
            <div className="min-w-0 flex-1"><h2 className="text-xl font-semibold">Explore everything FinCruiz can do</h2><p className="text-sm text-muted-foreground">Features stay visible even when the everyday navigation is simplified.</p></div>
            <button type="button" onClick={onClose} className="flex size-9 items-center justify-center rounded-xl hover:bg-muted"><X className="size-4" /></button>
          </div>
          <label className="mt-5 flex items-center gap-3 rounded-2xl border bg-muted/30 px-4 py-3">
            <Search className="size-4 text-muted-foreground" />
            <input autoFocus value={query} onChange={(e) => setQuery(e.target.value)} className="w-full bg-transparent text-sm outline-none" placeholder="Search a feature or type what you want to do, e.g. cash, forecast, board pack..." />
          </label>
        </div>
        <div className="flex-1 overflow-y-auto p-5 sm:p-6">
          {groups.length ? groups.map((group) => (
            <section key={group} className="mb-7 last:mb-0">
              <p className="mb-3 text-[11px] font-semibold uppercase tracking-[.18em] text-muted-foreground">{group}</p>
              <div className="grid gap-3 md:grid-cols-2">
                {results.filter((item) => item.group === group).map((item) => (
                  <Link key={`${item.group}-${item.label}`} href={item.href} onClick={onClose} className="group rounded-2xl border p-4 transition hover:border-primary/30 hover:bg-primary/[.03]">
                    <div className="flex items-start justify-between gap-4"><div><p className="font-semibold">{item.label}</p><p className="mt-1 text-sm leading-6 text-muted-foreground">{item.description}</p></div><ArrowRight className="mt-1 size-4 shrink-0 text-muted-foreground transition group-hover:translate-x-0.5 group-hover:text-primary" /></div>
                  </Link>
                ))}
              </div>
            </section>
          )) : <div className="py-14 text-center text-sm text-muted-foreground">No feature matched that search. Try a business outcome such as “cash”, “profit”, “forecast” or “board”.</div>}
        </div>
      </div>
    </div>
  );
}
