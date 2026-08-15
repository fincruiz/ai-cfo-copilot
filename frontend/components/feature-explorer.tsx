"use client";

import Link from "next/link";
import { ArrowRight, Search, Sparkles, WandSparkles, X } from "lucide-react";
import { useMemo, useRef, useState } from "react";
import { usageService } from "@/services/usage-service";

export interface Capability {
  group: string;
  label: string;
  description: string;
  href: string;
  keywords?: string;
}

export function FeatureExplorer({ capabilities, onClose }: { capabilities: Capability[]; onClose: () => void }) {
  const [query, setQuery] = useState("");
  const searchTracked = useRef(false);
  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return capabilities;
    return capabilities.filter((item) => `${item.label} ${item.description} ${item.group} ${item.keywords ?? ""}`.toLowerCase().includes(q));
  }, [capabilities, query]);
  const groups = Array.from(new Set(results.map((item) => item.group)));

  function updateQuery(value: string) {
    setQuery(value);
    if (value.trim() && !searchTracked.current) {
      searchTracked.current = true;
      usageService.track("explore_search_used"); // query text deliberately excluded for privacy
    }
  }

  return (
    <div className="fixed inset-0 z-[120] flex items-start justify-center bg-slate-950/55 p-4 pt-[5vh] backdrop-blur-md" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="flex max-h-[88vh] w-full max-w-5xl animate-modal-in flex-col overflow-hidden rounded-[30px] border bg-background shadow-[0_40px_120px_rgba(15,23,42,.35)]">
        <div className="border-b bg-gradient-to-br from-primary/[.07] via-background to-sky-500/[.05] p-6 sm:p-8">
          <div className="flex items-start gap-4">
            <div className="flex size-12 items-center justify-center rounded-2xl bg-primary text-primary-foreground shadow-lg"><WandSparkles className="size-5" /></div>
            <div className="min-w-0 flex-1">
              <p className="text-xs font-semibold uppercase tracking-[.18em] text-primary">Capability explorer</p>
              <h2 className="mt-1 text-2xl font-semibold sm:text-3xl">Explore everything FinCruiz can do</h2>
              <p className="mt-2 max-w-2xl text-sm leading-6 text-muted-foreground">Search by feature name or by the business outcome you want. FinCruiz keeps advanced tools discoverable without forcing every tool into the everyday navigation.</p>
            </div>
            <button type="button" onClick={onClose} className="flex size-10 items-center justify-center rounded-xl border bg-background/70 hover:bg-muted"><X className="size-4" /></button>
          </div>
          <label className="mt-6 flex items-center gap-3 rounded-2xl border bg-background px-5 py-4 shadow-sm focus-within:ring-2 focus-within:ring-primary/20">
            <Search className="size-5 text-primary" />
            <input autoFocus value={query} onChange={(e) => updateQuery(e.target.value)} className="w-full bg-transparent text-base outline-none" placeholder="Try: improve cash, model hiring, board pack, Xero, profit margin..." />
          </label>
          <div className="mt-4 flex flex-wrap gap-2">
            {["Cash & collections", "Model a decision", "Understand profit", "Board reporting"].map((prompt) => <button key={prompt} type="button" onClick={() => updateQuery(prompt)} className="rounded-full border bg-background/80 px-3 py-1.5 text-xs font-medium text-muted-foreground hover:border-primary/30 hover:text-foreground">{prompt}</button>)}
          </div>
        </div>
        <div className="flex-1 overflow-y-auto p-5 sm:p-7">
          {groups.length ? groups.map((group) => (
            <section key={group} className="mb-8 last:mb-0">
              <div className="mb-3 flex items-center gap-2"><Sparkles className="size-3.5 text-primary"/><p className="text-[11px] font-semibold uppercase tracking-[.18em] text-muted-foreground">{group}</p></div>
              <div className="grid gap-3 md:grid-cols-2">
                {results.filter((item) => item.group === group).map((item) => (
                  <Link key={`${item.group}-${item.label}`} href={item.href} onClick={() => { usageService.track("feature_opened_from_explorer", { feature: item.label, group: item.group }); onClose(); }} className="group rounded-2xl border p-4 transition duration-200 hover:-translate-y-0.5 hover:border-primary/30 hover:bg-primary/[.035] hover:shadow-lg">
                    <div className="flex items-start justify-between gap-4"><div><p className="font-semibold">{item.label}</p><p className="mt-1 text-sm leading-6 text-muted-foreground">{item.description}</p></div><ArrowRight className="mt-1 size-4 shrink-0 text-muted-foreground transition group-hover:translate-x-1 group-hover:text-primary" /></div>
                  </Link>
                ))}
              </div>
            </section>
          )) : <div className="py-16 text-center"><p className="font-semibold">No exact match</p><p className="mt-2 text-sm text-muted-foreground">Try a business outcome such as “cash”, “profit”, “forecast”, “customers” or “board”.</p></div>}
        </div>
      </div>
    </div>
  );
}
