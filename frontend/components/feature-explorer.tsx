"use client";

import Link from "next/link";
import { ArrowRight, Search, Sparkles, WandSparkles } from "lucide-react";
import { useMemo, useRef, useState } from "react";
import { usageService } from "@/services/usage-service";
import { ViewportModal } from "@/components/ui/viewport-modal";

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
    <ViewportModal
      open
      onClose={onClose}
      title="Search FinCruiz"
      description="Jump to a workflow by feature or business outcome. Advanced tools stay available without crowding the everyday workspace."
      maxWidthClass="max-w-5xl"
    >
      <div className="rounded-2xl border border-border/80 bg-muted/20 p-3">
        <div className="flex items-center gap-3 px-2 pb-3">
          <div className="flex size-10 items-center justify-center rounded-xl bg-primary text-primary-foreground"><WandSparkles className="size-4" /></div>
          <div>
            <p className="fincruiz-eyebrow">Command centre</p>
            <p className="mt-1 text-sm font-medium">Tell FinCruiz what you want to achieve.</p>
          </div>
        </div>
        <label className="flex items-center gap-3 rounded-xl border bg-background px-4 py-3 shadow-sm focus-within:ring-2 focus-within:ring-primary/15">
          <Search className="size-5 text-primary" />
          <input autoFocus value={query} onChange={(e) => updateQuery(e.target.value)} className="w-full bg-transparent text-base outline-none" placeholder="Try: improve cash, model hiring, board pack, Xero, profit margin..." />
        </label>
        <div className="mt-4 flex flex-wrap gap-2 px-1 pb-2">
          {["Cash & collections", "Model a decision", "Understand profit", "Board reporting"].map((prompt) => <button key={prompt} type="button" onClick={() => updateQuery(prompt)} className="rounded-full border bg-background/80 px-3 py-1.5 text-xs font-medium text-muted-foreground hover:border-primary/30 hover:text-foreground">{prompt}</button>)}
        </div>
      </div>
      <div className="mt-6">
        {groups.length ? groups.map((group) => (
          <section key={group} className="mb-8 last:mb-0">
            <div className="mb-3 flex items-center gap-2"><Sparkles className="size-3.5 text-primary"/><p className="text-[11px] font-semibold uppercase tracking-[.18em] text-muted-foreground">{group}</p></div>
            <div className="grid gap-3 md:grid-cols-2">
              {results.filter((item) => item.group === group).map((item) => (
                <Link key={`${item.group}-${item.label}`} href={item.href} onClick={() => { usageService.track("feature_opened_from_explorer", { feature: item.label, group: item.group }); onClose(); }} className="group rounded-xl border border-border/80 bg-card p-4 transition duration-200 hover:border-primary/25 hover:bg-primary/[.025] hover:shadow-md">
                  <div className="flex items-start justify-between gap-4"><div><p className="font-semibold">{item.label}</p><p className="mt-1 text-sm leading-6 text-muted-foreground">{item.description}</p></div><ArrowRight className="mt-1 size-4 shrink-0 text-muted-foreground transition group-hover:translate-x-1 group-hover:text-primary" /></div>
                </Link>
              ))}
            </div>
          </section>
        )) : <div className="py-16 text-center"><p className="font-semibold">No exact match</p><p className="mt-2 text-sm text-muted-foreground">Try a business outcome such as “cash”, “profit”, “forecast”, “customers” or “board”.</p></div>}
      </div>
    </ViewportModal>
  );
}