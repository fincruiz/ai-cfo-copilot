"use client";

import { useMemo } from "react";
import { usePathname } from "next/navigation";
import { MessageSquareText, Sparkles } from "lucide-react";
import { usageService } from "@/services/usage-service";

const PAGE_CONTEXT = [
  { match: (p: string) => p.includes("working-capital"), label: "working capital", prompt: "What needs collections or payment attention?" },
  { match: (p: string) => p.includes("forecast") || p.includes("planning"), label: "this plan", prompt: "Which assumption deserves the most attention?" },
  { match: (p: string) => p.includes("analytics") || p.includes("bi"), label: "these trends", prompt: "What is the most important trend here?" },
  { match: (p: string) => p.includes("reports") || p.includes("kpis"), label: "these financials", prompt: "Explain the most important movement here." },
  { match: (p: string) => p.includes("branches"), label: "branch performance", prompt: "Which branch needs management attention?" },
  { match: () => true, label: "this page", prompt: "What should I pay attention to here?" },
];

export function ContextualAIBar() {
  const pathname = usePathname();
  const context = useMemo(() => PAGE_CONTEXT.find((item) => item.match(pathname))!, [pathname]);
  if (pathname === "/dashboard") return null;

  function launch(question?: string) {
    usageService.track("contextual_ai_opened", { area: pathname.split("/")[2] || "dashboard" });
    window.dispatchEvent(new CustomEvent("fincruiz:open-ai", { detail: question ? { question } : {} }));
  }

  return (
    <div className="mt-8 flex flex-wrap items-center justify-between gap-3 rounded-2xl border border-dashed bg-muted/20 px-4 py-3">
      <div className="flex min-w-0 items-center gap-3">
        <span className="flex size-9 shrink-0 items-center justify-center rounded-xl bg-indigo-600/10 text-indigo-600 dark:text-indigo-300"><MessageSquareText className="size-4" /></span>
        <div className="min-w-0"><p className="text-sm font-semibold">Need context on {context.label}?</p><p className="truncate text-xs text-muted-foreground">Ask FinCruiz without leaving the page.</p></div>
      </div>
      <div className="flex gap-2">
        <button type="button" onClick={() => launch(context.prompt)} className="hidden rounded-xl border bg-background px-3 py-2 text-xs font-medium hover:bg-muted sm:inline-flex"><Sparkles className="mr-1.5 size-3.5" />Suggested question</button>
        <button type="button" aria-label="Ask about this page" onClick={() => launch()} className="rounded-xl bg-primary px-3 py-2 text-xs font-semibold text-primary-foreground">Ask FinCruiz</button>
      </div>
    </div>
  );
}