"use client";

import { FormEvent, useMemo, useState } from "react";
import { usePathname } from "next/navigation";
import { ArrowUp, BrainCircuit, Sparkles } from "lucide-react";
import { readWorkspaceScope } from "@/lib/workspace-scope";
import { usageService } from "@/services/usage-service";

const PAGE_CONTEXT: Array<{ match: (path: string) => boolean; label: string; prompts: string[] }> = [
  { match: (p) => p === "/dashboard", label: "this dashboard", prompts: ["What should management focus on?", "Why is profit moving differently from cash?"] },
  { match: (p) => p.includes("working-capital"), label: "working capital", prompts: ["Which customers need collections attention first?", "What is driving the cash conversion cycle?"] },
  { match: (p) => p.includes("native-planning") || p.includes("planning"), label: "this budget", prompts: ["Which budget assumptions look unrealistic?", "How should I phase this budget using historical seasonality?"] },
  { match: (p) => p.includes("forecast"), label: "this forecast", prompts: ["What could make us miss this forecast?", "What happens to cash if revenue falls 10%?"] },
  { match: (p) => p.includes("decision-simulator"), label: "this decision", prompts: ["What is the biggest risk in this scenario?", "Which assumption has the largest cash impact?"] },
  { match: (p) => p.includes("bi") || p.includes("analytics"), label: "these trends", prompts: ["What is the most important trend here?", "Show me the strongest and weakest movement."] },
  { match: (p) => p.includes("reports") || p.includes("kpis"), label: "these financials", prompts: ["Explain these results in plain English.", "Which number should management investigate first?"] },
  { match: (p) => p.includes("branches"), label: "branch performance", prompts: ["Which branch needs management attention?", "Compare branch profitability and growth."] },
  { match: () => true, label: "this page", prompts: ["What should I pay attention to here?", "What can FinCruiz help me do on this page?"] },
];

export function ContextualAIBar() {
  const pathname = usePathname();
  const [question, setQuestion] = useState("");
  const context = useMemo(() => PAGE_CONTEXT.find((item) => item.match(pathname))!, [pathname]);

  function launch(value: string) {
    const cleaned = value.trim();
    if (!cleaned) return;
    const scope = readWorkspaceScope();
    usageService.track("contextual_ai_opened", { area: pathname.split("/")[2] || "dashboard", scope: scope.mode });
    window.dispatchEvent(new CustomEvent("fincruiz:open-ai", { detail: { question: cleaned } }));
    setQuestion("");
  }

  function submit(event: FormEvent) { event.preventDefault(); launch(question); }

  return (
    <div className="sticky bottom-0 z-20 mt-8 pb-2 pt-3 pointer-events-none">
      <div className="pointer-events-auto mx-auto max-w-5xl rounded-[22px] border border-indigo-200/70 bg-background/95 p-2.5 shadow-[0_18px_55px_rgba(15,23,42,.14)] backdrop-blur-xl dark:border-indigo-500/20">
        <form onSubmit={submit} className="flex items-center gap-2">
          <div className="flex size-10 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-600 to-sky-500 text-white"><BrainCircuit className="size-4"/></div>
          <div className="hidden min-w-0 sm:block"><p className="text-[10px] font-bold uppercase tracking-[.14em] text-indigo-600 dark:text-indigo-300">Ask FinCruiz</p><p className="max-w-36 truncate text-[11px] text-muted-foreground">About {context.label}</p></div>
          <input value={question} onChange={(event) => setQuestion(event.target.value)} placeholder={`Ask FinCruiz about ${context.label}…`} className="min-w-0 flex-1 bg-transparent px-2 text-sm outline-none placeholder:text-muted-foreground"/>
          <button type="submit" disabled={!question.trim()} className="flex size-10 shrink-0 items-center justify-center rounded-xl bg-primary text-primary-foreground transition hover:opacity-90 disabled:opacity-35"><ArrowUp className="size-4"/></button>
        </form>
        <div className="mt-2 hidden gap-1.5 overflow-x-auto px-12 pb-0.5 sm:flex">
          {context.prompts.map((prompt) => <button key={prompt} type="button" onClick={() => launch(prompt)} className="whitespace-nowrap rounded-full bg-muted/70 px-3 py-1 text-[10px] font-medium text-muted-foreground transition hover:bg-indigo-50 hover:text-indigo-700 dark:hover:bg-indigo-950/40 dark:hover:text-indigo-200"><Sparkles className="mr-1 inline size-2.5"/>{prompt}</button>)}
        </div>
      </div>
    </div>
  );
}
