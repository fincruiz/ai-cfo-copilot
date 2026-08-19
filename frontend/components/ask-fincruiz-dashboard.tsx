"use client";

import { FormEvent, useEffect, useState } from "react";
import { useRouter } from "next/navigation";
import {
  ArrowRight,
  Bot,
  ChevronDown,
  ChevronUp,
  ExternalLink,
  FileCheck2,
  Gauge,
  Globe2,
  Lightbulb,
  Loader2,
  Navigation,
  Send,
  Sparkles,
  Target,
} from "lucide-react";

import { analyticsService } from "@/services/analytics-service";
import { usageService } from "@/services/usage-service";
import { InsightChart } from "@/components/insight-chart";
import type {
  AICFOAnswer,
  AICFOConversationTurn,
} from "@/types/analytics";

const COLLAPSE_STORAGE_KEY = "fincruiz.ask-fincruiz.collapsed";

const promptGroups = [
  {
    label: "Understand performance",
    items: [
      "Why did profit change this month?",
      "Which branch is hurting margins?",
      "What are my fastest-growing expenses?",
    ],
  },
  {
    label: "Improve cash",
    items: [
      "What is putting pressure on cash?",
      "Which customers are driving working-capital pressure?",
      "What happens if debtor days improve by 10 days?",
    ],
  },
  {
    label: "Plan ahead",
    items: [
      "Forecast the next 12 months.",
      "Can we afford to hire 3 people?",
      "What happens if revenue grows 15%?",
    ],
  },
  {
    label: "Challenge the business",
    items: [
      "What should management focus on today?",
      "What are the three biggest financial risks?",
      "What should the board discuss next?",
    ],
  },
];

export function AskFinCruizDashboard() {
  const router = useRouter();
  const [question, setQuestion] = useState("");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<AICFOAnswer | null>(null);
  const [liveContext, setLiveContext] = useState(true);
  const [conversation, setConversation] = useState<AICFOConversationTurn[]>([]);
  const [collapsed, setCollapsed] = useState(true);

  useEffect(() => {
    try {
      const saved = window.localStorage.getItem(COLLAPSE_STORAGE_KEY);
      if (saved === "false") {
        setCollapsed(false);
      } else if (saved === "true") {
        setCollapsed(true);
      }
    } catch {
      // Storage can be unavailable in privacy-restricted browsers.
      // FinCruiz still remains safely collapsed by default.
    }
  }, []);

  function setCollapsedPreference(next: boolean) {
    setCollapsed(next);
    try {
      window.localStorage.setItem(COLLAPSE_STORAGE_KEY, String(next));
    } catch {
      // The UI state still changes even when browser storage is unavailable.
    }

    usageService.track(
      next ? "ask_fincruiz_collapsed" : "ask_fincruiz_expanded",
      { source: "dashboard_ask_fincruiz" },
    );
  }

  async function ask(value: string) {
    const cleaned = value.trim();
    if (!cleaned || loading) return;

    setQuestion(cleaned);
    setLoading(true);
    setResult(null);

    usageService.track("ai_question_submitted", {
      source: "dashboard_ask_fincruiz",
    });

    try {
      const response = await analyticsService.askAiCfo(
        cleaned,
        liveContext,
        conversation,
      );

      setResult(response);
      setConversation(
        (current): AICFOConversationTurn[] =>
          [
            ...current,
            { role: "user" as const, content: cleaned },
            { role: "assistant" as const, content: response.answer },
          ].slice(-8),
      );
    } catch {
      setResult({
        answer:
          "I could not retrieve the company context just now. Check that the workspace data and backend are available, then try again.",
        mode: "error",
        suggested_questions: [],
        sources: [],
        external_context_used: false,
      });
    } finally {
      setLoading(false);
    }
  }

  async function submit(event: FormEvent) {
    event.preventDefault();
    await ask(question);
  }

  if (collapsed) {
    return (
      <div className="pointer-events-none fixed bottom-5 right-5 z-[80] sm:bottom-6 sm:right-6">
        <button
          type="button"
          onClick={() => setCollapsedPreference(false)}
          className="pointer-events-auto group flex size-14 items-center justify-center rounded-2xl bg-gradient-to-br from-indigo-600 via-blue-600 to-sky-500 text-white shadow-[0_18px_55px_rgba(79,70,229,.35)] ring-1 ring-white/20 transition duration-200 hover:-translate-y-0.5 hover:shadow-[0_22px_65px_rgba(79,70,229,.45)] sm:size-16"
          aria-expanded="false"
          aria-label="Open Ask FinCruiz"
          title="Ask FinCruiz"
        >
          <Bot className="size-6 transition-transform group-hover:scale-105 sm:size-7" />
          <span className="absolute -right-1 -top-1 flex size-5 items-center justify-center rounded-full border-2 border-background bg-emerald-500 text-[10px] font-black text-white">
            ✦
          </span>
        </button>
      </div>
    );
  }

  return (
    <div className="pointer-events-none fixed inset-x-0 bottom-0 z-[80] flex justify-end px-3 pb-3 sm:inset-x-auto sm:bottom-6 sm:right-6 sm:block sm:px-0 sm:pb-0">
      <section className="pointer-events-auto flex max-h-[82vh] w-full flex-col overflow-hidden rounded-t-[28px] border border-indigo-200/80 bg-background shadow-[0_24px_80px_rgba(15,23,42,.28)] dark:border-indigo-500/20 sm:max-h-[76vh] sm:w-[420px] sm:rounded-[28px]">
        <div className="flex shrink-0 items-center justify-between gap-3 border-b bg-gradient-to-r from-indigo-50 via-background to-sky-50 px-4 py-3 dark:from-indigo-950/30 dark:via-background dark:to-sky-950/20">
          <div className="flex min-w-0 items-center gap-3">
            <div className="flex size-9 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-600 to-sky-500 text-white shadow-sm">
              <Bot className="size-4" />
            </div>
            <div className="min-w-0">
              <p className="truncate text-sm font-semibold">Ask FinCruiz</p>
              <p className="truncate text-[11px] text-muted-foreground">
                Evidence-backed management intelligence
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={() => setCollapsedPreference(true)}
            className="inline-flex shrink-0 items-center gap-1.5 rounded-full border bg-background/90 px-3 py-1.5 text-xs font-semibold text-muted-foreground transition hover:text-foreground"
            aria-expanded="true"
            aria-label="Collapse Ask FinCruiz"
          >
            Minimise
            <ChevronDown className="size-3.5" />
          </button>
        </div>

        <div className="min-h-0 flex-1 overflow-y-auto overscroll-contain">
          <div className="p-4">
            <form onSubmit={submit}>
              <label className="text-xs font-semibold uppercase tracking-[.12em] text-muted-foreground">
                Ask a business question
              </label>

              <div className="mt-2 flex gap-2 rounded-2xl border bg-background p-2 shadow-sm focus-within:ring-2 focus-within:ring-indigo-500/20">
                <input
                  value={question}
                  onChange={(event) => setQuestion(event.target.value)}
                  placeholder="e.g. Why is profit up but cash down?"
                  className="min-w-0 flex-1 bg-transparent px-2 py-2 text-sm outline-none placeholder:text-muted-foreground"
                />

                <button
                  type="submit"
                  disabled={loading || !question.trim()}
                  className="flex size-10 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-600 to-sky-500 text-white shadow-sm transition hover:scale-[1.03] disabled:opacity-40"
                  aria-label="Ask FinCruiz"
                >
                  {loading ? (
                    <Loader2 className="size-4 animate-spin" />
                  ) : (
                    <Send className="size-4" />
                  )}
                </button>
              </div>
            </form>

            <label className="mt-3 flex cursor-pointer items-center gap-2 text-[11px] text-muted-foreground">
              <input
                type="checkbox"
                checked={liveContext}
                onChange={(event) => setLiveContext(event.target.checked)}
              />
              <Globe2 className="size-3.5" />
              Use live external context when relevant
            </label>

            {!result && !loading ? (
              <div className="mt-4">
                <p className="text-[10px] font-bold uppercase tracking-[.14em] text-muted-foreground">
                  Try asking
                </p>

                <div className="mt-2 space-y-3">
                  {promptGroups.map((group) => (
                    <div key={group.label}>
                      <p className="mb-1.5 text-[10px] font-semibold text-indigo-700 dark:text-indigo-300">
                        {group.label}
                      </p>
                      <div className="flex flex-wrap gap-1.5">
                        {group.items.map((prompt) => (
                          <button
                            key={prompt}
                            type="button"
                            onClick={() => void ask(prompt)}
                            className="rounded-full border bg-background px-2.5 py-1.5 text-[11px] leading-4 transition hover:border-indigo-400 hover:bg-indigo-50 dark:hover:bg-indigo-950/20"
                          >
                            {prompt}
                          </button>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ) : null}

            <div className="mt-4 min-h-24 rounded-2xl border border-dashed bg-muted/20 p-3">
              {loading ? (
                <div className="flex items-center gap-2 text-sm text-muted-foreground">
                  <Loader2 className="size-4 animate-spin" />
                  Reviewing company evidence...
                </div>
              ) : result ? (
                <div className="animate-step-in">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="flex items-center gap-2 text-[10px] font-semibold uppercase tracking-[.14em] text-indigo-700 dark:text-indigo-300">
                      <Sparkles className="size-3.5" />
                      FinCruiz response
                    </div>

                    <span className="rounded-full border px-2 py-1 text-[9px] font-semibold uppercase tracking-wide text-muted-foreground">
                      {result.mode?.replaceAll("_", " ") || "analysis"}
                    </span>
                  </div>

                  {result.interpreted_question &&
                  result.interpreted_question !== question ? (
                    <p className="mt-2 text-[10px] leading-4 text-muted-foreground">
                      Interpreted as:{" "}
                      <span className="font-medium text-foreground">
                        {result.interpreted_question}
                      </span>
                    </p>
                  ) : null}

                  <div className="mt-3 rounded-2xl border border-indigo-200/70 bg-gradient-to-br from-white to-indigo-50/60 p-3 dark:border-indigo-500/20 dark:from-white/[.04] dark:to-indigo-950/20">
                    <p className="flex items-center gap-2 text-[10px] font-semibold uppercase tracking-[.14em] text-indigo-700 dark:text-indigo-300">
                      <Lightbulb className="size-3.5" />
                      Management answer
                    </p>

                    <p className="mt-2 whitespace-pre-wrap text-sm font-medium leading-6">
                      {result.answer}
                    </p>
                  </div>

                  {result.visualization ? (
                    <div className="mt-3">
                      <div className="mb-2 flex items-center gap-2 text-[10px] font-semibold text-muted-foreground">
                        <Target className="size-3.5" />
                        Evidence visual
                      </div>
                      <InsightChart visualization={result.visualization} />
                    </div>
                  ) : null}

                  {result.evidence?.length || result.confidence ? (
                    <div className="mt-3 rounded-2xl border bg-background p-3">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <p className="flex items-center gap-2 text-[10px] font-semibold uppercase tracking-[.12em] text-muted-foreground">
                          <FileCheck2 className="size-3.5" />
                          Evidence
                        </p>

                        {result.confidence ? (
                          <span
                            className={`inline-flex items-center gap-1.5 rounded-full px-2 py-1 text-[9px] font-semibold uppercase ${
                              result.confidence === "high"
                                ? "bg-emerald-100 text-emerald-700"
                                : result.confidence === "low"
                                  ? "bg-rose-100 text-rose-700"
                                  : "bg-amber-100 text-amber-700"
                            }`}
                          >
                            <Gauge className="size-3" />
                            {result.confidence}
                          </span>
                        ) : null}
                      </div>

                      {result.evidence?.length ? (
                        <div className="mt-2 grid gap-2">
                          {result.evidence.slice(0, 5).map((item, index) => (
                            <div
                              key={`${item.label}-${index}`}
                              className="rounded-xl border bg-muted/20 px-3 py-2"
                            >
                              <div className="flex items-start justify-between gap-3">
                                <span className="text-[11px] text-muted-foreground">
                                  {item.label}
                                </span>
                                <span className="text-xs font-semibold tabular-nums">
                                  {item.value}
                                </span>
                              </div>
                              <p className="mt-1 text-[9px] text-muted-foreground">
                                {item.source}
                                {item.period ? ` · ${item.period}` : ""}
                              </p>
                            </div>
                          ))}
                        </div>
                      ) : null}

                      {result.confidence_reason ? (
                        <p className="mt-2 text-[10px] leading-4 text-muted-foreground">
                          {result.confidence_reason}
                        </p>
                      ) : null}
                    </div>
                  ) : null}

                  <div className="mt-3 flex flex-wrap gap-2">
                    {result.decision_handoff ? (
                      <button
                        type="button"
                        onClick={() => {
                          usageService.track("ai_decision_handoff_opened", {
                            source: "dashboard_ask_fincruiz",
                            feature: result.decision_handoff!.title,
                          });
                          router.push(result.decision_handoff!.route);
                          setCollapsedPreference(true);
                        }}
                        className="inline-flex items-center gap-2 rounded-xl bg-gradient-to-r from-violet-600 to-indigo-600 px-3 py-2 text-xs font-semibold text-white shadow-sm"
                      >
                        <Sparkles className="size-3.5" />
                        Model this
                        <ArrowRight className="size-3.5" />
                      </button>
                    ) : result.action ? (
                      <button
                        type="button"
                        onClick={() => {
                          usageService.track("ai_recommended_tool_opened", {
                            source: "dashboard_ask_fincruiz",
                            feature: result.action!.label,
                          });
                          router.push(result.action!.route);
                          setCollapsedPreference(true);
                        }}
                        className="inline-flex items-center gap-2 rounded-xl bg-indigo-600 px-3 py-2 text-xs font-semibold text-white shadow-sm"
                      >
                        <Navigation className="size-3.5" />
                        {result.action.label}
                        <ArrowRight className="size-3.5" />
                      </button>
                    ) : null}

                    {result.external_context_used ? (
                      <span className="inline-flex items-center gap-1.5 rounded-xl border bg-background px-2.5 py-2 text-[10px] text-muted-foreground">
                        <Globe2 className="size-3.5" />
                        External context used
                      </span>
                    ) : null}
                  </div>

                  {result.suggested_questions?.length ? (
                    <div className="mt-4 border-t pt-3">
                      <p className="text-[10px] font-semibold uppercase tracking-[.12em] text-muted-foreground">
                        Continue
                      </p>
                      <div className="mt-2 flex flex-wrap gap-1.5">
                        {result.suggested_questions
                          .slice(0, 4)
                          .map((suggestion) => (
                            <button
                              key={suggestion}
                              type="button"
                              onClick={() => void ask(suggestion)}
                              className="rounded-full border bg-background px-2.5 py-1.5 text-[10px] transition hover:border-indigo-400"
                            >
                              {suggestion}
                            </button>
                          ))}
                      </div>
                    </div>
                  ) : null}

                  {result.sources?.length ? (
                    <div className="mt-4 border-t pt-3">
                      <p className="text-[10px] font-semibold uppercase tracking-[.12em] text-muted-foreground">
                        External sources
                      </p>
                      <div className="mt-2 grid gap-2">
                        {result.sources.slice(0, 4).map((source) => (
                          <a
                            key={source.url}
                            href={source.url}
                            target="_blank"
                            rel="noreferrer"
                            className="flex items-start gap-2 rounded-xl border bg-background p-2.5 text-[10px] hover:border-indigo-300"
                          >
                            <ExternalLink className="mt-0.5 size-3.5 shrink-0 text-indigo-500" />
                            <span className="line-clamp-2">{source.title}</span>
                          </a>
                        ))}
                      </div>
                    </div>
                  ) : null}
                </div>
              ) : (
                <p className="text-xs leading-5 text-muted-foreground">
                  Ask about performance, cash, forecasting or management
                  decisions. Suggestions appear above until you start a
                  conversation.
                </p>
              )}
            </div>
          </div>
        </div>
      </section>

      <button
        type="button"
        onClick={() => setCollapsedPreference(true)}
        className="pointer-events-auto absolute -top-16 right-0 hidden size-14 items-center justify-center rounded-2xl bg-gradient-to-br from-indigo-600 via-blue-600 to-sky-500 text-white shadow-[0_18px_55px_rgba(79,70,229,.35)] sm:flex"
        aria-label="Minimise Ask FinCruiz"
        title="Minimise Ask FinCruiz"
      >
        <Bot className="size-6" />
      </button>
    </div>
  );

}