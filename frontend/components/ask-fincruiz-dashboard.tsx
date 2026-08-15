"use client";

import { FormEvent, useState } from "react";
import { useRouter } from "next/navigation";
import {
  ArrowRight,
  Bot,
  Globe2,
  Loader2,
  Navigation,
  Send,
  Sparkles,
} from "lucide-react";

import { analyticsService } from "@/services/analytics-service";
import { usageService } from "@/services/usage-service";
import type { AICFOAnswer } from "@/types/analytics";

const prompts = [
  "What should management focus on today?",
  "Why is profit changing?",
  "Can we afford to hire 3 people?",
  "What is putting pressure on cash?",
];

export function AskFinCruizDashboard() {
  const router = useRouter();
  const [question, setQuestion] = useState("");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<AICFOAnswer | null>(null);
  const [liveContext, setLiveContext] = useState(true);

  async function ask(value: string) {
    const cleaned = value.trim();
    if (!cleaned || loading) return;
    setQuestion(cleaned);
    setLoading(true);
    setResult(null);
    usageService.track("ai_question_submitted", { source: "dashboard_ask_fincruiz" });
    try {
      const response = await analyticsService.askAiCfo(cleaned, liveContext);
      setResult(response);
    } catch {
      setResult({
        answer: "I could not retrieve the company context just now. Check that the workspace data and backend are available, then try again.",
      } as AICFOAnswer);
    } finally {
      setLoading(false);
    }
  }

  async function submit(event: FormEvent) {
    event.preventDefault();
    await ask(question);
  }

  return (
    <section className="overflow-hidden rounded-[28px] border border-indigo-200/70 bg-gradient-to-br from-indigo-50 via-background to-sky-50 shadow-sm dark:border-indigo-500/20 dark:from-indigo-950/30 dark:via-background dark:to-sky-950/20">
      <div className="grid gap-0 lg:grid-cols-[.8fr_1.2fr]">
        <div className="border-b border-indigo-100/80 p-6 dark:border-white/10 lg:border-b-0 lg:border-r">
          <div className="flex items-center gap-3">
            <div className="flex size-12 items-center justify-center rounded-2xl bg-gradient-to-br from-indigo-600 to-sky-500 text-white shadow-lg">
              <Bot className="size-5" />
            </div>
            <div>
              <p className="text-xs font-semibold uppercase tracking-[.18em] text-indigo-700 dark:text-indigo-300">Ask FinCruiz</p>
              <h2 className="text-xl font-semibold">Ask about the business, not the software</h2>
            </div>
          </div>
          <p className="mt-4 text-sm leading-6 text-muted-foreground">
            Ask in plain English. FinCruiz can explain company performance, guide you to the right tool and route decision questions into models such as the Three-Way Forecast.
          </p>
          <div className="mt-5 flex flex-wrap gap-2">
            {prompts.map((prompt) => (
              <button
                key={prompt}
                type="button"
                onClick={() => void ask(prompt)}
                className="rounded-full border border-indigo-200 bg-white/80 px-3 py-1.5 text-xs font-medium text-slate-700 transition hover:-translate-y-0.5 hover:border-indigo-400 hover:shadow-sm dark:border-white/10 dark:bg-white/5 dark:text-slate-200"
              >
                {prompt}
              </button>
            ))}
          </div>
        </div>

        <div className="p-6">
          <form onSubmit={submit}>
            <label className="block text-sm font-semibold">What do you want to understand or decide?</label>
            <div className="mt-3 flex gap-2 rounded-2xl border bg-background/90 p-2 shadow-sm focus-within:ring-2 focus-within:ring-indigo-500/20">
              <input
                value={question}
                onChange={(event) => setQuestion(event.target.value)}
                placeholder="e.g. Can we afford to hire 3 people without putting cash under pressure?"
                className="min-w-0 flex-1 bg-transparent px-3 py-2 text-sm outline-none placeholder:text-muted-foreground"
              />
              <button
                type="submit"
                disabled={loading || !question.trim()}
                className="flex size-11 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-600 to-sky-500 text-white shadow-sm transition hover:scale-[1.03] disabled:opacity-40"
                aria-label="Ask FinCruiz"
              >
                {loading ? <Loader2 className="size-4 animate-spin" /> : <Send className="size-4" />}
              </button>
            </div>
          </form>

          <label className="mt-3 flex cursor-pointer items-center gap-2 text-xs text-muted-foreground">
            <input type="checkbox" checked={liveContext} onChange={(event) => setLiveContext(event.target.checked)} />
            <Globe2 className="size-3.5" /> Use live industry and economic context when relevant
          </label>

          <div className="mt-5 min-h-28 rounded-2xl border border-dashed bg-background/60 p-4">
            {loading ? (
              <div className="flex items-center gap-2 text-sm text-muted-foreground"><Loader2 className="size-4 animate-spin" /> Reviewing company evidence and choosing the right analysis...</div>
            ) : result ? (
              <div className="animate-step-in">
                <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-[.14em] text-indigo-700 dark:text-indigo-300"><Sparkles className="size-3.5" /> FinCruiz response</div>
                <p className="mt-3 whitespace-pre-wrap text-sm leading-7">{result.answer}</p>
                {result.action ? (
                  <button
                    type="button"
                    onClick={() => {
                      usageService.track("ai_recommended_tool_opened", { source: "dashboard_ask_fincruiz", feature: result.action!.label });
                      router.push(result.action!.route);
                    }}
                    className="mt-4 inline-flex items-center gap-2 rounded-xl bg-indigo-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-indigo-500"
                  >
                    <Navigation className="size-4" /> {result.action.label} <ArrowRight className="size-4" />
                  </button>
                ) : null}
              </div>
            ) : (
              <p className="text-sm leading-6 text-muted-foreground">Your answer will appear here. FinCruiz does not send the text of your question to product-usage analytics.</p>
            )}
          </div>
        </div>
      </div>
    </section>
  );
}
