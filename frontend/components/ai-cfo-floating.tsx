"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import { usePathname, useRouter } from "next/navigation";
import { Bot, ExternalLink, Loader2, Navigation, Send, Sparkles, X } from "lucide-react";

import { InsightChart } from "@/components/insight-chart";
import { analyticsService } from "@/services/analytics-service";
import { usageService } from "@/services/usage-service";
import type { AICFOAnswer } from "@/types/analytics";

type Message = { role: "user" | "assistant"; content: string; result?: AICFOAnswer };

function pageContext(pathname: string) {
  if (pathname === "/dashboard") return { label: "management dashboard", prompt: "What should management focus on?" };
  if (pathname.includes("reports")) return { label: "financial reports", prompt: "Explain the most important movement in these reports." };
  if (pathname.includes("working-capital")) return { label: "working capital", prompt: "What needs collections or payment attention?" };
  if (pathname.includes("forecast") || pathname.includes("planning")) return { label: "planning and forecast", prompt: "Which assumption deserves the most attention?" };
  if (pathname.includes("branches")) return { label: "branch performance", prompt: "Which branch needs management attention?" };
  if (pathname.includes("analytics") || pathname.includes("bi")) return { label: "performance trends", prompt: "What is the most important trend here?" };
  return { label: "this page", prompt: "What should I pay attention to here?" };
}

export function AICFOFloating() {
  const router = useRouter();
  const pathname = usePathname();
  const context = useMemo(() => pageContext(pathname), [pathname]);
  const [mounted, setMounted] = useState(false);
  const [open, setOpen] = useState(false);
  const [question, setQuestion] = useState("");
  const [loading, setLoading] = useState(false);
  const [messages, setMessages] = useState<Message[]>([
    { role: "assistant", content: "Ask a business question. I’ll use the active company data first and show the evidence you can drill into." },
  ]);

  useEffect(() => setMounted(true), []);

  async function ask(value: string) {
    const cleaned = value.trim();
    if (!cleaned || loading) return;
    setOpen(true);
    setMessages((current) => [...current, { role: "user", content: cleaned }]);
    setQuestion("");
    setLoading(true);
    usageService.track("ai_question_submitted", { source: "floating_assistant", area: pathname.split("/")[2] || "dashboard" });
    try {
      const conversation = messages
        .filter((message) => message.content && !(message.role === "assistant" && !message.result))
        .map((message) => ({ role: message.role, content: message.content }))
        .slice(-8);
      const response = await analyticsService.askAiCfo(cleaned, true, conversation);
      setMessages((current) => [...current, { role: "assistant", content: response.answer, result: response }]);
    } catch {
      setMessages((current) => [...current, { role: "assistant", content: "I could not retrieve the finance context. Confirm the workspace and backend are available, then try again." }]);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    const openFromContext = (event: Event) => {
      const detail = (event as CustomEvent<{ question?: string }>).detail;
      setOpen(true);
      if (detail?.question) window.setTimeout(() => void ask(detail.question!), 50);
    };
    window.addEventListener("fincruiz:open-ai", openFromContext as EventListener);
    return () => window.removeEventListener("fincruiz:open-ai", openFromContext as EventListener);
    // The event handler intentionally uses current conversation state.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loading, messages, pathname]);

  async function submit(event: FormEvent) {
    event.preventDefault();
    await ask(question);
  }

  if (!mounted) return null;

  return createPortal(
    pathname !== "/dashboard" ? <>
      {open ? (
        <section className="fixed bottom-20 right-4 z-[9997] flex h-[min(620px,calc(100vh-100px))] w-[min(420px,calc(100vw-24px))] flex-col overflow-hidden rounded-[24px] border bg-background shadow-[0_28px_90px_rgba(15,23,42,.28)] sm:right-6">
          <header className="flex items-center gap-3 border-b px-4 py-3.5">
            <span className="flex size-9 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-600 to-sky-500 text-white"><Bot className="size-4" /></span>
            <div className="min-w-0 flex-1">
              <p className="font-semibold">Ask FinCruiz</p>
              <p className="truncate text-xs text-muted-foreground">Context: {context.label}</p>
            </div>
            <button type="button" onClick={() => setOpen(false)} className="flex size-9 items-center justify-center rounded-lg hover:bg-muted" aria-label="Close Ask FinCruiz"><X className="size-4" /></button>
          </header>

          <div className="flex-1 space-y-3 overflow-y-auto p-4">
            {messages.map((message, index) => (
              <div key={`${message.role}-${index}`} className={message.role === "user" ? "ml-auto max-w-[88%]" : "max-w-[95%]"}>
                <div className={message.role === "user" ? "rounded-2xl bg-primary px-4 py-3 text-sm leading-6 text-primary-foreground" : "rounded-2xl bg-muted/60 px-4 py-3 text-sm leading-6"}>
                  <div className="whitespace-pre-wrap">{message.content}</div>
                </div>

                {message.result?.visualization ? <InsightChart visualization={message.result.visualization} /> : null}

                {message.result?.evidence?.length ? (
                  <div className="mt-2 rounded-2xl border bg-background p-3">
                    <p className="text-[10px] font-bold uppercase tracking-[.15em] text-muted-foreground">Evidence</p>
                    <div className="mt-2 space-y-1.5">
                      {message.result.evidence.slice(0, 6).map((item, evidenceIndex) => {
                        const content = (
                          <>
                            <span className="min-w-0 flex-1"><span className="block truncate text-xs font-medium">{item.label}</span><span className="block truncate text-[10px] text-muted-foreground">{item.source}{item.period ? ` · ${item.period}` : ""}</span></span>
                            <span className="text-xs font-semibold tabular-nums">{item.value}</span>
                          </>
                        );
                        return item.route ? (
                          <button key={`${item.label}-${evidenceIndex}`} type="button" onClick={() => { setOpen(false); router.push(item.route!); }} className="flex w-full items-center gap-3 rounded-xl px-2.5 py-2 text-left hover:bg-muted/60">{content}<Navigation className="size-3.5 shrink-0 text-muted-foreground" /></button>
                        ) : (
                          <div key={`${item.label}-${evidenceIndex}`} className="flex items-center gap-3 rounded-xl px-2.5 py-2">{content}</div>
                        );
                      })}
                    </div>
                  </div>
                ) : null}

                {message.result?.action ? (
                  <button type="button" onClick={() => { setOpen(false); router.push(message.result!.action!.route); }} className="mt-2 inline-flex items-center gap-2 rounded-xl border px-3 py-2 text-xs font-semibold hover:bg-muted">
                    <Navigation className="size-3.5" />{message.result.action.label}
                  </button>
                ) : null}

                {message.result?.sources?.length ? (
                  <details className="mt-2 rounded-xl border px-3 py-2 text-xs">
                    <summary className="cursor-pointer font-medium">External context sources</summary>
                    <div className="mt-2 space-y-2">
                      {message.result.sources.slice(0, 4).map((source) => <a key={source.url} href={source.url} target="_blank" rel="noreferrer" className="flex items-start gap-2 text-sky-700 hover:underline dark:text-sky-300"><ExternalLink className="mt-0.5 size-3 shrink-0" />{source.title}</a>)}
                    </div>
                  </details>
                ) : null}
              </div>
            ))}
            {loading ? <div className="flex items-center gap-2 rounded-2xl bg-muted/60 px-4 py-3 text-sm text-muted-foreground"><Loader2 className="size-4 animate-spin" />Reviewing company evidence…</div> : null}
          </div>

          <div className="border-t p-3">
            <button type="button" onClick={() => void ask(context.prompt)} disabled={loading} className="mb-2 inline-flex max-w-full items-center gap-1.5 rounded-full bg-muted px-3 py-1.5 text-[11px] font-medium text-muted-foreground hover:text-foreground"><Sparkles className="size-3" /><span className="truncate">{context.prompt}</span></button>
            <form onSubmit={submit} className="flex items-center gap-2 rounded-2xl border bg-background p-1.5">
              <input value={question} onChange={(event) => setQuestion(event.target.value)} placeholder="Ask a business question…" className="min-w-0 flex-1 bg-transparent px-3 text-sm outline-none" />
              <button type="submit" disabled={loading || !question.trim()} className="flex size-10 items-center justify-center rounded-xl bg-primary text-primary-foreground disabled:opacity-35"><Send className="size-4" /></button>
            </form>
          </div>
        </section>
      ) : null}

      <button type="button" onClick={() => setOpen((value) => !value)} className="fixed bottom-5 right-5 z-[9998] inline-flex h-11 items-center gap-2 rounded-full border bg-background/95 px-4 text-sm font-semibold shadow-lg backdrop-blur hover:-translate-y-0.5 sm:right-6" aria-label="Ask FinCruiz">
        <span className="flex size-7 items-center justify-center rounded-full bg-gradient-to-br from-indigo-600 to-sky-500 text-white"><Sparkles className="size-3.5" /></span>
        <span className="hidden sm:inline">Ask FinCruiz</span>
      </button>
    </> : null,
    document.body,
  );
}
