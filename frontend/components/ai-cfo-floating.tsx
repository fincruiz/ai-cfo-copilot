"use client";

import { FormEvent, useEffect, useState } from "react";
import { createPortal } from "react-dom";
import { usePathname, useRouter } from "next/navigation";
import {
  Bot,
  ExternalLink,
  Globe2,
  Loader2,
  MessageCircleQuestion,
  Navigation,
  Send,
  Sparkles,
  X,
} from "lucide-react";

import { analyticsService } from "@/services/analytics-service";
import { usageService } from "@/services/usage-service";
import { InsightChart } from "@/components/insight-chart";
import type { AICFOAnswer } from "@/types/analytics";

interface Message {
  role: "user" | "assistant";
  content: string;
  result?: AICFOAnswer;
}

export function AICFOFloating() {
  const router = useRouter();
  const pathname = usePathname();
  const [open, setOpen] = useState(false);
  const [mounted, setMounted] = useState(false);
  const [question, setQuestion] = useState("");
  const [loading, setLoading] = useState(false);
  const [liveContext, setLiveContext] = useState(true);
  const prompts = [
    "What should management focus on?",
    "Can we afford to hire 3 more people?",
    "Why is profit up but cash down?",
    "Where should I upload my GL?",
    "Check industry & economic risks",
  ];
  const [promptIndex, setPromptIndex] = useState(0);

  useEffect(() => { setMounted(true); }, []);

  useEffect(() => {
    const timer = window.setInterval(() => setPromptIndex((value) => (value + 1) % prompts.length), 2800);
    return () => window.clearInterval(timer);
  }, []);

  const [messages, setMessages] = useState<Message[]>([
    {
      role: "assistant",
      content:
        "Tell me what you are trying to decide in plain English. I can guide you to the right FinCruiz capability, analyse loaded company data, and use tools such as the three-way forecast when a decision needs modelling.",
    },
  ]);

  async function ask(value: string) {
    const cleaned = value.trim();
    if (!cleaned || loading) return;
    setMessages((current) => [...current, { role: "user", content: cleaned }]);
    setQuestion("");
    setLoading(true);
    usageService.track("ai_question_submitted", { source: "floating_assistant" }); // question text deliberately excluded
    try {
      const conversation = messages
        .filter((message) => message.content && !(message.role === "assistant" && !message.result))
        .map((message) => ({ role: message.role, content: message.content }))
        .slice(-8);
      const response = await analyticsService.askAiCfo(cleaned, liveContext, conversation);
      setMessages((current) => [...current, { role: "assistant", content: response.answer, result: response }]);
    } catch {
      setMessages((current) => [
        ...current,
        {
          role: "assistant",
          content:
            "I could not retrieve the finance context. Confirm the backend is running and your workspace is available.",
        },
      ]);
    } finally {
      setLoading(false);
    }
  }

  async function submit(event: FormEvent) {
    event.preventDefault();
    await ask(question);
  }

  useEffect(() => {
    function openFromContext(event: Event) {
      const detail = (event as CustomEvent<{ question?: string }>).detail;
      setOpen(true);
      if (detail?.question) window.setTimeout(() => void ask(detail.question!), 80);
    }
    window.addEventListener("fincruiz:open-ai", openFromContext as EventListener);
    return () => window.removeEventListener("fincruiz:open-ai", openFromContext as EventListener);
  // `ask` intentionally reads current component state; this listener is re-bound when the assistant renders.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [liveContext, loading]);

  if (!mounted) return null;

  return createPortal(
    <>
      {open ? (
        <section className="fixed bottom-24 right-5 z-[9997] flex h-[min(700px,calc(100vh-120px))] w-[min(460px,calc(100vw-32px))] animate-ai-drawer flex-col overflow-hidden rounded-[26px] border border-indigo-400/20 bg-slate-950 text-white shadow-[0_30px_100px_rgba(15,23,42,0.45)]">
          <header className="border-b border-white/10 bg-gradient-to-r from-indigo-600/25 to-sky-500/10 p-5">
            <div className="flex items-center gap-3">
              <div className="flex size-11 items-center justify-center rounded-2xl bg-white/10">
                <Bot className="size-5 text-indigo-200" />
              </div>
              <div className="min-w-0 flex-1">
                <p className="font-bold">AI CFO & Platform Guide</p>
                <p className="truncate text-xs text-slate-300">Finance data + optional live market context</p>
              </div>
              <button type="button" onClick={() => setOpen(false)} className="flex size-9 items-center justify-center rounded-xl bg-white/5 hover:bg-white/10">
                <X className="size-4" />
              </button>
            </div>
            <label className="mt-4 flex cursor-pointer items-center justify-between rounded-xl border border-white/10 bg-white/5 px-3 py-2 text-xs text-slate-300">
              <span className="flex items-center gap-2"><Globe2 className="size-3.5 text-sky-300" /> Use live industry & economic context when relevant</span>
              <input type="checkbox" checked={liveContext} onChange={(event) => setLiveContext(event.target.checked)} />
            </label>
          </header>

          <div className="flex-1 space-y-3 overflow-y-auto p-4">
            {messages.map((message, index) => (
              <div key={`${message.role}-${index}`} className={message.role === "user" ? "ml-auto max-w-[88%]" : "max-w-[94%]"}>
                <div className={[
                  "rounded-2xl px-4 py-3 text-sm leading-6 whitespace-pre-wrap",
                  message.role === "user" ? "bg-indigo-600 text-white" : "bg-white/8 text-slate-100",
                ].join(" ")}>
                  {message.content}
                </div>
                {message.role === "assistant" && message.result?.visualization ? <InsightChart visualization={message.result.visualization} /> : null}

                {message.result?.action ? (
                  <button
                    type="button"
                    onClick={() => { usageService.track("ai_recommended_tool_opened", { source: "floating_assistant", feature: message.result!.action!.label }); setOpen(false); router.push(message.result!.action!.route); }}
                    className="mt-2 inline-flex items-center gap-2 rounded-xl bg-indigo-500/20 px-3 py-2 text-xs font-semibold text-indigo-100 hover:bg-indigo-500/30"
                  >
                    <Navigation className="size-3.5" /> {message.result.action.label}
                  </button>
                ) : null}

                {message.result?.external_context_used ? (
                  <div className="mt-2 flex items-center gap-2 text-[11px] text-sky-300"><Globe2 className="size-3" /> Live external context used</div>
                ) : null}

                {message.result?.sources?.length ? (
                  <div className="mt-2 space-y-1 rounded-xl border border-white/10 bg-white/5 p-3">
                    <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-400">External sources</p>
                    {message.result.sources.slice(0, 4).map((source) => (
                      <a key={source.url} href={source.url} target="_blank" rel="noreferrer" className="flex items-start gap-2 text-xs text-sky-300 hover:text-sky-200">
                        <ExternalLink className="mt-0.5 size-3 shrink-0" /> <span className="line-clamp-2">{source.title}</span>
                      </a>
                    ))}
                  </div>
                ) : null}

                {message.result?.suggested_questions?.length ? (
                  <div className="mt-2 flex flex-wrap gap-1.5">
                    {message.result.suggested_questions.slice(0, 3).map((suggestion) => (
                      <button key={suggestion} type="button" onClick={() => void ask(suggestion)} className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-[11px] text-slate-300 hover:bg-white/10">
                        {suggestion}
                      </button>
                    ))}
                  </div>
                ) : null}
              </div>
            ))}

            {loading ? (
              <div className="flex items-center gap-2 rounded-2xl bg-white/8 px-4 py-3 text-sm text-slate-300">
                <Loader2 className="size-4 animate-spin" /> Reviewing company and relevant market context...
              </div>
            ) : null}
          </div>

          <div className="border-t border-white/10 p-4">
            <div className="mb-3 flex gap-2 overflow-x-auto pb-1">
              {["Where do I upload my GL?", "What should management focus on?", "What economic risks matter to us?"].map((item) => (
                <button key={item} type="button" onClick={() => void ask(item)} className="whitespace-nowrap rounded-full border border-white/10 px-3 py-1.5 text-[11px] text-slate-300 hover:bg-white/10">{item}</button>
              ))}
            </div>
            <form onSubmit={submit}>
              <div className="flex gap-2 rounded-2xl bg-white/8 p-2">
                <input
                  value={question}
                  onChange={(event) => setQuestion(event.target.value)}
                  placeholder="Ask about your business or how to use FinCruiz…"
                  className="min-w-0 flex-1 bg-transparent px-3 text-sm text-white outline-none placeholder:text-slate-400"
                />
                <button type="submit" disabled={loading || !question.trim()} className="flex size-11 items-center justify-center rounded-xl bg-indigo-500 text-white transition hover:bg-indigo-400 disabled:opacity-40">
                  <Send className="size-4" />
                </button>
              </div>
            </form>
          </div>
        </section>
      ) : null}

      {pathname !== "/dashboard" ? <div className="fixed bottom-5 right-5 z-[9997] flex items-center gap-3 sm:bottom-7 sm:right-7">
        {!open ? (
          <div className="hidden animate-soft-bob rounded-2xl border bg-background/95 px-4 py-3 text-sm font-semibold shadow-xl backdrop-blur sm:block">
            <span className="inline-block min-w-52 animate-prompt-swap">{prompts[promptIndex]}</span>
          </div>
        ) : null}
        <button type="button" onClick={() => setOpen((current) => !current)} className="group relative flex size-16 animate-ai-pulse items-center justify-center rounded-[22px] bg-gradient-to-br from-indigo-600 to-sky-500 text-white shadow-[0_18px_50px_rgba(79,70,229,0.45)] transition duration-300 hover:-translate-y-1 hover:scale-105" title="Open AI CFO Assistant">
          {open ? <X className="size-6" /> : <MessageCircleQuestion className="size-7" />}
          {!open ? (
            <>
              <span className="absolute -right-1 -top-1 flex size-5 items-center justify-center rounded-full bg-emerald-500 ring-4 ring-background"><Sparkles className="size-3" /></span>
              <span className="absolute inset-0 -z-10 animate-ping rounded-[22px] bg-indigo-500/25" />
            </>
          ) : null}
        </button>
      </div> : null}
    </>,
    document.body,
  );
}
