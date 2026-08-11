"use client";

import { FormEvent, useEffect, useState } from "react";
import {
  Bot,
  Loader2,
  MessageCircleQuestion,
  Send,
  Sparkles,
  X,
} from "lucide-react";

import { analyticsService } from "@/services/analytics-service";

interface Message {
  role: "user" | "assistant";
  content: string;
}

export function AICFOFloating() {
  const [open, setOpen] = useState(false);
  const [question, setQuestion] = useState("");
  const [loading, setLoading] = useState(false);
  const prompts = ["Ask about overdue AR", "Review vendor exposure", "Explain revenue movement", "Check cash and balance sheet"];
  const [promptIndex, setPromptIndex] = useState(0);
  useEffect(() => {
    const timer = window.setInterval(() => setPromptIndex((value) => (value + 1) % prompts.length), 2800);
    return () => window.clearInterval(timer);
  }, []);

  const [messages, setMessages] = useState<Message[]>([
    {
      role: "assistant",
      content:
        "Ask about revenue, margins, branch performance, AR collections, AP payments, mapping or forecasting.",
    },
  ]);

  async function submit(event: FormEvent) {
    event.preventDefault();
    const value = question.trim();
    if (!value || loading) return;

    setMessages((current) => [...current, { role: "user", content: value }]);
    setQuestion("");
    setLoading(true);

    try {
      const response = await analyticsService.askAiCfo(value);
      setMessages((current) => [
        ...current,
        { role: "assistant", content: response.answer },
      ]);
    } catch {
      setMessages((current) => [
        ...current,
        {
          role: "assistant",
          content:
            "I could not retrieve the finance context. Confirm the backend is running and the company has uploaded data.",
        },
      ]);
    } finally {
      setLoading(false);
    }
  }

  return (
    <>
      {open ? (
        <section className="fixed bottom-24 right-5 z-[80] flex h-[min(650px,calc(100vh-120px))] w-[min(430px,calc(100vw-32px))] animate-ai-drawer flex-col overflow-hidden rounded-[26px] border border-indigo-400/20 bg-slate-950 text-white shadow-[0_30px_100px_rgba(15,23,42,0.45)]">
          <header className="flex items-center gap-3 border-b border-white/10 bg-gradient-to-r from-indigo-600/25 to-sky-500/10 p-5">
            <div className="flex size-11 items-center justify-center rounded-2xl bg-white/10">
              <Bot className="size-5 text-indigo-200" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="font-bold">AI CFO Assistant</p>
              <p className="truncate text-xs text-slate-300">
                Grounded in your uploaded finance data
              </p>
            </div>
            <button
              type="button"
              onClick={() => setOpen(false)}
              className="flex size-9 items-center justify-center rounded-xl bg-white/5 hover:bg-white/10"
            >
              <X className="size-4" />
            </button>
          </header>

          <div className="flex-1 space-y-3 overflow-y-auto p-4">
            {messages.map((message, index) => (
              <div
                key={`${message.role}-${index}`}
                className={[
                  "max-w-[88%] rounded-2xl px-4 py-3 text-sm leading-6",
                  message.role === "user"
                    ? "ml-auto bg-indigo-600 text-white"
                    : "bg-white/8 text-slate-100",
                ].join(" ")}
              >
                {message.content}
              </div>
            ))}
            {loading ? (
              <div className="flex items-center gap-2 rounded-2xl bg-white/8 px-4 py-3 text-sm text-slate-300">
                <Loader2 className="size-4 animate-spin" />
                Reviewing finance context...
              </div>
            ) : null}
          </div>

          <form onSubmit={submit} className="border-t border-white/10 p-4">
            <div className="flex gap-2 rounded-2xl bg-white/8 p-2">
              <input
                value={question}
                onChange={(event) => setQuestion(event.target.value)}
                placeholder="Why is AR overdue?"
                className="min-w-0 flex-1 bg-transparent px-3 text-sm text-white outline-none placeholder:text-slate-400"
              />
              <button
                type="submit"
                disabled={loading || !question.trim()}
                className="flex size-11 items-center justify-center rounded-xl bg-indigo-500 text-white transition hover:bg-indigo-400 disabled:opacity-40"
              >
                <Send className="size-4" />
              </button>
            </div>
          </form>
        </section>
      ) : null}

      <div className="fixed bottom-6 right-6 z-[79] flex items-center gap-3">
        {!open ? (
          <div className="hidden animate-soft-bob rounded-2xl border bg-background/95 px-4 py-3 text-sm font-semibold shadow-xl backdrop-blur sm:block">
            <span className="inline-block min-w-44 animate-prompt-swap">{prompts[promptIndex]}</span>
          </div>
        ) : null}

        <button
          type="button"
          onClick={() => setOpen((current) => !current)}
          className="group relative flex size-16 animate-ai-pulse items-center justify-center rounded-[22px] bg-gradient-to-br from-indigo-600 to-sky-500 text-white shadow-[0_18px_50px_rgba(79,70,229,0.45)] transition duration-300 hover:-translate-y-1 hover:scale-105"
          title="Open AI CFO Assistant"
        >
          {open ? <X className="size-6" /> : <MessageCircleQuestion className="size-7" />}
          {!open ? (
            <>
              <span className="absolute -right-1 -top-1 flex size-5 items-center justify-center rounded-full bg-emerald-500 ring-4 ring-background">
                <Sparkles className="size-3" />
              </span>
              <span className="absolute inset-0 -z-10 animate-ping rounded-[22px] bg-indigo-500/25" />
            </>
          ) : null}
        </button>
      </div>
    </>
  );
}
