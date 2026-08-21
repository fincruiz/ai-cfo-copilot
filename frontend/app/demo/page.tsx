"use client";

import Link from "next/link";
import { FormEvent, useEffect, useMemo, useRef, useState } from "react";
import {
  ArrowRight,
  BadgeCheck,
  BarChart3,
  BrainCircuit,
  Building2,
  CheckCircle2,
  CircleDollarSign,
  FileBarChart,
  GitBranch,
  Loader2,
  Pause,
  Play,
  PlayCircle,
  Presentation,
  RotateCcw,
  Send,
  ShieldCheck,
  Sparkles,
  TrendingUp,
  Users,
  WalletCards,
  WandSparkles,
} from "lucide-react";

import { demoService, type DemoAnswer } from "@/services/demo-service";
import { marketingService } from "@/services/marketing-service";

type Audience = "owner" | "finance" | "advisor";

type Scene = {
  kicker: string;
  title: string;
  body: string;
  signal: string;
  outcome: string;
  presenter: string;
  proof: string[];
  icon: typeof BarChart3;
};

const scenes: Scene[] = [
  {
    kicker: "1 · Executive pulse",
    title: "Start with what management needs to know.",
    body: "FinCruiz turns prepared finance data into a concise management view: performance, cash, priorities and the issues that deserve attention first.",
    signal: "3 management priorities surfaced",
    outcome: "The owner starts with the answer, not a menu of reports.",
    presenter: "Open here. The point is not another dashboard — it is reducing the time between opening FinCruiz and knowing what deserves attention.",
    proof: ["Revenue ₹24.80M", "Net profit ₹4.12M", "Cash ₹6.21M", "West GP% 36.7%"],
    icon: BarChart3,
  },
  {
    kicker: "2 · Ask the business",
    title: "Ask FinCruiz instead of hunting for the right report.",
    body: "A company-wide copilot can answer management questions, show the evidence used and route the user into the relevant analysis or model.",
    signal: "Evidence shown before action",
    outcome: "Business users do not need to know where every finance function lives.",
    presenter: "Ask a question the prospect actually cares about. Then point to the evidence and confidence, not just the wording of the AI answer.",
    proof: ["AR ageing", "Branch P&L", "12-month trend", "Scenario assumptions"],
    icon: BrainCircuit,
  },
  {
    kicker: "3 · Branch intelligence",
    title: "See which part of the business is changing the group result.",
    body: "Compare Central, North and West in the same management context while preserving the consolidated company view.",
    signal: "West GP% 36.7% vs Central 44.8%",
    outcome: "Management can distinguish a growth problem from a margin-quality problem.",
    presenter: "Use this to show why consolidated numbers can hide the real operating issue. West is not simply smaller — its margin quality is weaker.",
    proof: ["Central GP 44.8%", "North GP 41.6%", "West GP 36.7%", "Group GP 42.4%"],
    icon: GitBranch,
  },
  {
    kicker: "4 · Working capital",
    title: "Connect profit to the cash management actually has.",
    body: "FinCruiz links overdue receivables, debtor days and cash conversion back to the management story instead of treating them as isolated finance KPIs.",
    signal: "₹1.18M overdue AR · 54 debtor days",
    outcome: "A profitable business can see why cash still feels tight.",
    presenter: "This is often the easiest value moment for an owner. Profit is healthy, but receivables are absorbing cash. The system explains the difference in management language.",
    proof: ["Overdue AR ₹1.18M", "28% of AR overdue", "54 debtor days", "Cash ₹6.21M"],
    icon: WalletCards,
  },
  {
    kicker: "5 · Forecast & decide",
    title: "Test the decision before management commits.",
    body: "Hiring, growth and collections assumptions can be pushed into scenario logic so the discussion moves from opinion to cash and profit impact.",
    signal: "3 hires: ₹4.08M closing cash · downside ₹2.92M",
    outcome: "The same question becomes a decision with a visible downside case.",
    presenter: "Show both cases. The base case says the hires are affordable, but slower collections break the management cash buffer. That is the decision insight.",
    proof: ["Base cash ₹4.08M", "Cash buffer ₹3.50M", "Downside cash ₹2.92M", "Scenario profit ₹3.46M"],
    icon: WandSparkles,
  },
  {
    kicker: "6 · Board story",
    title: "Carry the same evidence into the management conversation.",
    body: "FinCruiz can turn governed finance context into concise board and management reporting without rebuilding the story in a separate spreadsheet or slide process.",
    signal: "Performance · risk · action in one narrative",
    outcome: "Reporting closes the loop instead of becoming a separate monthly exercise.",
    presenter: "Finish here. The key message is continuity: the numbers used to understand and model the business are the same numbers used to communicate the decision.",
    proof: ["Performance summary", "Priority risks", "Scenario impact", "Management actions"],
    icon: Presentation,
  },
];

const audienceCopy: Record<Audience, { label: string; headline: string; focus: string; questions: string[]; closeHeadline: string; closeBody: string }> = {
  owner: {
    label: "Owner / CEO",
    headline: "Show the answer first.",
    focus: "Focus on priorities, cash and decisions rather than finance navigation.",
    questions: [
      "What should management focus on?",
      "Why is profit up but cash tight?",
      "Which branch needs attention?",
      "Can we afford to hire 3 people?",
    ],
    closeHeadline: "Bring one real management decision to the next session.",
    closeBody: "We can shape the demo around the priorities, cash pressure, branch performance or decision your management team is working through now.",
  },
  finance: {
    label: "CFO / Finance",
    headline: "Show the evidence and modelling depth.",
    focus: "Focus on deterministic finance context, variance explanation, forecasting and scenario handoffs.",
    questions: [
      "Where are we losing margin?",
      "Build a 12-month forecast.",
      "What happens if revenue grows 10%?",
      "What are the biggest financial risks?",
    ],
    closeHeadline: "Bring your reporting stack and one finance control problem.",
    closeBody: "We can walk through source-to-report traceability, reporting periods, integrations, forecasting and the controls needed before management relies on the numbers.",
  },
  advisor: {
    label: "Accountant / Advisor",
    headline: "Show how compliance data becomes advice.",
    focus: "Focus on repeatable management insight, evidence and the quality of the client conversation.",
    questions: [
      "How can we improve working capital?",
      "Which branch needs attention?",
      "What should the board discuss next?",
      "Where are we losing margin?",
    ],
    closeHeadline: "Bring one client workflow you want to make more advisory.",
    closeBody: "We can show how governed finance data, repeatable evidence and management questions can support a stronger client conversation without fabricating certainty.",
  },
};

const useCases = [
  {
    title: "Profit up, cash still tight",
    question: "Why is profit up but cash tight?",
    icon: WalletCards,
    value: "Connect P&L performance to receivables and cash conversion.",
  },
  {
    title: "One branch is diluting margin",
    question: "Which branch needs attention?",
    icon: GitBranch,
    value: "Move from consolidated performance to the operating driver.",
  },
  {
    title: "Management wants to hire",
    question: "Can we afford to hire 3 people?",
    icon: Users,
    value: "Show the base case, management buffer and downside cash risk.",
  },
  {
    title: "Growth target needs a model",
    question: "What happens if revenue grows 10%?",
    icon: TrendingUp,
    value: "Translate a headline target into forecast profit and cash implications.",
  },
];

function DemoVisualization({ visualization }: { visualization: NonNullable<DemoAnswer["visualization"]> }) {
  const values = visualization.series.flatMap((series) => series.data);
  const max = Math.max(...values, 1);
  const min = Math.min(...values, 0);
  const range = max - min || 1;

  if (visualization.type === "bar") {
    const first = visualization.series[0];
    return (
      <div className="mt-4 rounded-2xl border border-white/10 bg-white/[.04] p-4">
        <p className="text-xs font-black uppercase tracking-[.14em] text-slate-400">{visualization.title}</p>
        <div className="mt-5 flex h-36 items-end gap-3">
          {first.data.map((value, index) => (
            <div key={`${visualization.labels[index]}-${index}`} className="flex flex-1 flex-col items-center justify-end gap-2">
              <span className="text-[10px] font-bold text-indigo-200">{value}</span>
              <span className="w-full rounded-t-lg bg-indigo-400/65" style={{ height: `${Math.max(12, ((value - min) / range) * 105)}px` }} />
              <span className="text-[10px] text-slate-500">{visualization.labels[index]}</span>
            </div>
          ))}
        </div>
      </div>
    );
  }

  const width = 520;
  const height = 150;
  const lineData = visualization.series[0]?.data ?? [];
  const points = lineData.map((value, index, array) => {
    const x = array.length <= 1 ? 0 : (index / (array.length - 1)) * width;
    const y = height - ((value - min) / range) * (height - 16) - 8;
    return `${x},${y}`;
  }).join(" ");

  return (
    <div className="mt-4 rounded-2xl border border-white/10 bg-white/[.04] p-4">
      <p className="text-xs font-black uppercase tracking-[.14em] text-slate-400">{visualization.title}</p>
      <svg viewBox={`0 0 ${width} ${height}`} className="mt-4 h-36 w-full overflow-visible" role="img" aria-label={visualization.title}>
        <polyline fill="none" stroke="currentColor" strokeWidth="5" points={points} className="text-indigo-300" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
      <div className="mt-1 flex justify-between text-[10px] text-slate-500"><span>{visualization.labels[0]}</span><span>{visualization.labels[visualization.labels.length - 1]}</span></div>
    </div>
  );
}

export default function DemoPage() {
  const [scene, setScene] = useState(0);
  const [playing, setPlaying] = useState(false);
  const [presenterMode, setPresenterMode] = useState(false);
  const [audience, setAudience] = useState<Audience>("owner");
  const [question, setQuestion] = useState("Why is profit up but cash tight?");
  const [answer, setAnswer] = useState<DemoAnswer | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const initialAsked = useRef(false);

  const activeScene = scenes[scene];
  const activeAudience = audienceCopy[audience];
  const SceneIcon = activeScene.icon;

  useEffect(() => {
    marketingService.track("demo_viewed");
  }, []);

  useEffect(() => {
    if (!playing) return;
    const id = window.setInterval(() => setScene((current) => (current + 1) % scenes.length), 7200);
    return () => window.clearInterval(id);
  }, [playing]);

  useEffect(() => {
    if (initialAsked.current) return;
    initialAsked.current = true;
    void ask("Why is profit up but cash tight?", "initial");
    // Run once so the demo opens with a complete proof point.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function ask(value: string, source = "typed") {
    const q = value.trim();
    if (!q || loading) return;
    setQuestion(q);
    setLoading(true);
    setError("");
    marketingService.track("demo_question_submitted", { source, audience });
    try {
      setAnswer(await demoService.ask(q));
    } catch (caught: unknown) {
      const message = caught instanceof Error ? caught.message : "The demo AI is temporarily unavailable.";
      setError(message);
    } finally {
      setLoading(false);
    }
  }

  async function submit(event: FormEvent) {
    event.preventDefault();
    await ask(question, "typed");
  }

  function chooseScene(index: number) {
    setScene(index);
    setPlaying(false);
    marketingService.track("demo_guided_scene_clicked", { scene: index + 1 });
  }

  function chooseAudience(next: Audience) {
    setAudience(next);
    marketingService.track("demo_audience_changed", { audience: next });
  }

  function togglePresenter() {
    setPresenterMode((current) => {
      const next = !current;
      marketingService.track("demo_presenter_mode_toggled", { enabled: next });
      return next;
    });
  }

  function showAction(anchor?: string) {
    if (!anchor) return;
    document.getElementById(anchor)?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  const confidenceClass = useMemo(() => {
    if (answer?.confidence === "high") return "bg-emerald-300/10 text-emerald-200";
    if (answer?.confidence === "low") return "bg-rose-300/10 text-rose-200";
    return "bg-amber-300/10 text-amber-200";
  }, [answer?.confidence]);

  return (
    <main className="min-h-screen overflow-hidden bg-[#080d19] text-white">
      <div className="pointer-events-none fixed inset-0 demo-aurora" />

      <header className="sticky top-0 z-40 border-b border-white/10 bg-[#080d19]/90 backdrop-blur-xl">
        <div className="mx-auto flex max-w-7xl items-center justify-between gap-3 px-5 py-4 lg:px-8">
          <Link href="/" className="flex items-center gap-3 text-sm font-semibold text-slate-300 hover:text-white"><span className="flex size-8 items-center justify-center rounded-lg bg-indigo-500 text-white"><BarChart3 className="size-4" /></span><span>FinCruiz</span></Link>
          <div className="flex items-center gap-2">
            <button type="button" onClick={togglePresenter} className={`hidden rounded-xl border px-3 py-2 text-xs font-bold sm:inline-flex ${presenterMode ? "border-indigo-300/30 bg-indigo-300/10 text-indigo-100" : "border-white/10 bg-white/[.04] text-slate-300"}`}>
              <Presentation className="mr-1.5 size-3.5" />Presenter mode {presenterMode ? "on" : "off"}
            </button>
            <span className="hidden rounded-full border border-emerald-300/20 bg-emerald-300/10 px-3 py-1.5 text-xs text-emerald-200 lg:inline-flex"><ShieldCheck className="mr-1.5 size-3.5" />Synthetic data only</span>
            <Link href="/signup" onClick={() => marketingService.track("demo_signup_clicked", { source: "header" })} className="rounded-xl bg-white px-4 py-2.5 text-sm font-black text-slate-950">Use my business data</Link>
          </div>
        </div>
      </header>

      <section className="relative z-10 mx-auto max-w-7xl px-5 pb-24 pt-7 lg:px-8">
        <div className="mx-auto max-w-3xl text-center">
          <div className="inline-flex items-center gap-2 rounded-full border border-indigo-300/15 bg-indigo-300/[.07] px-3.5 py-2 text-xs font-bold text-indigo-100"><PlayCircle className="size-3.5" />Interactive product tour · synthetic data</div>
          <h1 className="mt-5 text-4xl font-black tracking-[-.055em] sm:text-6xl">See FinCruiz through <span className="text-indigo-300">your role.</span></h1>
          <p className="mx-auto mt-4 max-w-2xl text-base leading-7 text-slate-400">Choose how you work, follow a five-minute management story, then ask Nova Retail a real business question. No customer workspace is used.</p>
        </div>

        <div className="mx-auto mt-8 grid max-w-4xl gap-3 sm:grid-cols-3">
          {(Object.keys(audienceCopy) as Audience[]).map((key) => (
            <button key={key} type="button" onClick={() => chooseAudience(key)} className={`rounded-2xl border p-4 text-left transition ${audience === key ? "border-indigo-300/35 bg-indigo-300/10 shadow-[0_14px_40px_rgba(79,70,229,.12)]" : "border-white/10 bg-white/[.025] hover:border-white/20 hover:bg-white/[.045]"}`}>
              <p className="text-xs font-black uppercase tracking-[.13em] text-indigo-200">{audienceCopy[key].label}</p>
              <p className="mt-2 text-sm font-bold">{audienceCopy[key].headline}</p>
            </button>
          ))}
        </div>

        <div className="mt-10 grid gap-8 lg:grid-cols-[.72fr_1.28fr] lg:items-start">
          <div className="lg:sticky lg:top-28">
            <div className="inline-flex items-center gap-2 rounded-full border border-sky-300/15 bg-sky-300/10 px-3 py-1.5 text-xs font-bold text-sky-100"><Sparkles className="size-3.5" />Guided management story · about 5 minutes</div>
            <h2 className="mt-5 text-4xl font-black tracking-[-.05em] sm:text-5xl">Follow the <span className="text-indigo-300">decision loop.</span></h2>
            <p className="mt-5 max-w-xl text-base leading-8 text-slate-300">{activeAudience.focus} Choose a guided story or control each step manually. Presenter mode keeps the commercial talk track available without cluttering the prospect view.</p>

            <div className="mt-7 flex flex-wrap gap-3">
              <button type="button" onClick={() => setPlaying((value) => !value)} className="inline-flex items-center gap-2 rounded-xl bg-indigo-500 px-5 py-3 text-sm font-black hover:bg-indigo-400">{playing ? <Pause className="size-4" /> : <Play className="size-4" />}{playing ? "Pause story" : "Play guided story"}</button>
              <button type="button" onClick={() => { setScene(0); setPlaying(false); }} className="inline-flex items-center gap-2 rounded-xl border border-white/15 bg-white/5 px-5 py-3 text-sm font-semibold"><RotateCcw className="size-4" />Restart</button>
              <button type="button" onClick={togglePresenter} className="inline-flex items-center gap-2 rounded-xl border border-white/15 bg-white/5 px-5 py-3 text-sm font-semibold sm:hidden"><Presentation className="size-4" />Presenter notes</button>
            </div>

            <div className="mt-8 grid gap-2">
              {scenes.map((item, index) => (
                <button key={item.kicker} type="button" onClick={() => chooseScene(index)} className={`flex items-center gap-3 rounded-xl border px-3 py-3 text-left ${index === scene ? "border-indigo-300/30 bg-indigo-300/10" : "border-white/10 bg-white/[.03] hover:bg-white/[.05]"}`}>
                  <span className={`flex size-7 shrink-0 items-center justify-center rounded-lg text-xs font-black ${index === scene ? "bg-indigo-400 text-slate-950" : "bg-white/10 text-slate-400"}`}>{index + 1}</span>
                  <span className="min-w-0"><span className="block truncate text-sm font-bold">{item.title}</span><span className="mt-0.5 block text-[10px] uppercase tracking-wider text-slate-500">{item.kicker}</span></span>
                </button>
              ))}
            </div>
          </div>

          <div>
            <div className="fincruiz-demo-shell overflow-hidden p-1">
              <div className="min-h-[500px] rounded-[24px] bg-slate-950/72 p-6 sm:p-8">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <span className="text-xs font-black uppercase tracking-[.18em] text-indigo-200">{activeScene.kicker}</span>
                  <span className="flex items-center gap-2 text-xs text-slate-400"><span className={`size-2 rounded-full ${playing ? "animate-pulse bg-emerald-400" : "bg-slate-600"}`} />{playing ? "Auto story playing" : "Presenter controlled"}</span>
                </div>

                <div key={scene} className="mt-7 animate-scene-in">
                  <div className="flex size-12 items-center justify-center rounded-2xl bg-indigo-400/10 text-indigo-200"><SceneIcon className="size-6" /></div>
                  <h2 className="mt-5 text-3xl font-black sm:text-4xl">{activeScene.title}</h2>
                  <p className="mt-4 max-w-3xl leading-7 text-slate-300">{activeScene.body}</p>

                  <div className="mt-7 grid gap-3 sm:grid-cols-2">
                    <div className="rounded-2xl border border-emerald-300/10 bg-emerald-300/[.05] p-4">
                      <p className="text-xs font-black uppercase tracking-wider text-emerald-200">Signal to show</p>
                      <p className="mt-2 font-bold">{activeScene.signal}</p>
                    </div>
                    <div className="rounded-2xl border border-violet-300/10 bg-violet-300/[.05] p-4">
                      <p className="text-xs font-black uppercase tracking-wider text-violet-200">Customer outcome</p>
                      <p className="mt-2 text-sm font-bold leading-6">{activeScene.outcome}</p>
                    </div>
                  </div>

                  <div className="mt-5 grid gap-2 sm:grid-cols-2">
                    {activeScene.proof.map((item) => (
                      <div key={item} className="rounded-xl border border-white/10 bg-white/[.04] px-3 py-3 text-sm font-semibold text-slate-200"><BadgeCheck className="mr-2 inline size-3.5 text-indigo-300" />{item}</div>
                    ))}
                  </div>

                  {presenterMode ? (
                    <div className="mt-5 rounded-2xl border border-amber-300/15 bg-amber-300/[.06] p-4">
                      <p className="text-xs font-black uppercase tracking-[.14em] text-amber-200">Presenter talk track</p>
                      <p className="mt-2 text-sm leading-6 text-amber-50/90">{activeScene.presenter}</p>
                    </div>
                  ) : null}
                </div>

                <div className="mt-8 flex items-center justify-between border-t border-white/10 pt-5">
                  <button type="button" disabled={scene === 0} onClick={() => chooseScene(Math.max(0, scene - 1))} className="rounded-xl border border-white/10 px-4 py-2 text-xs font-bold disabled:opacity-30">Previous</button>
                  <span className="text-xs text-slate-500">{scene + 1} / {scenes.length}</span>
                  <button type="button" disabled={scene === scenes.length - 1} onClick={() => chooseScene(Math.min(scenes.length - 1, scene + 1))} className="inline-flex items-center gap-2 rounded-xl bg-white px-4 py-2 text-xs font-black text-slate-950 disabled:opacity-30">Next<ArrowRight className="size-3.5" /></button>
                </div>
              </div>
            </div>
          </div>
        </div>

        <div className="mt-20 text-center">
          <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-300">Use-case shortcuts</p>
          <h2 className="mt-3 text-3xl font-black sm:text-4xl">Pick the problem the prospect already recognises.</h2>
          <p className="mx-auto mt-4 max-w-2xl text-slate-400">Each card asks the demo a real management question and jumps into the evidence-backed answer.</p>
        </div>

        <div className="mt-8 grid gap-4 md:grid-cols-2 xl:grid-cols-4">
          {useCases.map((item) => {
            const Icon = item.icon;
            return (
              <button
                key={item.title}
                type="button"
                onClick={() => {
                  marketingService.track("demo_scenario_clicked", { scenario: item.title, audience });
                  void ask(item.question, "scenario");
                  document.getElementById("ask-demo")?.scrollIntoView({ behavior: "smooth", block: "start" });
                }}
                className="rounded-[26px] border border-white/10 bg-white/[.04] p-5 text-left hover:-translate-y-1 hover:border-indigo-300/25 hover:bg-white/[.06]"
              >
                <Icon className="size-6 text-indigo-300" />
                <h3 className="mt-4 font-black">{item.title}</h3>
                <p className="mt-2 text-sm leading-6 text-slate-400">{item.value}</p>
                <span className="mt-4 inline-flex items-center gap-2 text-xs font-black text-indigo-200">Run scenario<ArrowRight className="size-3.5" /></span>
              </button>
            );
          })}
        </div>

        <div id="ask-demo" className="scroll-mt-28 mt-20 text-center">
          <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-300">Conversational BI</p>
          <h2 className="mt-3 text-3xl font-black sm:text-4xl">Now use the prospect's own question.</h2>
          <p className="mx-auto mt-4 max-w-2xl text-slate-400">The public demo receives only fixed Nova Retail evidence. If the evidence is not in that dataset, the answer should say so rather than inventing it.</p>
        </div>

        <div className="mt-8 grid gap-5 lg:grid-cols-[.68fr_1.32fr]">
          <div className="rounded-[22px] border border-white/10 bg-white/[.05] p-6">
            <div className="flex items-start justify-between gap-3">
              <div><p className="font-black">Nova Retail</p><p className="mt-1 text-xs text-slate-400">Synthetic executive view · 3 branches · 12 months</p></div>
              <span className="rounded-full bg-emerald-300/10 px-2.5 py-1 text-[10px] font-bold text-emerald-200">DEMO</span>
            </div>

            <div className="mt-5 grid grid-cols-2 gap-3">
              {[["Revenue", "₹24.80M"], ["Net profit", "₹4.12M"], ["Cash", "₹6.21M"], ["Gross margin", "42.4%"]].map(([label, value]) => (
                <div key={label} className="rounded-2xl border border-white/10 bg-slate-950/45 p-4"><p className="text-xs text-slate-400">{label}</p><p className="mt-2 text-xl font-black">{value}</p></div>
              ))}
            </div>

            <div className="mt-5 rounded-2xl border border-white/10 bg-slate-950/45 p-4">
              <p className="text-xs font-black uppercase tracking-wider text-slate-400">Best questions for {activeAudience.label}</p>
              <div className="mt-3 flex flex-wrap gap-2">
                {activeAudience.questions.map((item) => (
                  <button key={item} type="button" onClick={() => void ask(item, "suggested")} className="rounded-full border border-white/10 bg-white/[.04] px-3 py-2 text-left text-xs hover:border-indigo-300/40 hover:bg-indigo-300/10">{item}</button>
                ))}
              </div>
            </div>
          </div>

          <div className="rounded-[22px] border border-indigo-300/15 bg-indigo-300/[.06] p-6">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <p className="flex items-center gap-2 text-sm font-black"><BrainCircuit className="size-5 text-indigo-300" />Ask FinCruiz</p>
              <span className="text-xs text-slate-400">Evidence-first public demo</span>
            </div>

            <form onSubmit={submit} className="mt-4 flex gap-2 rounded-2xl border border-white/10 bg-slate-950/55 p-2">
              <input value={question} onChange={(event) => setQuestion(event.target.value)} className="min-w-0 flex-1 bg-transparent px-3 py-2 text-sm outline-none placeholder:text-slate-500" placeholder="e.g. What should management focus on?" />
              <button disabled={loading || !question.trim()} className="flex size-11 items-center justify-center rounded-xl bg-indigo-500 disabled:opacity-40">{loading ? <Loader2 className="size-4 animate-spin" /> : <Send className="size-4" />}</button>
            </form>

            {error ? <p className="mt-3 rounded-xl bg-red-400/10 p-3 text-sm text-red-200">{error}</p> : null}

            {answer ? (
              <div className="mt-5 rounded-2xl bg-slate-950/60 p-5">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <p className="text-xs font-black uppercase tracking-wider text-indigo-300">Management answer</p>
                  <span className={`rounded-full px-2.5 py-1 text-[10px] font-bold uppercase ${confidenceClass}`}>{answer.confidence} confidence</span>
                </div>
                <p className="mt-3 whitespace-pre-wrap text-sm leading-7 text-slate-100">{answer.answer}</p>

                {answer.visualization ? <DemoVisualization visualization={answer.visualization} /> : null}

                {answer.evidence?.length ? (
                  <div className="mt-4 grid gap-2 sm:grid-cols-2">
                    {answer.evidence.map((item, index) => (
                      <div key={`${item.label}-${index}`} className="rounded-xl border border-white/10 bg-white/[.04] p-3"><p className="text-xs text-slate-400">{item.label}</p><p className="mt-1 font-bold">{item.value}</p><p className="mt-1 text-[10px] text-slate-500">{item.source}</p></div>
                    ))}
                  </div>
                ) : (
                  <div className="mt-4 rounded-xl border border-amber-300/10 bg-amber-300/[.05] p-3 text-xs leading-5 text-amber-100">No supporting Nova Retail evidence was available for this question. That refusal is intentional.</div>
                )}

                <p className="mt-4 text-xs leading-5 text-slate-400"><ShieldCheck className="mr-1 inline size-3.5" />{answer.confidence_reason}</p>

                <div className="mt-4 flex flex-wrap gap-2">
                  {answer.action?.demo_anchor ? (
                    <button type="button" onClick={() => showAction(answer.action?.demo_anchor)} className="inline-flex items-center gap-2 rounded-xl bg-indigo-500 px-4 py-2.5 text-xs font-black">{answer.action.label}<ArrowRight className="size-3.5" /></button>
                  ) : null}
                  <Link href="/signup" onClick={() => marketingService.track("demo_signup_clicked", { source: "answer" })} className="inline-flex items-center gap-2 rounded-xl border border-white/10 px-4 py-2.5 text-xs font-bold">Run this on my data<ArrowRight className="size-3.5" /></Link>
                </div>

                {answer.suggested_questions?.length ? (
                  <div className="mt-4 border-t border-white/10 pt-4">
                    <p className="text-[10px] font-black uppercase tracking-[.13em] text-slate-500">Continue the conversation</p>
                    <div className="mt-2 flex flex-wrap gap-2">{answer.suggested_questions.map((item) => <button key={item} type="button" onClick={() => void ask(item, "followup")} className="rounded-full border border-white/10 px-3 py-1.5 text-xs hover:border-indigo-300/40">{item}</button>)}</div>
                  </div>
                ) : null}
              </div>
            ) : loading ? (
              <div className="mt-5 flex items-center gap-2 text-sm text-slate-400"><Loader2 className="size-4 animate-spin" />Reviewing the synthetic evidence…</div>
            ) : null}
          </div>
        </div>

        <div className="mt-20 grid gap-5 lg:grid-cols-3">
          <div id="reporting" className="scroll-mt-28 rounded-[22px] border border-white/10 bg-white/[.04] p-6"><FileBarChart className="size-6 text-indigo-300" /><h3 className="mt-4 text-xl font-black">Management reporting</h3><p className="mt-2 text-sm leading-6 text-slate-400">Prepared P&L, balance-sheet and KPI context becomes the source for management explanation rather than a separate narrative process.</p></div>
          <div id="branches" className="scroll-mt-28 rounded-[22px] border border-white/10 bg-white/[.04] p-6"><GitBranch className="size-6 text-sky-300" /><h3 className="mt-4 text-xl font-black">Branch intelligence</h3><p className="mt-2 text-sm leading-6 text-slate-400">Compare Central, North and West while keeping the group view intact.</p></div>
          <div id="working-capital" className="scroll-mt-28 rounded-[22px] border border-white/10 bg-white/[.04] p-6"><WalletCards className="size-6 text-emerald-300" /><h3 className="mt-4 text-xl font-black">Working capital</h3><p className="mt-2 text-sm leading-6 text-slate-400">Connect AR ageing, debtor days and cash conversion to management action.</p></div>
          <div id="forecasting" className="scroll-mt-28 rounded-[22px] border border-white/10 bg-white/[.04] p-6"><TrendingUp className="size-6 text-violet-300" /><h3 className="mt-4 text-xl font-black">Forecasting</h3><p className="mt-2 text-sm leading-6 text-slate-400">Translate management assumptions into forward-looking profit and cash outcomes.</p></div>
          <div id="decision" className="scroll-mt-28 rounded-[22px] border border-white/10 bg-white/[.04] p-6"><CircleDollarSign className="size-6 text-amber-300" /><h3 className="mt-4 text-xl font-black">Decision simulation</h3><p className="mt-2 text-sm leading-6 text-slate-400">Test hiring, growth, pricing and collection scenarios before committing.</p></div>
          <div id="board" className="scroll-mt-28 rounded-[22px] border border-white/10 bg-white/[.04] p-6"><Building2 className="size-6 text-rose-300" /><h3 className="mt-4 text-xl font-black">Board story</h3><p className="mt-2 text-sm leading-6 text-slate-400">Carry performance, risks, scenarios and actions into a concise management narrative.</p></div>
        </div>

        <div className="mt-16 rounded-[24px] border border-indigo-300/15 bg-gradient-to-br from-indigo-500/15 to-sky-500/10 p-8 text-center sm:p-10">
          <CheckCircle2 className="mx-auto size-8 text-emerald-300" />
          <p className="mt-4 text-xs font-black uppercase tracking-[.18em] text-indigo-200">Next step for {activeAudience.label}</p>
          <h2 className="mx-auto mt-3 max-w-3xl text-3xl font-black">{activeAudience.closeHeadline}</h2>
          <p className="mx-auto mt-3 max-w-2xl text-slate-300">{activeAudience.closeBody}</p>
          <div className="mt-6 flex flex-col justify-center gap-3 sm:flex-row">
            <Link href={`/book-demo?persona=${audience}`} onClick={() => marketingService.track("demo_book_demo_clicked", { persona: audience, source: "final" })} className="inline-flex items-center justify-center gap-2 rounded-xl bg-white px-5 py-3 text-sm font-black text-slate-950">Book this demo<ArrowRight className="size-4" /></Link>
            <Link href="/signup" onClick={() => marketingService.track("demo_signup_clicked", { source: "final" })} className="inline-flex items-center justify-center gap-2 rounded-xl border border-white/15 bg-white/[.04] px-5 py-3 text-sm font-bold">Create workspace<ArrowRight className="size-4" /></Link>
            <Link href="/pricing" onClick={() => marketingService.track("demo_pricing_clicked")} className="inline-flex items-center justify-center gap-2 rounded-xl border border-white/15 bg-white/[.04] px-5 py-3 text-sm font-bold">View pricing<ArrowRight className="size-4" /></Link>
          </div>
        </div>
      </section>
    </main>
  );
}
