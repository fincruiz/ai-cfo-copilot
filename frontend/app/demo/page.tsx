"use client";

import Link from "next/link";
import { useEffect, useMemo, useState } from "react";
import {
  ArrowLeft,
  ArrowRight,
  BadgeCheck,
  BarChart3,
  Bot,
  BrainCircuit,
  Building2,
  CircleDollarSign,
  GitBranch,
  LineChart,
  Pause,
  Play,
  PlayCircle,
  RotateCcw,
  ShieldCheck,
  Sparkles,
  TrendingUp,
  WalletCards,
} from "lucide-react";

type DemoQuestion = {
  question: string;
  answer: string;
  evidence: string;
  confidence: "High" | "Medium";
  action: string;
  chartTitle: string;
  chart: number[];
};

const scenes = [
  {
    kicker: "1 · Connect",
    title: "Bring the business together.",
    body: "Finance, branches, budgets and connected systems feed the same governed intelligence layer.",
    signal: "18,426 GL lines validated",
  },
  {
    kicker: "2 · Understand",
    title: "Turn accounting into a management story.",
    body: "FinCruiz calculates the statements and KPIs first, then highlights what changed and why it matters.",
    signal: "3 management priorities identified",
  },
  {
    kicker: "3 · Ask",
    title: "Use conversational BI instead of hunting for a report.",
    body: "Ask a business question. FinCruiz chooses the evidence, graph and analysis path automatically.",
    signal: "Evidence + chart + confidence",
  },
  {
    kicker: "4 · Decide",
    title: "Model the decision before management acts.",
    body: "Hiring, pricing, capex and working-capital questions can flow into the three-way decision model.",
    signal: "P&L + Balance Sheet + Cash Flow",
  },
  {
    kicker: "5 · Plan",
    title: "Set targets at the level management actually thinks.",
    body: "Enter Revenue, GP% and NP targets. FinCruiz can allocate them to branch, month and mapped COA using historical behaviour.",
    signal: "High-level target → GL-level budget",
  },
];

const questions: DemoQuestion[] = [
  {
    question: "Why is profit up but cash tight?",
    answer: "Working capital is the main reason. Receivables grew faster than revenue, so ₹1.18M of accounting profit has not converted into cash yet.",
    evidence: "P&L · AR ageing · monthly cash trend",
    confidence: "High",
    action: "Open Working Capital",
    chartTitle: "Revenue vs cash conversion",
    chart: [68, 71, 74, 78, 82, 86, 91, 96, 101, 107, 114, 121],
  },
  {
    question: "Which branch needs attention?",
    answer: "West branch revenue is growing, but GP% is 4.2 points below the group average. Freight and discounting explain most of the variance.",
    evidence: "Branch P&L · mapped COA · margin bridge",
    confidence: "High",
    action: "Compare branch drivers",
    chartTitle: "Branch gross margin trend",
    chart: [91, 89, 86, 84, 82, 79, 76, 74, 72, 69, 67, 65],
  },
  {
    question: "Can we afford to hire 3 people?",
    answer: "Yes in the base case. Closing cash stays above the ₹3.5M management buffer, but slower collections would create a pressure point in November.",
    evidence: "Three-way forecast · payroll assumption · cash buffer",
    confidence: "High",
    action: "Model hiring scenario",
    chartTitle: "Base vs hiring cash outlook",
    chart: [100, 98, 96, 92, 88, 84, 80, 77, 75, 78, 82, 86],
  },
  {
    question: "Build next year at ₹30M revenue and 42% GP.",
    answer: "FinCruiz can derive the gross-profit envelope, allowable operating cost and monthly phasing, then allocate the budget to mapped COA using the historical account mix.",
    evidence: "Historical P&L mix · seasonality · mapped COA",
    confidence: "Medium",
    action: "Open Target Budget Builder",
    chartTitle: "Target monthly revenue phasing",
    chart: [73, 75, 77, 80, 84, 88, 92, 96, 100, 105, 111, 118],
  },
];

const branchRows = [
  { name: "Central", revenue: "₹11.8M", gp: "44.8%", trend: "+8.9%", state: "Strong" },
  { name: "North", revenue: "₹7.4M", gp: "41.6%", trend: "+5.2%", state: "Stable" },
  { name: "West", revenue: "₹5.6M", gp: "36.7%", trend: "+12.1%", state: "Watch margin" },
];

export default function DemoPage() {
  const [scene, setScene] = useState(0);
  const [playing, setPlaying] = useState(true);
  const [questionIndex, setQuestionIndex] = useState(0);
  const [scenario, setScenario] = useState<"base" | "hire" | "downside">("base");
  const [budgetMode, setBudgetMode] = useState<"high" | "gl">("high");

  useEffect(() => {
    if (!playing) return;
    const id = window.setInterval(() => setScene((v) => (v + 1) % scenes.length), 5400);
    return () => window.clearInterval(id);
  }, [playing]);

  const currentQuestion = questions[questionIndex];
  const maxChart = useMemo(() => Math.max(...currentQuestion.chart), [currentQuestion.chart]);
  const scenarioData = useMemo(() => {
    if (scenario === "hire") return { revenue: "₹25.10M", profit: "₹3.46M", cash: "₹4.08M", note: "Affordable · buffer remains above target" };
    if (scenario === "downside") return { revenue: "₹22.30M", profit: "₹2.18M", cash: "₹2.92M", note: "Cash buffer breached in November" };
    return { revenue: "₹24.80M", profit: "₹4.12M", cash: "₹6.21M", note: "Current plan" };
  }, [scenario]);

  return (
    <main className="min-h-screen overflow-hidden bg-[#07101f] text-white">
      <div className="fixed inset-0 pointer-events-none demo-aurora" />

      <header className="sticky top-0 z-40 border-b border-white/10 bg-[#07101f]/82 backdrop-blur-xl">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-5 py-4 lg:px-8">
          <Link href="/" className="flex items-center gap-2 text-sm text-slate-300 hover:text-white"><ArrowLeft className="size-4" />FinCruiz</Link>
          <div className="flex items-center gap-2">
            <span className="hidden rounded-full border border-emerald-300/20 bg-emerald-300/10 px-3 py-1.5 text-xs text-emerald-200 sm:inline-flex"><ShieldCheck className="mr-1.5 size-3.5" />Interactive demo · synthetic data</span>
            <Link href="/login" className="hidden rounded-xl border border-white/15 px-4 py-2.5 text-sm font-semibold sm:block">Sign in</Link>
            <Link href="/signup" className="rounded-xl bg-white px-4 py-2.5 text-sm font-black text-slate-950">Use my business data</Link>
          </div>
        </div>
      </header>

      <section className="relative z-10 mx-auto max-w-7xl px-5 pb-20 pt-8 lg:px-8">
        <div className="rounded-2xl border border-indigo-300/15 bg-indigo-300/[.07] px-4 py-3 text-sm text-indigo-100 sm:flex sm:items-center sm:justify-between">
          <span className="flex items-center gap-2"><PlayCircle className="size-4" /><b>Demo Mode.</b>&nbsp; No customer data is used here.</span>
          <span className="mt-2 text-xs text-indigo-200 sm:mt-0">Guided story + interactive management workspace</span>
        </div>

        <div className="mt-10 grid items-center gap-10 lg:grid-cols-[.78fr_1.22fr]">
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-sky-300/15 bg-sky-300/10 px-3 py-1.5 text-xs font-bold text-sky-100"><Sparkles className="size-3.5" />5-minute executive demo</div>
            <h1 className="mt-5 text-4xl font-black tracking-[-.05em] sm:text-6xl">Don't tour software. <span className="bg-gradient-to-r from-sky-300 via-indigo-300 to-violet-300 bg-clip-text text-transparent">Watch it answer the business.</span></h1>
            <p className="mt-5 max-w-xl text-base leading-8 text-slate-300 sm:text-lg">Experience conversational BI, evidence-backed answers, branch intelligence, decision simulation and target-driven planning without signing in.</p>
            <div className="mt-7 flex flex-wrap gap-3">
              <button onClick={() => setPlaying((v) => !v)} className="inline-flex items-center gap-2 rounded-xl bg-indigo-500 px-5 py-3 text-sm font-black hover:bg-indigo-400">{playing ? <Pause className="size-4" /> : <Play className="size-4" />}{playing ? "Pause guided story" : "Continue guided story"}</button>
              <button onClick={() => { setScene(0); setPlaying(true); }} className="inline-flex items-center gap-2 rounded-xl border border-white/15 bg-white/5 px-5 py-3 text-sm font-semibold hover:bg-white/10"><RotateCcw className="size-4" />Restart</button>
            </div>
            <div className="mt-8 grid grid-cols-5 gap-2">{scenes.map((item, index) => <button key={item.kicker} onClick={() => { setScene(index); setPlaying(false); }} className={`h-1.5 rounded-full ${index === scene ? "bg-indigo-300" : "bg-white/10"}`} aria-label={`Show scene ${index + 1}`} />)}</div>
          </div>

          <div className="relative overflow-hidden rounded-[34px] border border-white/10 bg-white/[.04] p-1 shadow-[0_40px_120px_rgba(20,30,80,.5)]">
            <div className="min-h-[510px] rounded-[31px] bg-slate-950/78 p-6 backdrop-blur sm:p-8">
              <div className="flex items-center justify-between"><span className="text-xs font-black uppercase tracking-[.18em] text-indigo-200">{scenes[scene].kicker}</span><span className="flex items-center gap-2 text-xs text-slate-400"><span className="size-2 animate-pulse rounded-full bg-emerald-400" />Guided simulation</span></div>
              <div key={scene} className="mt-8 animate-scene-in">
                <h2 className="text-3xl font-black sm:text-4xl">{scenes[scene].title}</h2>
                <p className="mt-4 max-w-2xl leading-7 text-slate-300">{scenes[scene].body}</p>
                <div className="mt-8 grid gap-4 sm:grid-cols-3">
                  {["Business data", "FinCruiz engine", "Management outcome"].map((label, index) => <div key={label} className={`rounded-2xl border p-4 ${index === Math.min(scene, 2) ? "border-indigo-300/25 bg-indigo-300/[.08]" : "border-white/10 bg-white/[.04]"}`}><p className="text-xs uppercase tracking-wider text-slate-500">{index + 1}</p><p className="mt-2 font-bold">{label}</p></div>)}
                </div>
                <div className="mt-5 flex items-center gap-3 rounded-2xl border border-emerald-300/10 bg-emerald-300/[.05] p-4"><BadgeCheck className="size-5 shrink-0 text-emerald-300" /><div><p className="text-xs uppercase tracking-wider text-emerald-200">Live demo signal</p><p className="mt-1 text-sm font-bold">{scenes[scene].signal}</p></div></div>
              </div>
            </div>
          </div>
        </div>

        <div className="mt-20 text-center"><p className="text-xs font-black uppercase tracking-[.2em] text-indigo-300">Conversational BI</p><h2 className="mt-3 text-3xl font-black sm:text-4xl">Ask the sample business what management actually wants to know.</h2><p className="mx-auto mt-4 max-w-2xl text-slate-400">The answer changes with the question: narrative, chart, evidence, confidence and next action stay connected.</p></div>

        <div className="mt-8 grid gap-5 lg:grid-cols-[.92fr_1.08fr]">
          <div className="rounded-[30px] border border-white/10 bg-white/[.05] p-6">
            <div className="flex items-center justify-between"><div><p className="font-bold">Nova Retail · Executive view</p><p className="mt-1 text-xs text-slate-400">Synthetic multi-branch business</p></div><span className="rounded-full bg-emerald-400/10 px-3 py-1 text-xs text-emerald-200">Evidence ready</span></div>
            <div className="mt-5 grid grid-cols-3 gap-3">{[["Revenue", "₹24.80M"], ["Net profit", "₹4.12M"], ["Cash", "₹6.21M"]].map(([l, v]) => <div key={l} className="rounded-2xl border border-white/10 bg-slate-950/45 p-4"><p className="text-xs text-slate-400">{l}</p><p className="mt-2 text-xl font-black">{v}</p></div>)}</div>
            <div className="mt-5 rounded-2xl border border-white/10 bg-slate-950/45 p-5">
              <div className="flex items-center justify-between"><p className="text-sm font-bold">{currentQuestion.chartTitle}</p><span className="text-xs text-slate-500">12 months</span></div>
              <div key={questionIndex} className="mt-5 flex h-48 items-end gap-2 animate-scene-in">{currentQuestion.chart.map((value, index) => <div key={index} className="flex-1 rounded-t-md bg-gradient-to-t from-indigo-600 to-sky-300" style={{ height: `${Math.max(18, (value / maxChart) * 100)}%` }} />)}</div>
            </div>
          </div>

          <div className="rounded-[30px] border border-indigo-300/15 bg-indigo-300/[.06] p-6">
            <p className="flex items-center gap-2 text-sm font-black"><BrainCircuit className="size-5 text-indigo-300" />Ask FinCruiz</p>
            <div className="mt-4 grid gap-2 sm:grid-cols-2">{questions.map((item, index) => <button key={item.question} onClick={() => setQuestionIndex(index)} className={`rounded-xl border px-3 py-2.5 text-left text-xs font-semibold ${questionIndex === index ? "border-indigo-300/40 bg-indigo-300/15 text-indigo-100" : "border-white/10 text-slate-300 hover:bg-white/[.04]"}`}>{item.question}</button>)}</div>
            <div key={questionIndex} className="mt-5 animate-scene-in rounded-2xl bg-slate-950/60 p-5">
              <div className="flex items-center justify-between gap-3"><p className="text-xs font-black uppercase tracking-wider text-indigo-300">Management answer</p><span className={`rounded-full px-2.5 py-1 text-[10px] font-black ${currentQuestion.confidence === "High" ? "bg-emerald-400/10 text-emerald-200" : "bg-amber-400/10 text-amber-200"}`}>Confidence · {currentQuestion.confidence}</span></div>
              <p className="mt-3 text-sm leading-7 text-slate-100">{currentQuestion.answer}</p>
              <div className="mt-4 rounded-xl border border-white/10 bg-white/[.05] p-3 text-xs text-slate-300"><BadgeCheck className="mr-2 inline size-4 text-emerald-300" /><b>Evidence:</b> {currentQuestion.evidence}</div>
              <div className="mt-3 flex items-center gap-2 rounded-xl bg-violet-400/10 p-3 text-xs font-bold text-violet-200"><TrendingUp className="size-4" />Next action: {currentQuestion.action}</div>
            </div>
          </div>
        </div>

        <div className="mt-16 grid gap-5 lg:grid-cols-2">
          <div className="rounded-[30px] border border-white/10 bg-white/[.04] p-6">
            <div className="flex items-center gap-3"><div className="flex size-11 items-center justify-center rounded-2xl bg-sky-300/10 text-sky-200"><GitBranch className="size-5" /></div><div><p className="text-xs font-black uppercase tracking-[.16em] text-sky-300">Branch intelligence</p><h3 className="mt-1 text-xl font-black">See where group performance is really coming from.</h3></div></div>
            <div className="mt-5 space-y-2">{branchRows.map((row) => <div key={row.name} className="grid grid-cols-[1fr_auto_auto] items-center gap-4 rounded-2xl border border-white/10 bg-slate-950/40 p-4"><div><p className="font-bold">{row.name}</p><p className="text-xs text-slate-500">Revenue {row.revenue} · {row.trend}</p></div><div className="text-right"><p className="text-xs text-slate-500">GP%</p><p className="font-black">{row.gp}</p></div><span className={`rounded-full px-2.5 py-1 text-[10px] font-black ${row.state === "Watch margin" ? "bg-amber-400/10 text-amber-200" : "bg-emerald-400/10 text-emerald-200"}`}>{row.state}</span></div>)}</div>
          </div>

          <div className="rounded-[30px] border border-white/10 bg-white/[.04] p-6">
            <div className="flex items-center justify-between gap-4"><div><p className="text-xs font-black uppercase tracking-[.16em] text-violet-300">Decision simulator</p><h3 className="mt-1 text-xl font-black">Watch one decision move through all three statements.</h3></div><div className="flex gap-1.5">{(["base", "hire", "downside"] as const).map(v => <button key={v} onClick={() => setScenario(v)} className={`rounded-xl px-3 py-2 text-xs font-bold capitalize ${scenario === v ? "bg-white text-slate-950" : "border border-white/10 bg-white/5"}`}>{v}</button>)}</div></div>
            <div key={scenario} className="mt-5 grid animate-scene-in grid-cols-2 gap-3">{[["Revenue", scenarioData.revenue, BarChart3], ["Net profit", scenarioData.profit, CircleDollarSign], ["Closing cash", scenarioData.cash, WalletCards], ["Assessment", scenarioData.note, LineChart]].map(([label, value, Icon]) => { const C = Icon as typeof BarChart3; return <div key={String(label)} className="rounded-2xl border border-white/10 bg-slate-950/40 p-4"><C className="size-4 text-indigo-300" /><p className="mt-3 text-xs text-slate-500">{String(label)}</p><p className="mt-2 text-sm font-black">{String(value)}</p></div>; })}</div>
            <div className="mt-4 rounded-xl border border-violet-300/10 bg-violet-300/[.06] p-3 text-xs text-violet-100">The finance engine calculates the scenario. AI explains the management implication.</div>
          </div>
        </div>

        <div className="mt-16 rounded-[30px] border border-white/10 bg-white/[.04] p-6 sm:p-8">
          <div className="grid gap-8 lg:grid-cols-[.7fr_1.3fr] lg:items-center">
            <div>
              <p className="text-xs font-black uppercase tracking-[.16em] text-emerald-300">Target-driven planning</p>
              <h3 className="mt-2 text-3xl font-black">Management sets the target. FinCruiz builds the detail.</h3>
              <p className="mt-4 text-sm leading-7 text-slate-300">Enter Revenue, GP% and Net Profit at consolidated or branch level. FinCruiz can use historical ratios and seasonality to phase the plan and allocate it to mapped COA.</p>
              <div className="mt-5 flex gap-2"><button onClick={() => setBudgetMode("high")} className={`rounded-xl px-4 py-2 text-xs font-bold ${budgetMode === "high" ? "bg-emerald-300 text-slate-950" : "border border-white/10"}`}>High-level view</button><button onClick={() => setBudgetMode("gl")} className={`rounded-xl px-4 py-2 text-xs font-bold ${budgetMode === "gl" ? "bg-emerald-300 text-slate-950" : "border border-white/10"}`}>GL allocation</button></div>
            </div>
            <div key={budgetMode} className="animate-scene-in rounded-[26px] bg-slate-950/55 p-5">
              {budgetMode === "high" ? <div className="grid gap-3 sm:grid-cols-3">{[["Revenue target", "₹30.00M"], ["Gross margin", "42.00%"], ["Net profit target", "₹3.20M"]].map(([l, v]) => <div key={l} className="rounded-2xl border border-white/10 bg-white/[.04] p-4"><p className="text-xs text-slate-500">{l}</p><p className="mt-2 text-xl font-black">{v}</p></div>)}</div> : <div className="space-y-2">{[["4000 · Product revenue", "₹18.4M", "Historical mix"], ["4010 · Service revenue", "₹11.6M", "Historical mix"], ["5000 · COGS", "₹17.4M", "Derived from 42% GP"], ["6100 · Payroll", "₹4.1M", "Prior-year ratio"], ["6200 · Marketing", "₹1.2M", "Manual override"]].map(([a, v, source]) => <div key={a} className="grid grid-cols-[1fr_auto] gap-4 rounded-xl border border-white/10 bg-white/[.035] p-3"><div><p className="text-sm font-bold">{a}</p><p className="text-[10px] text-slate-500">{source}</p></div><p className="font-black">{v}</p></div>)}</div>}
              <div className="mt-4 flex items-center gap-2 rounded-xl bg-emerald-400/10 p-3 text-xs text-emerald-200"><BadgeCheck className="size-4" />Every generated value keeps its source and can be overridden or restored.</div>
            </div>
          </div>
        </div>

        <div className="mt-16 rounded-[34px] bg-white p-8 text-slate-950 sm:p-10">
          <div className="grid gap-8 lg:grid-cols-[1fr_auto] lg:items-center">
            <div><p className="text-sm font-black uppercase tracking-[.18em] text-indigo-600">Ready when you are</p><h2 className="mt-3 text-3xl font-black">Use the same intelligence layer with your own business.</h2><p className="mt-4 max-w-2xl text-slate-600">Create a workspace, connect or upload data, and let FinCruiz build the management view. You control connections, resets, permissions and deletion.</p></div>
            <div className="flex flex-col gap-2"><Link href="/signup" className="inline-flex min-h-13 items-center justify-center gap-2 rounded-2xl bg-slate-950 px-6 font-black text-white">Create workspace <ArrowRight className="size-4" /></Link><Link href="/pricing" className="text-center text-sm font-bold text-indigo-600">View regional pricing</Link></div>
          </div>
        </div>
      </section>
    </main>
  );
}
