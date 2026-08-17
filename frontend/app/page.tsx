"use client";

import Link from "next/link";
import { useEffect, useMemo, useState } from "react";
import {
  ArrowRight,
  BadgeCheck,
  BarChart3,
  BrainCircuit,
  Building2,
  CheckCircle2,
  ChevronRight,
  CircleDollarSign,
  Database,
  GitBranch,
  LineChart,
  LockKeyhole,
  PlayCircle,
  ShieldCheck,
  Sparkles,
  TrendingUp,
  WandSparkles,
} from "lucide-react";

const proofPoints = [
  "Conversational BI",
  "Evidence-backed AI",
  "Branch intelligence",
  "Target-driven budgeting",
  "Three-way decision modelling",
  "Forecasting & scenarios",
  "Board & management reporting",
  "Regional plans & terminology",
];

const questions = [
  {
    q: "Why is profit up but cash tight?",
    answer:
      "Receivables are the main pressure point. Revenue grew 9.4%, but overdue AR increased faster and absorbed ₹1.18M of cash.",
    evidence: "18,426 GL lines · AR ageing · 12-month trend",
    chart: [64, 68, 66, 73, 78, 81, 79, 88, 94, 98, 104, 112],
  },
  {
    q: "Can we afford to hire 3 people?",
    answer:
      "Yes in the base case. The three-way model keeps closing cash above the ₹3.5M management buffer, but slower collections create a November pressure point.",
    evidence: "P&L + Balance Sheet + Cash Flow · decision scenario",
    chart: [93, 91, 88, 84, 82, 79, 76, 74, 72, 76, 80, 84],
  },
  {
    q: "Which branch needs attention?",
    answer:
      "West branch is growing revenue, but its GP% is 4.2 points below the group average. Freight and discounting explain most of the variance.",
    evidence: "Branch P&L · mapped COA · gross margin bridge",
    chart: [82, 79, 76, 73, 70, 68, 66, 63, 61, 59, 58, 56],
  },
];

export default function HomePage() {
  const [activeQuestion, setActiveQuestion] = useState(0);
  const [activeStep, setActiveStep] = useState(0);

  useEffect(() => {
    const id = window.setInterval(() => {
      setActiveQuestion((v) => (v + 1) % questions.length);
      setActiveStep((v) => (v + 1) % 4);
    }, 4800);
    return () => window.clearInterval(id);
  }, []);

  const current = questions[activeQuestion];
  const maxChart = useMemo(() => Math.max(...current.chart), [current.chart]);

  return (
    <main className="min-h-screen overflow-hidden bg-[#f5f7fb] text-slate-950">
      <div className="pointer-events-none fixed inset-0 landing-aurora" />

      <header className="sticky top-0 z-50 border-b border-white/70 bg-white/78 backdrop-blur-xl">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-5 py-4 lg:px-8">
          <Link href="/" className="flex items-center gap-3">
            <div className="flex size-11 items-center justify-center rounded-2xl bg-slate-950 text-white shadow-lg">
              <BarChart3 className="size-5" />
            </div>
            <div>
              <p className="text-lg font-black tracking-tight">FinCruiz</p>
              <p className="text-xs text-slate-500">Management intelligence</p>
            </div>
          </Link>

          <nav className="hidden items-center gap-1 lg:flex">
            <a href="#how-it-works" className="rounded-xl px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-100">How it works</a>
            <a href="#capabilities" className="rounded-xl px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-100">Capabilities</a>
            <Link href="/pricing" className="rounded-xl px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-100">Pricing</Link>
          </nav>

          <div className="flex items-center gap-2 sm:gap-3">
            <Link href="/demo" className="inline-flex items-center gap-2 rounded-xl border border-indigo-200 bg-indigo-50 px-3 py-2.5 text-sm font-bold text-indigo-700 shadow-sm hover:-translate-y-0.5 hover:shadow-md sm:px-5">
              <PlayCircle className="size-4" />
              <span className="hidden sm:inline">Try interactive demo</span>
              <span className="sm:hidden">Demo</span>
            </Link>
            <Link href="/login" className="rounded-xl px-3 py-2.5 text-sm font-semibold text-slate-700 hover:bg-white sm:px-5">Sign in</Link>
            <Link href="/signup" className="hidden items-center gap-2 rounded-xl bg-slate-950 px-5 py-2.5 text-sm font-bold text-white shadow-lg hover:-translate-y-0.5 hover:bg-slate-800 md:inline-flex">
              Create workspace <ArrowRight className="size-4" />
            </Link>
          </div>
        </div>
      </header>

      <section className="relative z-10 mx-auto grid max-w-7xl items-center gap-12 px-5 pb-20 pt-16 lg:grid-cols-[.9fr_1.1fr] lg:px-8 lg:pb-28 lg:pt-24">
        <div className="animate-rise">
          <div className="inline-flex items-center gap-2 rounded-full border border-indigo-100 bg-white/90 px-4 py-2 text-sm font-bold text-slate-700 shadow-sm">
            <Sparkles className="size-4 text-indigo-600" />
            AI that starts with your evidence, not guesses
          </div>

          <h1 className="mt-7 max-w-4xl text-5xl font-black leading-[1.01] tracking-[-.055em] sm:text-6xl lg:text-[72px]">
            Your business data should
            <span className="block bg-gradient-to-r from-indigo-600 via-violet-600 to-sky-500 bg-clip-text text-transparent">help you decide.</span>
          </h1>

          <p className="mt-7 max-w-2xl text-lg leading-8 text-slate-600">
            FinCruiz connects finance and operational data, explains what changed, shows the evidence, visualises the trend, and models what happens before management makes the next move.
          </p>

          <div className="mt-9 flex flex-col gap-3 sm:flex-row">
            <Link href="/demo" className="group inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-indigo-600 px-7 text-base font-black text-white shadow-[0_18px_45px_rgba(79,70,229,.28)] hover:-translate-y-1 hover:bg-indigo-500">
              <PlayCircle className="size-5" />
              See FinCruiz think
              <ArrowRight className="size-4 transition group-hover:translate-x-1" />
            </Link>
            <Link href="/signup" className="inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-slate-950 px-7 text-base font-bold text-white shadow-xl hover:-translate-y-1 hover:bg-slate-800">
              Use my business data <ArrowRight className="size-4" />
            </Link>
          </div>

          <div className="mt-7 flex flex-wrap gap-x-5 gap-y-2 text-sm text-slate-600">
            {["No data needed for demo", "Evidence before narrative", "Reset or delete your data"].map((item) => (
              <span key={item} className="flex items-center gap-2"><CheckCircle2 className="size-4 text-emerald-600" />{item}</span>
            ))}
          </div>
        </div>

        <div className="relative min-h-[640px] animate-float-in">
          <div className="absolute inset-x-0 top-0 overflow-hidden rounded-[34px] border border-white/90 bg-white/88 p-5 shadow-[0_36px_110px_rgba(30,41,59,.15)] backdrop-blur-xl sm:p-6">
            <div className="flex items-center justify-between gap-4">
              <div>
                <p className="text-xs font-bold uppercase tracking-[.18em] text-slate-400">Ask FinCruiz</p>
                <p className="mt-1 text-xl font-black">{current.q}</p>
              </div>
              <span className="flex shrink-0 items-center gap-2 rounded-full bg-emerald-50 px-3 py-1.5 text-xs font-bold text-emerald-700">
                <span className="size-2 animate-pulse rounded-full bg-emerald-500" /> Evidence ready
              </span>
            </div>

            <div key={activeQuestion} className="mt-5 animate-scene-in rounded-2xl bg-slate-950 p-5 text-white sm:p-6">
              <div className="flex items-center gap-2 text-xs font-bold uppercase tracking-[.16em] text-indigo-300"><BrainCircuit className="size-4" /> Management answer</div>
              <p className="mt-3 text-sm leading-7 text-slate-100 sm:text-base">{current.answer}</p>
              <div className="mt-4 flex items-center gap-2 rounded-xl border border-white/10 bg-white/[.06] px-3 py-2 text-xs text-slate-300">
                <BadgeCheck className="size-4 shrink-0 text-emerald-300" />
                <span><b className="text-white">Based on:</b> {current.evidence}</span>
              </div>
            </div>

            <div className="mt-4 rounded-2xl border bg-slate-50 p-4">
              <div className="flex items-center justify-between"><p className="text-xs font-bold text-slate-500">Visual BI · trend evidence</p><p className="text-xs text-slate-400">12 months</p></div>
              <div className="mt-4 flex h-32 items-end gap-1.5">
                {current.chart.map((value, index) => (
                  <div key={index} className="flex-1 rounded-t-lg bg-gradient-to-t from-indigo-600 to-sky-300 transition-all duration-700" style={{ height: `${Math.max(18, (value / maxChart) * 100)}%` }} />
                ))}
              </div>
            </div>

            <div className="mt-4 grid grid-cols-3 gap-2">
              {questions.map((q, index) => (
                <button key={q.q} onClick={() => setActiveQuestion(index)} className={`rounded-xl border px-3 py-2 text-left text-[11px] font-semibold leading-4 ${index === activeQuestion ? "border-indigo-200 bg-indigo-50 text-indigo-700" : "bg-white text-slate-500"}`}>
                  {q.q}
                </button>
              ))}
            </div>
          </div>

          <div className="absolute -left-3 bottom-2 w-[62%] rounded-[26px] border border-slate-200 bg-slate-950 p-5 text-white shadow-2xl animate-soft-bob">
            <p className="text-xs font-bold uppercase tracking-[.16em] text-violet-300">Decision intelligence</p>
            <p className="mt-2 font-bold">Ask the business question.</p>
            <p className="mt-1 text-xs leading-5 text-slate-400">FinCruiz can route hiring, pricing, expansion or working-capital questions into the appropriate model.</p>
          </div>

          <div className="absolute -right-2 bottom-8 w-[42%] rounded-[24px] border border-indigo-100 bg-white p-4 shadow-xl">
            <div className="flex items-center gap-2 text-xs font-bold text-indigo-700"><ShieldCheck className="size-4" /> Confidence: High</div>
            <p className="mt-2 text-xs leading-5 text-slate-500">Calculated company data and model outputs stay separate from AI interpretation.</p>
          </div>
        </div>
      </section>

      <section className="relative z-10 overflow-hidden border-y border-slate-200/70 bg-slate-950 py-5 text-white">
        <div className="animate-feature-marquee flex min-w-max gap-10 whitespace-nowrap text-sm font-semibold tracking-wide text-slate-200">
          {proofPoints.concat(proofPoints).map((item, index) => <span key={`${item}-${index}`} className="flex items-center gap-2"><Sparkles className="size-3 text-indigo-300" />{item}</span>)}
        </div>
      </section>

      <section id="how-it-works" className="relative z-10 mx-auto max-w-7xl px-5 py-24 lg:px-8">
        <div className="grid gap-10 lg:grid-cols-[.75fr_1.25fr] lg:items-end">
          <div>
            <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-600">See FinCruiz think</p>
            <h2 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">From raw numbers to a management decision.</h2>
          </div>
          <p className="max-w-2xl text-lg leading-8 text-slate-600">The customer does not need to know whether the answer lives in working capital, forecasting, BI or a three-way model. FinCruiz starts with the question and chooses the right evidence path.</p>
        </div>

        <div className="mt-12 grid gap-4 lg:grid-cols-4">
          {[
            [Database, "1. Connect", "ERP, GL, budgets and operational data enter one governed workspace."],
            [LineChart, "2. Understand", "Statements, KPIs, trends and branch performance are calculated deterministically."],
            [BrainCircuit, "3. Explain", "AI translates the evidence into management language and shows where it came from."],
            [TrendingUp, "4. Decide", "Forecast, budget and scenario models show the financial impact before management acts."],
          ].map(([Icon, title, text], index) => {
            const C = Icon as typeof Database;
            return <button key={String(title)} onClick={() => setActiveStep(index)} className={`group rounded-[26px] border p-6 text-left shadow-sm transition duration-300 hover:-translate-y-1 hover:shadow-xl ${activeStep === index ? "border-indigo-200 bg-indigo-50/80" : "bg-white/85"}`}><C className="size-5 text-indigo-600" /><p className="mt-5 text-lg font-black">{String(title)}</p><p className="mt-2 text-sm leading-6 text-slate-600">{String(text)}</p></button>;
          })}
        </div>
      </section>

      <section id="capabilities" className="relative z-10 bg-slate-950 py-24 text-white">
        <div className="mx-auto max-w-7xl px-5 lg:px-8">
          <div className="mx-auto max-w-3xl text-center">
            <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-300">Built for management, not just finance teams</p>
            <h2 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">One platform. Multiple management conversations.</h2>
            <p className="mt-5 text-lg leading-8 text-slate-300">Start high level, then drill down only when the decision requires it.</p>
          </div>

          <div className="mt-12 grid gap-4 md:grid-cols-2 lg:grid-cols-3">
            {[
              [BrainCircuit, "Conversational BI", "Ask in plain English and receive the right chart, explanation and next action."],
              [ShieldCheck, "Evidence-backed AI", "Every important answer can expose period, source, confidence and calculation evidence."],
              [GitBranch, "Branch intelligence", "Compare branches, plan branch targets and understand where performance diverges."],
              [CircleDollarSign, "Target-driven planning", "Set Revenue, GP% and NP targets and let FinCruiz allocate to months and mapped COA."],
              [TrendingUp, "Decision simulator", "Test hiring, pricing, capex, working-capital and revenue assumptions through the three-way model."],
              [Building2, "Management reporting", "Executive dashboards, board packs, forecasts and financial reports stay connected to one data model."],
            ].map(([Icon, title, text]) => {
              const C = Icon as typeof BrainCircuit;
              return <div key={String(title)} className="group rounded-[26px] border border-white/10 bg-white/[.055] p-6 transition hover:-translate-y-1 hover:border-indigo-300/30 hover:bg-white/[.08]"><div className="flex size-11 items-center justify-center rounded-2xl bg-indigo-400/10 text-indigo-200"><C className="size-5" /></div><p className="mt-5 text-xl font-black">{String(title)}</p><p className="mt-2 text-sm leading-7 text-slate-300">{String(text)}</p><Link href="/demo" className="mt-5 inline-flex items-center gap-1 text-sm font-bold text-indigo-300">See in demo <ChevronRight className="size-4" /></Link></div>;
            })}
          </div>
        </div>
      </section>

      <section className="relative z-10 mx-auto max-w-7xl px-5 py-24 lg:px-8">
        <div className="overflow-hidden rounded-[36px] border border-indigo-100 bg-gradient-to-br from-white via-indigo-50/80 to-sky-50 p-7 shadow-[0_30px_100px_rgba(79,70,229,.12)] sm:p-10 lg:p-12">
          <div className="grid gap-10 lg:grid-cols-[1fr_.95fr] lg:items-center">
            <div>
              <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-600">A safer way to use AI in finance</p>
              <h2 className="mt-3 text-4xl font-black tracking-tight">Numbers first. AI second.</h2>
              <p className="mt-5 max-w-xl text-lg leading-8 text-slate-600">FinCruiz keeps financial calculations, scenario outputs and source evidence separate from AI interpretation. That makes answers easier to challenge, trace and trust.</p>
              <div className="mt-7 grid gap-3 sm:grid-cols-2">
                {["Deterministic finance engine", "Evidence & confidence", "Tenant-aware permissions", "Reset & deletion controls"].map(item => <div key={item} className="flex items-center gap-2 rounded-xl bg-white/80 px-3 py-3 text-sm font-bold text-slate-700"><BadgeCheck className="size-4 text-emerald-600" />{item}</div>)}
              </div>
            </div>
            <div className="rounded-[28px] bg-slate-950 p-6 text-white shadow-2xl">
              <p className="text-xs font-black uppercase tracking-[.16em] text-slate-500">Example evidence chain</p>
              <div className="mt-5 space-y-3">
                {["18,426 GL transactions", "AR ageing + branch P&L", "12-month trend", "Three-way scenario output"].map((item, index) => <div key={item} className="flex items-center gap-3 rounded-2xl border border-white/10 bg-white/[.05] p-3"><span className="flex size-7 items-center justify-center rounded-full bg-indigo-400/15 text-xs font-black text-indigo-200">{index + 1}</span><span className="text-sm text-slate-200">{item}</span></div>)}
              </div>
              <div className="mt-4 rounded-2xl border border-emerald-300/15 bg-emerald-300/[.07] p-4 text-sm text-emerald-100"><LockKeyhole className="mb-2 size-4" /><b>Management answer:</b> explain the conclusion, then let the user open the evidence.</div>
            </div>
          </div>
        </div>
      </section>

      <section className="relative z-10 mx-auto max-w-7xl px-5 pb-24 lg:px-8">
        <div className="rounded-[36px] bg-slate-950 p-8 text-white sm:p-10 lg:flex lg:items-center lg:justify-between lg:gap-10">
          <div>
            <p className="text-sm font-black uppercase tracking-[.18em] text-indigo-300">See it before you connect anything</p>
            <h2 className="mt-3 text-3xl font-black sm:text-4xl">Give FinCruiz five minutes.</h2>
            <p className="mt-4 max-w-2xl leading-7 text-slate-300">The interactive demo uses synthetic data and deliberately showcases conversational BI, evidence-backed AI, branch analysis, decision modelling, budgeting and forecasting.</p>
          </div>
          <div className="mt-6 flex shrink-0 flex-col gap-3 sm:flex-row lg:mt-0">
            <Link href="/demo" className="inline-flex min-h-13 items-center justify-center gap-2 rounded-2xl bg-white px-6 font-black text-slate-950">Try interactive demo <ArrowRight className="size-4" /></Link>
            <Link href="/pricing" className="inline-flex min-h-13 items-center justify-center rounded-2xl border border-white/15 px-6 font-bold">View local pricing</Link>
          </div>
        </div>
      </section>

      <footer className="relative z-10 border-t bg-white/70 py-8">
        <div className="mx-auto flex max-w-7xl flex-col gap-4 px-5 text-sm text-slate-500 sm:flex-row sm:items-center sm:justify-between lg:px-8">
          <span>© FinCruiz · Management intelligence</span>
          <div className="flex gap-5"><Link href="/demo">Demo</Link><Link href="/pricing">Pricing</Link><Link href="/login">Sign in</Link></div>
        </div>
      </footer>
    </main>
  );
}
