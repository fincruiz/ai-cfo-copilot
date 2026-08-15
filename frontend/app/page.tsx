"use client";

import Link from "next/link";
import { useEffect, useState } from "react";
import {
  ArrowRight, BarChart3, BrainCircuit, CheckCircle2, CircleDollarSign, Database,
  LineChart, LockKeyhole, PlayCircle, ShieldCheck, Sparkles, TrendingUp, WandSparkles,
} from "lucide-react";

const ticker = ["Executive intelligence", "Financial statements", "Working capital", "Three-way forecasting", "Industry benchmarking", "Board reporting", "Xero integration", "AI management briefing"];
const briefing = [
  { tone: "amber", title: "Cash conversion needs attention", body: "Receivables are growing faster than revenue. Five customers explain most of the movement." },
  { tone: "emerald", title: "Revenue momentum remains positive", body: "Revenue is tracking 9.4% above the prior period while operating cost growth remains contained." },
  { tone: "sky", title: "Decision ready", body: "Ask whether you can hire, invest or expand and FinCruiz can route the question into the three-way model." },
];

export default function HomePage() {
  const [active, setActive] = useState(0);
  useEffect(() => { const id = window.setInterval(() => setActive((v) => (v + 1) % briefing.length), 3600); return () => window.clearInterval(id); }, []);

  return (
    <main className="min-h-screen overflow-hidden bg-[#f6f8fc] text-slate-950">
      <div className="pointer-events-none fixed inset-0 landing-aurora" />
      <header className="sticky top-0 z-50 border-b border-white/60 bg-white/75 backdrop-blur-xl">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-5 py-4 lg:px-8">
          <Link href="/" className="flex items-center gap-3">
            <div className="flex size-11 items-center justify-center rounded-2xl bg-slate-950 text-white shadow-lg"><BarChart3 className="size-5" /></div>
            <div><p className="text-lg font-bold tracking-tight">FinCruiz</p><p className="text-xs text-slate-500">Organizational intelligence</p></div>
          </Link>
          <div className="flex items-center gap-2 sm:gap-3">
            <Link href="/demo" className="inline-flex items-center gap-2 rounded-xl border border-indigo-200 bg-indigo-50 px-3 py-2.5 text-sm font-bold text-indigo-700 shadow-sm hover:-translate-y-0.5 hover:shadow-md sm:px-5"><PlayCircle className="size-4"/><span className="hidden sm:inline">Try interactive demo</span><span className="sm:hidden">Demo</span></Link>
            <Link href="/login" className="rounded-xl px-3 py-2.5 text-sm font-semibold text-slate-700 hover:bg-white sm:px-5">Sign in</Link>
            <Link href="/signup" className="hidden items-center gap-2 rounded-xl bg-slate-950 px-5 py-2.5 text-sm font-bold text-white shadow-lg hover:-translate-y-0.5 hover:bg-slate-800 md:inline-flex">Create workspace<ArrowRight className="size-4"/></Link>
          </div>
        </div>
      </header>

      <section className="relative z-10 mx-auto grid max-w-7xl items-center gap-12 px-5 pb-20 pt-16 lg:grid-cols-[.92fr_1.08fr] lg:px-8 lg:pb-28 lg:pt-24">
        <div className="animate-rise">
          <div className="inline-flex items-center gap-2 rounded-full border border-indigo-100 bg-white/85 px-4 py-2 text-sm font-semibold text-slate-700 shadow-sm"><Sparkles className="size-4 text-indigo-600"/>Your business already has the answers. FinCruiz connects them.</div>
          <h1 className="mt-7 max-w-3xl text-5xl font-black leading-[1.01] tracking-[-.055em] sm:text-6xl lg:text-7xl">Know what is happening. <span className="block bg-gradient-to-r from-indigo-600 via-violet-600 to-sky-500 bg-clip-text text-transparent">Know what to do next.</span></h1>
          <p className="mt-7 max-w-2xl text-lg leading-8 text-slate-600">FinCruiz turns finance and connected business data into a management briefing: what changed, why it matters, what needs attention, and which decision model to use next.</p>
          <div className="mt-9 flex flex-col gap-3 sm:flex-row">
            <Link href="/demo" className="group inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-indigo-600 px-7 text-base font-bold text-white shadow-[0_18px_45px_rgba(79,70,229,.28)] hover:-translate-y-1 hover:bg-indigo-500"><PlayCircle className="size-5"/>Experience the demo<ArrowRight className="size-4 transition group-hover:translate-x-1"/></Link>
            <Link href="/signup" className="inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-slate-950 px-7 text-base font-bold text-white shadow-xl hover:-translate-y-1 hover:bg-slate-800">Use my own business data<ArrowRight className="size-4"/></Link>
          </div>
          <div className="mt-7 flex flex-wrap gap-x-5 gap-y-2 text-sm text-slate-600">{["No data needed for demo", "Evidence before AI narrative", "Reset or delete your data"].map((item) => <span key={item} className="flex items-center gap-2"><CheckCircle2 className="size-4 text-emerald-600"/>{item}</span>)}</div>
        </div>

        <div className="relative min-h-[610px] animate-float-in">
          <div className="absolute inset-x-0 top-4 mx-auto w-[94%] overflow-hidden rounded-[32px] border border-white/80 bg-white/86 p-5 shadow-[0_35px_100px_rgba(31,41,55,.14)] backdrop-blur-xl sm:p-6">
            <div className="flex items-center justify-between"><div><p className="text-xs font-semibold uppercase tracking-[.18em] text-slate-400">Monday executive briefing</p><p className="mt-1 text-xl font-bold">Here is what management should know</p></div><span className="flex items-center gap-2 rounded-full bg-emerald-50 px-3 py-1 text-xs font-bold text-emerald-700"><span className="size-2 animate-pulse rounded-full bg-emerald-500"/>Live</span></div>
            <div className="mt-5 grid grid-cols-3 gap-3">{[["Revenue","₹24.80M","+9.4%"],["Net profit","₹4.21M","+4.1%"],["Cash","₹6.32M","Watch AR"]].map(([label,value,note]) => <div key={label} className="rounded-2xl border bg-white p-4"><p className="text-xs text-slate-500">{label}</p><p className="mt-2 text-lg font-bold sm:text-xl">{value}</p><p className="mt-1 text-xs text-indigo-600">{note}</p></div>)}</div>
            <div className="mt-4 h-44 rounded-2xl bg-slate-950 p-5"><div className="flex h-full items-end gap-2">{[42,48,51,57,62,68,66,74,81,88,95,104].map((height,index)=><div key={index} className="flex-1 rounded-t-md bg-gradient-to-t from-indigo-600 to-sky-300 transition-all duration-700" style={{height:`${height}px`,animationDelay:`${index*55}ms`}}/>)}</div></div>
          </div>
          <div className="absolute bottom-8 right-0 w-[78%] rounded-[28px] border border-indigo-100 bg-white p-5 shadow-[0_28px_80px_rgba(79,70,229,.18)] sm:p-6">
            <div className="flex items-center gap-3"><div className="flex size-11 items-center justify-center rounded-2xl bg-indigo-50 text-indigo-600"><BrainCircuit className="size-5"/></div><div><p className="font-bold">FinCruiz Intelligence</p><p className="text-xs text-slate-500">Evidence → meaning → action</p></div></div>
            <div key={active} className="mt-4 animate-scene-in"><p className="font-semibold">{briefing[active].title}</p><p className="mt-2 text-sm leading-6 text-slate-600">{briefing[active].body}</p></div>
            <div className="mt-4 flex gap-1.5">{briefing.map((_,i)=><button key={i} onClick={()=>setActive(i)} className={`h-1.5 flex-1 rounded-full ${i===active?"bg-indigo-500":"bg-slate-200"}`} aria-label={`Insight ${i+1}`}/>)}</div>
          </div>
          <div className="absolute left-0 top-[390px] rounded-2xl border bg-slate-950 px-4 py-3 text-white shadow-xl animate-soft-bob"><div className="flex items-center gap-3"><WandSparkles className="size-5 text-sky-300"/><div><p className="text-sm font-semibold">Ask a decision, not a report</p><p className="text-xs text-slate-400">“Can we afford to hire 3 people?”</p></div></div></div>
        </div>
      </section>

      <section className="relative z-10 overflow-hidden border-y border-slate-200/70 bg-slate-950 py-5 text-white"><div className="animate-feature-marquee flex min-w-max gap-10 whitespace-nowrap text-sm font-semibold tracking-wide text-slate-200">{ticker.concat(ticker).map((item,index)=><span key={`${item}-${index}`} className="flex items-center gap-2"><Sparkles className="size-3 text-indigo-300"/>{item}</span>)}</div></section>

      <section className="relative z-10 mx-auto max-w-7xl px-5 py-24 lg:px-8">
        <div className="max-w-3xl"><p className="text-xs font-bold uppercase tracking-[.2em] text-indigo-600">One organizational brain</p><h2 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">From scattered systems to one management conversation.</h2><p className="mt-5 text-lg leading-8 text-slate-600">The customer does not need to know which finance module answers the question. FinCruiz can guide them to the right capability.</p></div>
        <div className="mt-10 grid gap-4 md:grid-cols-4">{[
          [Database,"Connect","ERP, uploads and supporting data enter a governed workspace."],
          [LineChart,"Understand","Financial statements, KPIs, working capital and trends are calculated first."],
          [BrainCircuit,"Explain","FinCruiz translates the evidence into management language."],
          [TrendingUp,"Decide","Questions such as hiring or expansion can flow into scenario and three-way modelling."],
        ].map(([Icon,title,text]) => { const C=Icon as typeof Database; return <div key={String(title)} className="group rounded-[24px] border bg-white/80 p-6 shadow-sm transition duration-300 hover:-translate-y-1 hover:shadow-xl"><C className="size-5 text-indigo-600"/><p className="mt-5 text-lg font-bold">{String(title)}</p><p className="mt-2 text-sm leading-6 text-slate-600">{String(text)}</p></div>; })}</div>
      </section>

      <section className="relative z-10 mx-auto mb-20 max-w-7xl px-5 lg:px-8"><div className="overflow-hidden rounded-[34px] bg-slate-950 p-8 text-white shadow-2xl sm:p-12"><div className="grid gap-8 lg:grid-cols-[1fr_auto] lg:items-center"><div><div className="flex items-center gap-2 text-indigo-300"><PlayCircle className="size-5"/><span className="text-sm font-bold uppercase tracking-[.18em]">No login required</span></div><h2 className="mt-4 text-3xl font-black sm:text-4xl">See FinCruiz think before you trust it with your data.</h2><p className="mt-4 max-w-2xl text-slate-300">The guided demo uses synthetic company data and shows the full path from raw finance information to an executive insight and a modelled decision.</p></div><Link href="/demo" className="inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-white px-7 font-bold text-slate-950 hover:-translate-y-1">Start interactive demo<ArrowRight className="size-4"/></Link></div></div></section>

      <footer className="relative z-10 border-t bg-white/75"><div className="mx-auto flex max-w-7xl flex-col gap-4 px-5 py-8 text-sm text-slate-500 sm:flex-row sm:items-center sm:justify-between lg:px-8"><span>FinCruiz · Financial and organizational intelligence</span><div className="flex gap-4"><Link href="/demo">Demo</Link><Link href="/login">Sign in</Link><Link href="/signup">Create workspace</Link></div></div></footer>
    </main>
  );
}
