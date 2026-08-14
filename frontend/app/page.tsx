"use client";

import Link from "next/link";
import {
  ArrowRight,
  BarChart3,
  BrainCircuit,
  Building2,
  CheckCircle2,
  FileBarChart2,
  LineChart,
  ShieldCheck,
  Sparkles,
  UploadCloud,
  PlayCircle,
} from "lucide-react";

const features = [
  {
    icon: UploadCloud,
    title: "Upload once. Keep the history.",
    text: "Validate ledgers, preserve versions and build a reusable finance data layer.",
  },
  {
    icon: FileBarChart2,
    title: "Reports that reconcile",
    text: "Generate Trial Balance, P&L, Balance Sheet, KPIs and branch views from one source.",
  },
  {
    icon: LineChart,
    title: "Plan forward",
    text: "Turn monthly actuals into scenarios, forecasts and board-ready insights.",
  },
  {
    icon: BrainCircuit,
    title: "AI CFO intelligence",
    text: "Surface trends, risks, anomalies and management commentary without rebuilding spreadsheets.",
  },
];

export default function HomePage() {
  return (
    <main className="min-h-screen overflow-hidden bg-[#f7f8fb] text-slate-950">
      <div className="absolute inset-x-0 top-0 h-[720px] bg-[radial-gradient(circle_at_20%_20%,rgba(99,102,241,0.16),transparent_34%),radial-gradient(circle_at_82%_18%,rgba(14,165,233,0.16),transparent_30%)]" />

      <header className="relative z-10 mx-auto flex max-w-7xl items-center justify-between px-6 py-6 lg:px-8">
        <Link href="/" className="flex items-center gap-3">
          <div className="flex size-11 items-center justify-center rounded-2xl bg-slate-950 text-white shadow-lg">
            <BarChart3 className="size-5" />
          </div>
          <div>
            <p className="text-lg font-bold tracking-tight">FinCruiz</p>
            <p className="text-xs text-slate-500">Finance intelligence platform</p>
          </div>
        </Link>

        <div className="flex items-center gap-3">
          <Link href="/demo" className="hidden items-center gap-2 rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm transition hover:-translate-y-0.5 hover:shadow md:inline-flex">
            <PlayCircle className="size-4" /> Try demo
          </Link>
          <Link
            href="/login"
            className="hidden rounded-xl px-5 py-3 text-sm font-semibold text-slate-700 transition hover:bg-white hover:shadow-sm sm:block"
          >
            Sign in
          </Link>
          <Link
            href="/signup"
            className="inline-flex items-center gap-2 rounded-xl bg-slate-950 px-6 py-3 text-sm font-semibold text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-slate-800"
          >
            Start free setup
            <ArrowRight className="size-4" />
          </Link>
        </div>
      </header>

      <section className="relative z-10 mx-auto grid max-w-7xl items-center gap-14 px-6 pb-24 pt-16 lg:grid-cols-[0.9fr_1.1fr] lg:px-8 lg:pt-24">
        <div className="animate-rise">
          <div className="mb-6 inline-flex items-center gap-2 rounded-full border border-white/70 bg-white/80 px-4 py-2 text-sm font-medium text-slate-700 shadow-sm backdrop-blur">
            <Sparkles className="size-4" />
            Built for finance teams that have outgrown manual reporting
          </div>

          <h1 className="max-w-3xl text-5xl font-black leading-[1.02] tracking-[-0.045em] sm:text-6xl lg:text-7xl">
            Turn finance data into
            <span className="block bg-gradient-to-r from-indigo-600 via-violet-600 to-sky-500 bg-clip-text text-transparent">
              confident decisions.
            </span>
          </h1>

          <p className="mt-7 max-w-2xl text-lg leading-8 text-slate-600">
            FinCruiz brings reporting, branch consolidation, KPIs, forecasting,
            AI commentary and board reporting into one secure workspace.
          </p>

          <div className="mt-9 flex flex-col gap-3 sm:flex-row">
            <Link
              href="/demo"
              className="inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl border border-slate-200 bg-white px-8 text-base font-bold text-slate-900 shadow-lg transition hover:-translate-y-1 hover:shadow-xl"
            >
              <PlayCircle className="size-5" />
              Explore interactive demo
            </Link>
            <Link
              href="/signup"
              className="inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-slate-950 px-8 text-base font-bold text-white shadow-2xl transition hover:-translate-y-1 hover:bg-slate-800"
            >
              Create your workspace
              <ArrowRight className="size-5" />
            </Link>
            <Link
              href="/login"
              className="inline-flex min-h-14 items-center justify-center rounded-2xl border border-slate-200 bg-white px-8 text-base font-bold text-slate-800 shadow-sm transition hover:-translate-y-1 hover:shadow-lg"
            >
              Sign in to FinCruiz
            </Link>
          </div>

          <div className="mt-9 grid gap-3 text-sm text-slate-600 sm:grid-cols-3">
            {["Email-verified accounts", "Saved finance history", "Branch-ready reporting"].map((item) => (
              <div key={item} className="flex items-center gap-2">
                <CheckCircle2 className="size-4 text-emerald-600" />
                {item}
              </div>
            ))}
          </div>
        </div>

        <div className="relative min-h-[560px] animate-float-in">
          <div className="absolute left-4 top-14 w-[88%] rotate-[-4deg] rounded-[28px] border border-white/80 bg-white/70 p-5 shadow-2xl backdrop-blur-md">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-400">Management dashboard</p>
                <p className="mt-1 text-xl font-bold">Executive overview</p>
              </div>
              <div className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-semibold text-emerald-700">Live</div>
            </div>
            <div className="mt-5 grid grid-cols-3 gap-3">
              {[
                ["Revenue", "$4.84M", "+18.4%"],
                ["EBITDA", "$786K", "+9.2%"],
                ["Cash", "$1.12M", "Healthy"],
              ].map(([label, value, note]) => (
                <div key={label} className="rounded-2xl border bg-white p-4">
                  <p className="text-xs text-slate-500">{label}</p>
                  <p className="mt-2 text-xl font-bold">{value}</p>
                  <p className="mt-1 text-xs text-emerald-600">{note}</p>
                </div>
              ))}
            </div>
            <div className="mt-4 flex h-44 items-end gap-2 rounded-2xl bg-slate-950 p-5">
              {[42, 58, 47, 70, 64, 78, 88, 82, 96, 91, 108, 118].map((height, index) => (
                <div
                  key={index}
                  className="flex-1 rounded-t-md bg-gradient-to-t from-indigo-500 to-sky-300"
                  style={{ height: `${Math.min(height, 120)}px` }}
                />
              ))}
            </div>
          </div>

          <div className="absolute bottom-8 right-0 w-[72%] rotate-[3deg] rounded-[26px] border border-white/80 bg-white p-5 shadow-2xl">
            <div className="flex items-center gap-3">
              <div className="flex size-11 items-center justify-center rounded-2xl bg-indigo-50 text-indigo-600">
                <BrainCircuit className="size-5" />
              </div>
              <div>
                <p className="font-bold">AI CFO Brief</p>
                <p className="text-xs text-slate-500">Updated from the latest actuals</p>
              </div>
            </div>
            <div className="mt-4 space-y-3">
              <div className="rounded-xl bg-amber-50 p-3 text-sm text-amber-900">
                Gross margin reduced by 2.8 percentage points in the latest month.
              </div>
              <div className="rounded-xl bg-emerald-50 p-3 text-sm text-emerald-900">
                Receivable collection improved by 6 days against the prior quarter.
              </div>
            </div>
          </div>

          <div className="absolute right-8 top-0 flex items-center gap-2 rounded-2xl border bg-white px-4 py-3 shadow-xl animate-soft-bob">
            <ShieldCheck className="size-5 text-indigo-600" />
            <div>
              <p className="text-sm font-semibold">Controlled finance data</p>
              <p className="text-xs text-slate-500">Validated and traceable</p>
            </div>
          </div>
        </div>
      </section>

      
      <section className="relative z-10 overflow-hidden border-y border-slate-200/70 bg-slate-950 py-5 text-white">
        <div className="animate-feature-marquee flex min-w-max gap-10 whitespace-nowrap text-sm font-semibold tracking-wide text-slate-200">
          {["Financial Statements","Branch Consolidation","Business Analytics","Working Capital","Budgets & Forecasts","AI CFO","Board Reports","Board Packs","PowerPoint Export","Audit-ready History"].concat(["Financial Statements","Branch Consolidation","Business Analytics","Working Capital","Budgets & Forecasts","AI CFO","Board Reports","Board Packs","PowerPoint Export","Audit-ready History"]).map((item,index)=><span key={`${item}-${index}`} className="flex items-center gap-2"><Sparkles className="size-3 text-indigo-300"/>{item}</span>)}
        </div>
      </section>

<section className="relative z-10 border-y bg-white/80 py-20 backdrop-blur">
        <div className="mx-auto max-w-7xl px-6 lg:px-8">
          <div className="mx-auto max-w-3xl text-center">
            <p className="text-sm font-bold uppercase tracking-[0.2em] text-indigo-600">One finance operating system</p>
            <h2 className="mt-4 text-4xl font-black tracking-tight">From ledger upload to board conversation</h2>
            <p className="mt-5 text-lg text-slate-600">
              Replace disconnected spreadsheets with a repeatable reporting and planning workflow.
            </p>
          </div>

          <div className="mt-14 grid gap-5 md:grid-cols-2 lg:grid-cols-4">
            {features.map(({ icon: Icon, title, text }, index) => (
              <article
                key={title}
                className="group rounded-3xl border bg-white p-6 shadow-sm transition duration-300 hover:-translate-y-2 hover:shadow-2xl"
                style={{ animationDelay: `${index * 90}ms` }}
              >
                <div className="flex size-12 items-center justify-center rounded-2xl bg-slate-950 text-white transition group-hover:rotate-3 group-hover:scale-110">
                  <Icon className="size-5" />
                </div>
                <h3 className="mt-6 text-lg font-bold">{title}</h3>
                <p className="mt-3 leading-7 text-slate-600">{text}</p>
              </article>
            ))}
          </div>
        </div>
      </section>

      <section className="relative z-10 mx-auto max-w-7xl px-6 py-24 lg:px-8">
        <div className="overflow-hidden rounded-[36px] bg-slate-950 px-7 py-12 text-white shadow-2xl sm:px-12 lg:flex lg:items-center lg:justify-between">
          <div>
            <div className="flex items-center gap-2 text-indigo-300">
              <Building2 className="size-5" />
              Built for growing multi-entity and branch-based businesses
            </div>
            <h2 className="mt-4 max-w-3xl text-3xl font-black tracking-tight sm:text-4xl">
              Set up the company once. Keep reporting, forecasting and board packs moving every month.
            </h2>
          </div>
          <Link
            href="/signup"
            className="mt-8 inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-white px-8 font-bold text-slate-950 transition hover:-translate-y-1 lg:mt-0"
          >
            Start company setup
            <ArrowRight className="size-5" />
          </Link>
        </div>
      </section>
    </main>
  );
}
