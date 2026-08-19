"use client";

import Link from "next/link";
import { FormEvent, useEffect, useState } from "react";
import {
  Activity,
  ArrowRight,
  BadgeCheck,
  BarChart3,
  BrainCircuit,
  Building2,
  CheckCircle2,
  CircleDollarSign,
  Database,
  FileBarChart,
  GitBranch,
  Globe2,
  KeyRound,
  LineChart,
  Loader2,
  PlayCircle,
  Presentation,
  Send,
  ShieldCheck,
  Sparkles,
  TrendingUp,
  UploadCloud,
  Users,
  WandSparkles,
} from "lucide-react";

import { demoService, type DemoAnswer } from "@/services/demo-service";
import { marketingService } from "@/services/marketing-service";

type PersonaKey = "owner" | "finance" | "advisor";

type Persona = {
  label: string;
  eyebrow: string;
  headline: string;
  body: string;
  questions: string[];
  outcomes: string[];
};

const personas: Record<PersonaKey, Persona> = {
  owner: {
    label: "Owner / CEO",
    eyebrow: "See the business, not the ledger",
    headline: "Land on the dashboard, ask the question, make the decision.",
    body: "FinCruiz turns finance data into a concise management view so leaders do not need to hunt through reports before every decision.",
    questions: [
      "What should I focus on today?",
      "Why is cash getting tighter?",
      "Which branch is underperforming?",
    ],
    outcomes: ["Executive priorities", "Cash visibility", "Scenario-backed decisions"],
  },
  finance: {
    label: "CFO / Finance",
    eyebrow: "One governed finance layer",
    headline: "Move from reporting the past to challenging the next decision.",
    body: "Keep the calculations deterministic, then use AI to explain movements, surface evidence and route management questions into forecasting and modelling.",
    questions: [
      "Where are we losing margin?",
      "What happens if revenue grows 10%?",
      "Build a 12-month forecast.",
    ],
    outcomes: ["Evidence-backed narrative", "Three-way forecasting", "Board-ready reporting"],
  },
  advisor: {
    label: "Accountant / Advisor",
    eyebrow: "Turn compliance data into advisory conversations",
    headline: "Give clients a management story they can actually act on.",
    body: "Use the same structured finance data to explain performance, compare branches, test assumptions and prepare clearer management conversations.",
    questions: [
      "What are the three biggest risks?",
      "How can we improve working capital?",
      "What should the board discuss next?",
    ],
    outcomes: ["Faster client insight", "Clear evidence trail", "Repeatable management reporting"],
  },
};

const managementLoop = [
  {
    title: "Bring the numbers together",
    text: "Connect or import finance data, branches and planning inputs into one governed workspace.",
    icon: Database,
  },
  {
    title: "Understand what changed",
    text: "Statements, KPIs and management signals are calculated before AI writes the story.",
    icon: LineChart,
  },
  {
    title: "Ask the business",
    text: "Ask FinCruiz in plain English and see the evidence behind the answer.",
    icon: BrainCircuit,
  },
  {
    title: "Model the decision",
    text: "Test hiring, pricing, growth, capex and working-capital assumptions before acting.",
    icon: WandSparkles,
  },
  {
    title: "Report with confidence",
    text: "Carry the same governed numbers into management and board reporting.",
    icon: Presentation,
  },
];

const capabilities = [
  {
    title: "Management reporting",
    text: "P&L, balance sheet, KPIs and management commentary from one prepared finance layer.",
    icon: FileBarChart,
    href: "/demo#reporting",
    event: "homepage_reporting_cta_clicked",
    cta: "See reporting",
  },
  {
    title: "Conversational BI",
    text: "Ask why something moved, where the pressure is coming from and what deserves attention.",
    icon: BrainCircuit,
    href: "/demo#ask-demo",
    event: "homepage_capability_demo_clicked",
    cta: "Ask the demo",
  },
  {
    title: "Branch intelligence",
    text: "Compare locations without building another spreadsheet or losing the consolidated view.",
    icon: GitBranch,
    href: "/demo#branches",
    event: "homepage_capability_demo_clicked",
    cta: "Compare branches",
  },
  {
    title: "Forecasting & planning",
    text: "Start from historical account behaviour, then shape targets and assumptions at management level.",
    icon: TrendingUp,
    href: "/demo#forecasting",
    event: "homepage_forecasting_cta_clicked",
    cta: "See forecasting",
  },
  {
    title: "Decision simulation",
    text: "Model the cash and profit effect of management choices before you commit.",
    icon: CircleDollarSign,
    href: "/demo#decision",
    event: "homepage_capability_demo_clicked",
    cta: "Model a decision",
  },
  {
    title: "Board communication",
    text: "Move from raw finance outputs to concise management and board-ready explanations.",
    icon: Building2,
    href: "/demo#board",
    event: "homepage_reporting_cta_clicked",
    cta: "See board story",
  },
];

const faqs = [
  {
    q: "Does the AI calculate my financial statements?",
    a: "No. FinCruiz prepares and validates the finance context first. AI is used to explain, investigate and help management navigate the evidence rather than replacing the underlying calculations.",
  },
  {
    q: "Can FinCruiz replace my accounting system?",
    a: "FinCruiz is designed as a management-intelligence layer around your finance data. It is not positioned as a replacement for the accounting system that records your books and statutory transactions.",
  },
  {
    q: "Can I use multiple branches or entities?",
    a: "The product supports branch-aware analysis and consolidated management views. Workspace limits depend on the plan and deployment configuration.",
  },
  {
    q: "What can Ask FinCruiz see?",
    a: "Inside a customer workspace, answers are built from the governed company context available to that workspace and role. The public demo uses a completely separate synthetic company dataset.",
  },
  {
    q: "How do users get access to a company?",
    a: "Existing company access is invitation-based. Roles and permissions control what each member can see and manage inside the workspace.",
  },
];

const homepageStructuredData = {
  "@context": "https://schema.org",
  "@graph": [
    {
      "@type": "SoftwareApplication",
      name: "FinCruiz",
      applicationCategory: "BusinessApplication",
      operatingSystem: "Web",
      description:
        "AI CFO and management intelligence software for evidence-backed reporting, business questions, forecasting, scenario modelling and board communication.",
    },
    {
      "@type": "FAQPage",
      mainEntity: faqs.map((item) => ({
        "@type": "Question",
        name: item.q,
        acceptedAnswer: {
          "@type": "Answer",
          text: item.a,
        },
      })),
    },
  ],
};

export default function HomePage() {
  const [question, setQuestion] = useState("Why is cash getting tighter?");
  const [answer, setAnswer] = useState<DemoAnswer | null>(null);
  const [asking, setAsking] = useState(false);
  const [persona, setPersona] = useState<PersonaKey>("owner");

  useEffect(() => {
    marketingService.track("homepage_viewed");
  }, []);

  async function ask(value: string) {
    const q = value.trim();
    if (!q || asking) return;
    setQuestion(q);
    setAsking(true);
    marketingService.track("homepage_ai_question_submitted", { source: "homepage" });
    try {
      setAnswer(await demoService.ask(q));
    } catch {
      setAnswer({
        answer:
          "The interactive evidence service is temporarily unavailable. Open the guided demo to continue with the synthetic Nova Retail story.",
        mode: "error",
        evidence: [],
        confidence: "low",
        confidence_reason: "Demo service unavailable.",
        suggested_questions: [],
      });
    } finally {
      setAsking(false);
    }
  }

  function submit(event: FormEvent) {
    event.preventDefault();
    void ask(question);
  }

  function choosePersona(next: PersonaKey) {
    setPersona(next);
    marketingService.track("homepage_persona_changed", { persona: next });
  }

  const activePersona = personas[persona];
  const evidence = answer?.evidence ?? [];

  return (
    <main className="min-h-screen overflow-hidden bg-[#f7f8fc] text-slate-950">
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(homepageStructuredData) }}
      />
      <div className="pointer-events-none fixed inset-0 landing-aurora" />

      <header className="sticky top-0 z-50 border-b border-white/80 bg-white/85 backdrop-blur-xl">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-5 py-4 lg:px-8">
          <Link href="/" className="flex items-center gap-3">
            <span className="flex size-11 items-center justify-center rounded-2xl bg-slate-950 text-white shadow-sm">
              <BarChart3 className="size-5" />
            </span>
            <span>
              <span className="block text-lg font-black tracking-tight">FinCruiz</span>
              <span className="block text-[11px] font-medium text-slate-500">AI CFO & management intelligence</span>
            </span>
          </Link>

          <nav className="hidden items-center gap-1 lg:flex">
            <a href="#product" className="rounded-xl px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-100">Product</a>
            <a href="#teams" className="rounded-xl px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-100">For teams</a>
            <a href="#trust" className="rounded-xl px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-100">Trust</a>
            <a href="#faq" className="rounded-xl px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-100">FAQ</a>
          </nav>

          <div className="flex items-center gap-2">
            <Link
              href="/pricing"
              onClick={() => marketingService.track("homepage_pricing_cta_clicked", { source: "nav" })}
              className="hidden rounded-xl px-4 py-2 text-sm font-semibold text-slate-600 sm:block"
            >
              Pricing
            </Link>
            <Link href="/login" className="rounded-xl px-3 py-2 text-sm font-semibold sm:px-4">Sign in</Link>
            <Link
              href="/signup"
              onClick={() => marketingService.track("homepage_hero_signup_clicked", { source: "nav" })}
              className="hidden rounded-xl bg-slate-950 px-5 py-2.5 text-sm font-bold text-white md:inline-flex"
            >
              Create workspace
            </Link>
          </div>
        </div>
      </header>

      <section className="relative z-10 mx-auto grid max-w-7xl gap-12 px-5 pb-20 pt-14 lg:grid-cols-[0.92fr_1.08fr] lg:items-center lg:px-8 lg:pt-24">
        <div>
          <div className="inline-flex items-center gap-2 rounded-full border border-indigo-100 bg-white/90 px-4 py-2 text-sm font-bold shadow-sm">
            <Sparkles className="size-4 text-indigo-600" />
            Finance tells you what happened. FinCruiz helps you decide what happens next.
          </div>

          <h1 className="mt-7 max-w-3xl text-5xl font-black leading-[1.01] tracking-[-.055em] sm:text-6xl lg:text-[68px]">
            Turn financial data into
            <span className="block bg-gradient-to-r from-indigo-600 via-violet-600 to-sky-500 bg-clip-text text-transparent">
              management decisions.
            </span>
          </h1>

          <p className="mt-6 max-w-2xl text-lg leading-8 text-slate-600">
            FinCruiz brings reporting, branches, budgets and scenarios into one evidence-backed workspace. Ask a business question in plain English, see the numbers behind the answer, and model the next move.
          </p>

          <div className="mt-8 flex flex-col gap-3 sm:flex-row">
            <Link
              href="/demo"
              onClick={() => marketingService.track("homepage_hero_demo_clicked", { source: "hero" })}
              className="inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-indigo-600 px-7 font-black text-white shadow-[0_18px_45px_rgba(79,70,229,.24)] hover:-translate-y-0.5"
            >
              <PlayCircle className="size-5" />
              See the guided demo
              <ArrowRight className="size-4" />
            </Link>
            <Link
              href="/signup"
              onClick={() => marketingService.track("homepage_hero_signup_clicked", { source: "hero" })}
              className="inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-slate-950 px-7 font-bold text-white hover:-translate-y-0.5"
            >
              Use my business data
              <ArrowRight className="size-4" />
            </Link>
          </div>

          <div className="mt-7 grid max-w-2xl gap-3 text-sm text-slate-600 sm:grid-cols-2">
            {[
              "Synthetic demo — no customer data needed",
              "Evidence before AI narrative",
              "Invitation-based company access",
              "Branch-aware reporting and planning",
            ].map((item) => (
              <span key={item} className="flex items-center gap-2">
                <CheckCircle2 className="size-4 shrink-0 text-emerald-600" />
                {item}
              </span>
            ))}
          </div>
        </div>

        <div className="rounded-[34px] border border-white/90 bg-white/92 p-4 shadow-[0_40px_120px_rgba(30,41,59,.15)] backdrop-blur-xl sm:p-6">
          <div className="rounded-[26px] bg-slate-950 p-5 text-white sm:p-6">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div>
                <p className="text-[11px] font-black uppercase tracking-[.18em] text-indigo-300">Executive dashboard</p>
                <h2 className="mt-2 text-2xl font-black">Morning, Nova Retail.</h2>
              </div>
              <span className="rounded-full border border-emerald-300/15 bg-emerald-300/10 px-3 py-1.5 text-xs font-bold text-emerald-200">Financial confidence · High</span>
            </div>

            <div className="mt-5 grid gap-3 sm:grid-cols-3">
              {[
                ["Revenue", "₹24.80M", "+9.4% YoY"],
                ["Net profit", "₹4.12M", "16.6% margin"],
                ["Cash", "₹6.21M", "AR needs attention"],
              ].map(([label, value, note]) => (
                <div key={label} className="rounded-2xl border border-white/10 bg-white/[.055] p-4">
                  <p className="text-xs text-slate-400">{label}</p>
                  <p className="mt-2 text-xl font-black">{value}</p>
                  <p className="mt-1 text-xs text-indigo-200">{note}</p>
                </div>
              ))}
            </div>

            <div className="mt-4 rounded-2xl border border-indigo-300/15 bg-indigo-300/[.08] p-4">
              <div className="flex items-start gap-3">
                <span className="flex size-9 shrink-0 items-center justify-center rounded-xl bg-indigo-500 text-white"><BrainCircuit className="size-4" /></span>
                <div>
                  <p className="text-xs font-black uppercase tracking-[.14em] text-indigo-200">Ask FinCruiz — Your AI CFO</p>
                  <p className="mt-2 font-bold">What should I focus on today?</p>
                  <p className="mt-2 text-sm leading-6 text-slate-300">Working-capital conversion and West branch margin are the two issues most likely to affect the next management decision.</p>
                </div>
              </div>
              <div className="mt-4 grid gap-2 sm:grid-cols-2">
                <div className="rounded-xl border border-white/10 bg-slate-950/45 p-3 text-xs text-slate-300"><BadgeCheck className="mr-1.5 inline size-3.5 text-emerald-300" />Evidence: AR ageing · branch P&L · cash forecast</div>
                <div className="rounded-xl border border-white/10 bg-slate-950/45 p-3 text-xs text-slate-300"><WandSparkles className="mr-1.5 inline size-3.5 text-violet-300" />Next: model collections or branch margin</div>
              </div>
            </div>
          </div>

          <div className="mt-4 grid gap-3 sm:grid-cols-2">
            <div className="rounded-2xl border bg-slate-50 p-4">
              <p className="text-xs font-black uppercase tracking-[.14em] text-slate-500">Management priorities</p>
              <div className="mt-3 space-y-2">
                {["1 · Reduce overdue receivables", "2 · Investigate West margin", "3 · Protect downside cash buffer"].map((item) => (
                  <div key={item} className="rounded-xl bg-white px-3 py-2.5 text-sm font-semibold shadow-sm">{item}</div>
                ))}
              </div>
            </div>
            <div className="rounded-2xl border bg-slate-50 p-4">
              <p className="text-xs font-black uppercase tracking-[.14em] text-slate-500">Decision simulator</p>
              <p className="mt-3 text-sm font-semibold">Hire 3 people?</p>
              <div className="mt-3 flex items-end gap-2">
                {[72, 61, 48, 66].map((height, index) => (
                  <span key={index} className="flex-1 rounded-t-lg bg-indigo-100" style={{ height: `${height}px` }} />
                ))}
              </div>
              <p className="mt-3 text-xs leading-5 text-slate-500">Base case remains above the management cash buffer. Downside collections do not.</p>
            </div>
          </div>
        </div>
      </section>

      <section className="relative z-10 border-y border-slate-800 bg-slate-950 py-5 text-white">
        <div className="mx-auto flex max-w-7xl flex-wrap justify-center gap-x-8 gap-y-3 px-5 text-sm font-semibold text-slate-200">
          {["Evidence-backed AI", "Conversational BI", "Multi-branch intelligence", "Native planning", "Three-way forecasting", "Board reporting"].map((item) => (
            <span key={item} className="flex items-center gap-2"><Sparkles className="size-3 text-indigo-300" />{item}</span>
          ))}
        </div>
      </section>

      <section id="product" className="relative z-10 mx-auto max-w-7xl px-5 py-24 lg:px-8">
        <div className="max-w-3xl">
          <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-600">One management loop</p>
          <h2 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">From accounting data to action — without changing tools every step.</h2>
          <p className="mt-5 text-lg leading-8 text-slate-600">FinCruiz is designed to keep understanding, investigation, planning, decisions and reporting connected to the same finance context.</p>
        </div>

        <div className="mt-10 grid gap-4 lg:grid-cols-5">
          {managementLoop.map((item, index) => {
            const Icon = item.icon;
            return (
              <div key={item.title} className="rounded-[26px] border bg-white p-5 shadow-sm transition hover:-translate-y-1 hover:shadow-lg">
                <div className="flex items-center justify-between">
                  <span className="flex size-10 items-center justify-center rounded-2xl bg-indigo-50 text-indigo-700"><Icon className="size-5" /></span>
                  <span className="text-xs font-black text-slate-300">0{index + 1}</span>
                </div>
                <h3 className="mt-5 font-black">{item.title}</h3>
                <p className="mt-2 text-sm leading-6 text-slate-600">{item.text}</p>
              </div>
            );
          })}
        </div>
      </section>

      <section id="teams" className="relative z-10 border-y bg-white/75 py-24">
        <div className="mx-auto max-w-7xl px-5 lg:px-8">
          <div className="grid gap-12 lg:grid-cols-[0.72fr_1.28fr] lg:items-start">
            <div>
              <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-600">Built around the person asking</p>
              <h2 className="mt-3 text-4xl font-black tracking-tight">Different roles. One company truth.</h2>
              <p className="mt-5 text-base leading-7 text-slate-600">The product should feel useful whether the user starts with a management question, a finance review or an advisory conversation.</p>

              <div className="mt-7 grid gap-2">
                {(Object.keys(personas) as PersonaKey[]).map((key) => (
                  <button
                    key={key}
                    type="button"
                    onClick={() => choosePersona(key)}
                    className={`rounded-2xl border px-4 py-4 text-left transition ${persona === key ? "border-indigo-300 bg-indigo-50 shadow-sm" : "bg-white hover:border-slate-300"}`}
                  >
                    <span className="flex items-center justify-between gap-3">
                      <span className="font-bold">{personas[key].label}</span>
                      <ArrowRight className={`size-4 ${persona === key ? "text-indigo-600" : "text-slate-300"}`} />
                    </span>
                  </button>
                ))}
              </div>
            </div>

            <div className="rounded-[32px] border bg-slate-950 p-6 text-white shadow-xl sm:p-8">
              <p className="text-xs font-black uppercase tracking-[.18em] text-indigo-300">{activePersona.eyebrow}</p>
              <h3 className="mt-3 text-3xl font-black tracking-tight">{activePersona.headline}</h3>
              <p className="mt-4 max-w-2xl leading-7 text-slate-300">{activePersona.body}</p>

              <div className="mt-7 grid gap-4 md:grid-cols-[1.1fr_.9fr]">
                <div className="rounded-2xl border border-white/10 bg-white/[.05] p-5">
                  <p className="text-xs font-black uppercase tracking-[.14em] text-slate-400">Questions they can start with</p>
                  <div className="mt-3 space-y-2">
                    {activePersona.questions.map((item) => (
                      <button key={item} type="button" onClick={() => void ask(item)} className="flex w-full items-center justify-between gap-3 rounded-xl border border-white/10 bg-slate-950/50 px-3 py-3 text-left text-sm font-semibold hover:border-indigo-300/30 hover:bg-indigo-300/10">
                        {item}<ArrowRight className="size-3.5 shrink-0 text-indigo-300" />
                      </button>
                    ))}
                  </div>
                </div>
                <div className="rounded-2xl border border-white/10 bg-white/[.05] p-5">
                  <p className="text-xs font-black uppercase tracking-[.14em] text-slate-400">What they get back</p>
                  <div className="mt-3 space-y-3">
                    {activePersona.outcomes.map((item) => (
                      <div key={item} className="flex items-start gap-2 text-sm text-slate-200"><CheckCircle2 className="mt-0.5 size-4 shrink-0 text-emerald-300" />{item}</div>
                    ))}
                  </div>
                  <Link href="/demo" onClick={() => marketingService.track("homepage_product_tour_clicked", { persona })} className="mt-6 inline-flex items-center gap-2 text-sm font-black text-indigo-200 hover:text-white">Show this workflow <ArrowRight className="size-4" /></Link>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>

      <section className="relative z-10 mx-auto max-w-7xl px-5 py-24 lg:px-8">
        <div className="mx-auto max-w-3xl text-center">
          <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-600">Try the intelligence layer</p>
          <h2 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">Ask FinCruiz before you give it any of your data.</h2>
          <p className="mt-5 text-lg leading-8 text-slate-600">The public demo is isolated from customer workspaces and uses a fixed synthetic multi-branch company.</p>
        </div>

        <div className="mx-auto mt-10 max-w-5xl rounded-[32px] border bg-white p-5 shadow-[0_28px_90px_rgba(30,41,59,.12)] sm:p-7">
          <form onSubmit={submit} className="flex flex-col gap-3 sm:flex-row">
            <div className="flex min-h-14 min-w-0 flex-1 items-center gap-3 rounded-2xl border bg-slate-50 px-4 focus-within:border-indigo-300 focus-within:ring-4 focus-within:ring-indigo-50">
              <BrainCircuit className="size-5 shrink-0 text-indigo-600" />
              <input value={question} onChange={(event) => setQuestion(event.target.value)} placeholder="e.g. Why is cash getting tighter?" className="min-w-0 flex-1 bg-transparent text-sm outline-none" />
            </div>
            <button disabled={!question.trim() || asking} className="inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-slate-950 px-6 font-bold text-white disabled:opacity-50">
              {asking ? <Loader2 className="size-4 animate-spin" /> : <Send className="size-4" />}
              Ask FinCruiz
            </button>
          </form>

          <div className="mt-3 flex flex-wrap gap-2">
            {["What should management focus on?", "Which branch needs attention?", "Can we afford to hire 3 people?", "What happens if revenue grows 10%?"].map((item) => (
              <button key={item} type="button" onClick={() => void ask(item)} className="rounded-full bg-indigo-50 px-3 py-1.5 text-xs font-semibold text-indigo-700 hover:bg-indigo-100">{item}</button>
            ))}
          </div>

          {answer ? (
            <div className="mt-6 grid gap-4 lg:grid-cols-[1.25fr_.75fr]">
              <div className="rounded-2xl bg-slate-950 p-5 text-white">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <p className="text-xs font-black uppercase tracking-[.14em] text-indigo-300">Management answer</p>
                  <span className="rounded-full bg-emerald-300/10 px-2.5 py-1 text-[10px] font-bold uppercase text-emerald-200">{answer.confidence} confidence</span>
                </div>
                <p className="mt-3 whitespace-pre-wrap text-sm leading-7 text-slate-100">{answer.answer}</p>
                <p className="mt-4 text-xs leading-5 text-slate-400"><ShieldCheck className="mr-1.5 inline size-3.5" />{answer.confidence_reason}</p>
                <div className="mt-5 flex flex-wrap gap-2">
                  <Link href="/demo" className="inline-flex items-center gap-2 rounded-xl bg-white px-4 py-2.5 text-xs font-black text-slate-950">Open full guided demo <ArrowRight className="size-3.5" /></Link>
                  <Link href="/signup" onClick={() => marketingService.track("homepage_ai_signup_clicked")} className="inline-flex items-center gap-2 rounded-xl border border-white/15 px-4 py-2.5 text-xs font-bold">Ask about my business <ArrowRight className="size-3.5" /></Link>
                </div>
              </div>

              <div className="rounded-2xl border bg-slate-50 p-5">
                <p className="text-xs font-black uppercase tracking-[.14em] text-slate-500">Evidence used</p>
                {evidence.length ? (
                  <div className="mt-3 space-y-2">
                    {evidence.slice(0, 4).map((item, index) => (
                      <div key={`${item.label}-${index}`} className="rounded-xl bg-white p-3 shadow-sm">
                        <div className="flex items-start justify-between gap-3">
                          <span className="text-xs text-slate-500">{item.label}</span>
                          <span className="text-sm font-black">{item.value}</span>
                        </div>
                        <p className="mt-1 text-[10px] text-slate-400">{item.source}</p>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="mt-3 text-sm leading-6 text-slate-500">No supported evidence was available for that question — which is exactly when the demo should refuse to invent an answer.</p>
                )}
              </div>
            </div>
          ) : null}
        </div>
      </section>

      <section className="relative z-10 border-y bg-slate-950 py-24 text-white">
        <div className="mx-auto max-w-7xl px-5 lg:px-8">
          <div className="flex flex-col justify-between gap-6 lg:flex-row lg:items-end">
            <div className="max-w-3xl">
              <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-300">Capability without the maze</p>
              <h2 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">Start with the question. FinCruiz takes you to the right analysis.</h2>
            </div>
            <Link href="/pricing" onClick={() => marketingService.track("homepage_pricing_cta_clicked", { source: "capabilities" })} className="inline-flex items-center gap-2 text-sm font-black text-indigo-200 hover:text-white">See regional pricing <ArrowRight className="size-4" /></Link>
          </div>

          <div className="mt-10 grid gap-4 md:grid-cols-2 xl:grid-cols-3">
            {capabilities.map((item) => {
              const Icon = item.icon;
              return (
                <Link
                  key={item.title}
                  href={item.href}
                  onClick={() => marketingService.track(item.event, { capability: item.title })}
                  className="group rounded-[26px] border border-white/10 bg-white/[.05] p-6 hover:-translate-y-1 hover:border-indigo-300/25 hover:bg-white/[.075]"
                >
                  <span className="flex size-11 items-center justify-center rounded-2xl bg-indigo-400/10 text-indigo-200"><Icon className="size-5" /></span>
                  <h3 className="mt-5 text-xl font-black">{item.title}</h3>
                  <p className="mt-2 text-sm leading-6 text-slate-400">{item.text}</p>
                  <span className="mt-5 inline-flex items-center gap-2 text-sm font-bold text-indigo-200">{item.cta}<ArrowRight className="size-4 transition group-hover:translate-x-1" /></span>
                </Link>
              );
            })}
          </div>
        </div>
      </section>

      <section className="relative z-10 mx-auto max-w-7xl px-5 py-24 lg:px-8">
        <div className="grid gap-6 lg:grid-cols-2">
          <div className="rounded-[32px] border bg-white p-7 shadow-sm sm:p-8">
            <span className="flex size-12 items-center justify-center rounded-2xl bg-sky-50 text-sky-700"><UploadCloud className="size-6" /></span>
            <p className="mt-5 text-xs font-black uppercase tracking-[.18em] text-sky-700">Connect or import</p>
            <h2 className="mt-2 text-3xl font-black">Meet finance teams where their data already lives.</h2>
            <p className="mt-4 leading-7 text-slate-600">FinCruiz is designed for finance data coming from accounting integrations and structured file workflows, so the management layer does not depend on one accounting vendor.</p>
            <div className="mt-6 flex flex-wrap gap-2">
              {["Xero", "Zoho", "Tally", "CSV / file import"].map((item) => <span key={item} className="rounded-full border bg-slate-50 px-3 py-2 text-sm font-bold text-slate-700">{item}</span>)}
            </div>
            <p className="mt-4 text-xs leading-5 text-slate-500">Integration availability depends on the environment and configured connection. File-based workflows remain available where supported.</p>
          </div>

          <div id="trust" className="rounded-[32px] border bg-white p-7 shadow-sm sm:p-8">
            <span className="flex size-12 items-center justify-center rounded-2xl bg-emerald-50 text-emerald-700"><ShieldCheck className="size-6" /></span>
            <p className="mt-5 text-xs font-black uppercase tracking-[.18em] text-emerald-700">Trust is part of the workflow</p>
            <h2 className="mt-2 text-3xl font-black">AI should explain the numbers — not quietly become the numbers.</h2>
            <div className="mt-6 grid gap-3 sm:grid-cols-2">
              {[
                [KeyRound, "Invitation-based access", "Existing company access is tied to invited identity and role."],
                [Users, "Role-based permissions", "Owner, finance and viewer experiences can be separated."],
                [BadgeCheck, "Evidence-backed answers", "Finance evidence is prepared before AI interpretation."],
                [Activity, "Audit & controls", "Important workspace actions can be traced and reviewed."],
              ].map(([Icon, title, text]) => {
                const ItemIcon = Icon as typeof ShieldCheck;
                return (
                  <div key={String(title)} className="rounded-2xl border bg-slate-50 p-4">
                    <ItemIcon className="size-4 text-emerald-700" />
                    <p className="mt-3 font-black">{String(title)}</p>
                    <p className="mt-1 text-xs leading-5 text-slate-500">{String(text)}</p>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </section>

      <section id="faq" className="relative z-10 border-y bg-white/75 py-24">
        <div className="mx-auto max-w-5xl px-5 lg:px-8">
          <div className="text-center">
            <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-600">Questions buyers ask</p>
            <h2 className="mt-3 text-4xl font-black tracking-tight">Clear answers before the sales call.</h2>
          </div>
          <div className="mt-10 grid gap-3">
            {faqs.map((item) => (
              <details key={item.q} className="group rounded-2xl border bg-white p-5 shadow-sm open:border-indigo-200">
                <summary className="cursor-pointer list-none font-black">{item.q}</summary>
                <p className="mt-3 max-w-4xl text-sm leading-7 text-slate-600">{item.a}</p>
              </details>
            ))}
          </div>
        </div>
      </section>

      <section className="relative z-10 mx-auto max-w-7xl px-5 py-24 lg:px-8">
        <div className="overflow-hidden rounded-[36px] bg-slate-950 p-8 text-white shadow-[0_30px_90px_rgba(15,23,42,.25)] sm:p-12">
          <div className="grid gap-8 lg:grid-cols-[1.2fr_.8fr] lg:items-end">
            <div>
              <p className="text-xs font-black uppercase tracking-[.2em] text-indigo-300">See the management loop end to end</p>
              <h2 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">Give FinCruiz a question. See the evidence. Then model the decision.</h2>
              <p className="mt-5 max-w-2xl text-base leading-7 text-slate-300">The guided demo is built for a prospect conversation: synthetic company, clear story, interactive questions and no setup required.</p>
            </div>
            <div className="flex flex-col gap-3 sm:flex-row lg:flex-col">
              <Link href="/demo" onClick={() => marketingService.track("homepage_final_demo_clicked")} className="inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl bg-white px-6 font-black text-slate-950"><PlayCircle className="size-5" />Open guided demo<ArrowRight className="size-4" /></Link>
              <Link href="/signup" onClick={() => marketingService.track("homepage_final_signup_clicked")} className="inline-flex min-h-14 items-center justify-center gap-2 rounded-2xl border border-white/15 bg-white/[.05] px-6 font-bold">Create workspace<ArrowRight className="size-4" /></Link>
            </div>
          </div>
        </div>
      </section>

      <footer className="relative z-10 border-t bg-white py-8">
        <div className="mx-auto flex max-w-7xl flex-col gap-5 px-5 text-sm text-slate-500 sm:flex-row sm:items-center sm:justify-between lg:px-8">
          <div className="flex items-center gap-2 font-black text-slate-900"><BarChart3 className="size-4" />FinCruiz</div>
          <div className="flex flex-wrap gap-5">
            <Link href="/demo">Demo</Link>
            <Link href="/pricing" onClick={() => marketingService.track("homepage_pricing_cta_clicked", { source: "footer" })}>Pricing</Link>
            <Link href="/login">Sign in</Link>
          </div>
          <p>Management intelligence for better business decisions.</p>
        </div>
      </footer>
    </main>
  );
}
