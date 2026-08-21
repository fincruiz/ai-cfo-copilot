import Link from "next/link";
import { ArrowLeft, BarChart3, CheckCircle2, CircleAlert, DatabaseZap, FileSearch, ShieldCheck } from "lucide-react";

const truth = [
  "Accounting integrations and uploaded files converge on the same canonical validated GL model.",
  "Reporting periods and Data as of context are surfaced separately from AI wording.",
  "Report evidence can drill through toward accounts and transactions instead of stopping at a narrative answer.",
  "Financial calculations remain deterministic; AI is used to explain governed evidence, not silently rewrite it.",
];

export default function TrustPage() {
  return <main className="min-h-screen bg-[#f7f8fc] text-slate-950">
    <header className="border-b bg-white"><div className="mx-auto flex max-w-5xl items-center justify-between px-5 py-5"><Link href="/" className="flex items-center gap-2 font-black"><BarChart3 className="size-5"/>FinCruiz</Link><Link href="/book-demo" className="rounded-xl bg-slate-950 px-4 py-2.5 text-sm font-black text-white">Book a demo</Link></div></header>
    <section className="mx-auto max-w-5xl px-5 py-16 sm:py-24">
      <Link href="/" className="inline-flex items-center gap-2 text-sm text-slate-500 hover:text-slate-950"><ArrowLeft className="size-4"/>Back</Link>
      <div className="mt-10 max-w-3xl"><p className="text-xs font-black uppercase tracking-[.2em] text-indigo-600">Trust centre</p><h1 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">Trust the financial path before trusting the answer.</h1><p className="mt-5 text-base leading-8 text-slate-600">FinCruiz is being hardened around finance truth, tenant isolation, source traceability and explicit launch gates. This page separates implemented product controls from checks that still require production/operator certification.</p></div>
      <div className="mt-10 grid gap-5 md:grid-cols-2">
        <div className="rounded-3xl border bg-white p-6 shadow-sm"><FileSearch className="size-6 text-indigo-600"/><h2 className="mt-4 text-xl font-black">Finance truth</h2><div className="mt-4 space-y-3">{truth.map(item => <p key={item} className="flex gap-3 text-sm leading-6 text-slate-600"><CheckCircle2 className="mt-0.5 size-4 shrink-0 text-emerald-600"/>{item}</p>)}</div></div>
        <div className="rounded-3xl border bg-white p-6 shadow-sm"><ShieldCheck className="size-6 text-emerald-600"/><h2 className="mt-4 text-xl font-black">Access and control</h2><p className="mt-3 text-sm leading-7 text-slate-600">Company membership, role-aware access, server-side integration secrets and audit-oriented workflows form the application control layer. Administrative controls remain secondary to the management experience rather than exposed as the product itself.</p><div className="mt-5"><Link href="/security" className="text-sm font-black text-indigo-600">Read security details →</Link></div></div>
        <div className="rounded-3xl border bg-white p-6 shadow-sm"><DatabaseZap className="size-6 text-sky-600"/><h2 className="mt-4 text-xl font-black">Production operations</h2><p className="mt-3 text-sm leading-7 text-slate-600">Paid launch is gated on production environment configuration, persistent ingestion storage, database/API performance, billing lifecycle evidence, backup/restore, monitoring and a support process. Local code readiness alone does not prove those deployed checks.</p></div>
        <div className="rounded-3xl border border-amber-200 bg-amber-50 p-6"><CircleAlert className="size-6 text-amber-700"/><h2 className="mt-4 text-xl font-black">Claims we will not invent</h2><p className="mt-3 text-sm leading-7 text-amber-900/75">No customer testimonial, case-study outcome, uptime history, penetration-test result or external security certification should appear publicly until it exists, is verified and is approved for publication.</p></div>
      </div>
      <div className="mt-10 flex flex-wrap gap-3"><Link href="/security" className="rounded-xl border bg-white px-4 py-3 text-sm font-bold">Security</Link><Link href="/privacy" className="rounded-xl border bg-white px-4 py-3 text-sm font-bold">Privacy</Link><Link href="/book-demo" className="rounded-xl bg-slate-950 px-4 py-3 text-sm font-black text-white">Discuss your requirements</Link></div>
    </section>
  </main>;
}
