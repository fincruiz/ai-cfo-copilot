import Link from "next/link";
import { ArrowLeft, BarChart3, CheckCircle2, LockKeyhole, ShieldCheck } from "lucide-react";

const controls = [
  ["Company-scoped access", "Authenticated requests are resolved against company membership and role so one tenant is not intentionally exposed through another tenant's workspace."],
  ["Role-based permissions", "Owner, finance and viewer-level experiences can be separated, with administrative controls kept behind authenticated workspace access."],
  ["Server-side secrets", "Accounting and billing provider credentials are configured on the backend. Browser code is not intended to receive secret keys or integration encryption material."],
  ["Audit and traceability", "Important finance and workspace actions are designed to retain source, user and timing context so support and finance teams can investigate changes."],
  ["Finance validation before AI", "Financial reports are prepared from validated finance data. AI interpretation is not used as a replacement for the ledger or report calculation engine."],
  ["Fail-closed paid launch gate", "Live payment processing remains controlled by an explicit backend safety switch and provider-mode checks."],
];

export default function SecurityPage() {
  return <main className="min-h-screen bg-slate-950 text-white">
    <header className="border-b border-white/10"><div className="mx-auto flex max-w-5xl items-center justify-between px-5 py-5"><Link href="/" className="flex items-center gap-2 font-black"><BarChart3 className="size-5"/>FinCruiz</Link><Link href="/trust" className="text-sm font-semibold text-slate-300 hover:text-white">Trust centre</Link></div></header>
    <section className="mx-auto max-w-5xl px-5 py-16 sm:py-24">
      <Link href="/" className="inline-flex items-center gap-2 text-sm text-slate-400 hover:text-white"><ArrowLeft className="size-4"/>Back</Link>
      <div className="mt-10 max-w-3xl"><p className="text-xs font-black uppercase tracking-[.2em] text-emerald-300">Security</p><h1 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">Controls built around financial data and tenant boundaries.</h1><p className="mt-5 text-base leading-8 text-slate-300">This page describes product controls currently represented in the FinCruiz application. It does not claim an external certification or audit that has not been completed.</p></div>
      <div className="mt-10 grid gap-4 md:grid-cols-2">{controls.map(([title,body]) => <div key={title} className="rounded-2xl border border-white/10 bg-white/[.05] p-5"><ShieldCheck className="size-5 text-emerald-300"/><h2 className="mt-4 font-black">{title}</h2><p className="mt-2 text-sm leading-6 text-slate-400">{body}</p></div>)}</div>
      <div className="mt-10 rounded-2xl border border-amber-300/15 bg-amber-300/[.05] p-5"><LockKeyhole className="size-5 text-amber-200"/><h2 className="mt-3 font-black">Certification status</h2><p className="mt-2 text-sm leading-6 text-amber-50/80">FinCruiz should not display SOC 2, ISO 27001, PCI-DSS or similar certification claims unless the relevant independent scope has actually been completed and approved for public use. Payment card handling is delegated to configured payment providers rather than being presented here as a FinCruiz certification.</p></div>
      <div className="mt-8 flex flex-wrap gap-3"><Link href="/privacy" className="rounded-xl border border-white/10 px-4 py-3 text-sm font-bold">Privacy</Link><Link href="/trust" className="rounded-xl bg-white px-4 py-3 text-sm font-black text-slate-950">Trust details</Link></div>
    </section>
  </main>;
}
