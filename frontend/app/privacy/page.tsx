import Link from "next/link";
import { ArrowLeft, BarChart3, Database, Eye, ShieldCheck } from "lucide-react";

export default function PrivacyPage() {
  return <main className="min-h-screen bg-slate-950 text-white">
    <header className="border-b border-white/10"><div className="mx-auto flex max-w-5xl items-center justify-between px-5 py-5"><Link href="/" className="flex items-center gap-2 font-black"><BarChart3 className="size-5"/>FinCruiz</Link><Link href="/book-demo" className="text-sm font-semibold text-slate-300 hover:text-white">Book a demo</Link></div></header>
    <section className="mx-auto max-w-5xl px-5 py-16 sm:py-24">
      <Link href="/" className="inline-flex items-center gap-2 text-sm text-slate-400 hover:text-white"><ArrowLeft className="size-4"/>Back</Link>
      <div className="mt-10 max-w-3xl"><p className="text-xs font-black uppercase tracking-[.2em] text-indigo-300">Privacy</p><h1 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">A clear product-level view of how data is used.</h1><p className="mt-5 text-base leading-8 text-slate-300">This is a product transparency page, not a substitute for the final legal Privacy Policy, customer agreement or data-processing agreement that should be reviewed for each launch market.</p></div>
      <div className="mt-10 space-y-4">
        <section className="rounded-2xl border border-white/10 bg-white/[.05] p-6"><Database className="size-5 text-indigo-300"/><h2 className="mt-4 text-xl font-black">Workspace and finance data</h2><p className="mt-2 text-sm leading-7 text-slate-400">FinCruiz processes workspace profile, membership, integration/import and financial data needed to provide reporting, planning and management features. Source identifiers and report context are retained where needed for reconciliation and traceability.</p></section>
        <section className="rounded-2xl border border-white/10 bg-white/[.05] p-6"><Eye className="size-5 text-sky-300"/><h2 className="mt-4 text-xl font-black">AI features</h2><p className="mt-2 text-sm leading-7 text-slate-400">The AI CFO is designed to interpret prepared finance context and evidence rather than act as the system of record. The production deployment should document the configured AI provider, retention settings and contractual data terms before customer launch.</p></section>
        <section className="rounded-2xl border border-white/10 bg-white/[.05] p-6"><ShieldCheck className="size-5 text-emerald-300"/><h2 className="mt-4 text-xl font-black">Sales enquiries</h2><p className="mt-2 text-sm leading-7 text-slate-400">Information submitted through Book a Demo is used to respond to the enquiry and manage the sales conversation. The demo form does not create a paid subscription or start billing.</p></section>
      </div>
      <div className="mt-10 rounded-2xl border border-white/10 bg-white/[.03] p-5 text-sm leading-7 text-slate-400"><strong className="text-white">Before paid launch:</strong> publish the reviewed legal privacy terms for the operating entity and launch countries, including retention, subprocessors, contact route, data-subject rights and cross-border processing where applicable.</div>
      <div className="mt-8 flex flex-wrap gap-3"><Link href="/security" className="rounded-xl border border-white/10 px-4 py-3 text-sm font-bold">Security</Link><Link href="/trust" className="rounded-xl bg-white px-4 py-3 text-sm font-black text-slate-950">Trust details</Link></div>
    </section>
  </main>;
}
