"use client";

import Link from "next/link";
import { FormEvent, useEffect, useState } from "react";
import { ArrowLeft, ArrowRight, BarChart3, CheckCircle2, Loader2, ShieldCheck } from "lucide-react";
import { marketingService } from "@/services/marketing-service";

type Persona = "owner" | "finance" | "advisor" | "";

const personaCopy: Record<Exclude<Persona, "">, { title: string; body: string }> = {
  owner: { title: "Owner / CEO demo", body: "Focus the session on management priorities, cash visibility and decisions you need to make." },
  finance: { title: "CFO / Finance demo", body: "Go deeper on financial truth, controls, reporting, integrations, forecasting and evidence traceability." },
  advisor: { title: "Accountant / Advisor demo", body: "See how governed finance data can support a repeatable advisory and management-reporting workflow." },
};

export default function BookDemoPage() {
  const [persona, setPersona] = useState<Persona>("");
  const [form, setForm] = useState({ name: "", work_email: "", company_name: "", role: "", country: "", team_size: "", message: "", website: "" });
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState("");
  const [submitted, setSubmitted] = useState(false);

  useEffect(() => {
    const requested = new URLSearchParams(window.location.search).get("persona") as Persona | null;
    if (requested && ["owner", "finance", "advisor"].includes(requested)) setPersona(requested);
    marketingService.track("homepage_book_demo_clicked", { source: "book_demo_page" });
  }, []);

  async function submit(event: FormEvent) {
    event.preventDefault();
    if (submitting) return;
    setSubmitting(true); setError("");
    try {
      const result = await marketingService.bookDemo({ ...form, persona: persona || undefined, source_path: window.location.pathname });
      if (!result.accepted) throw new Error("Please try again later or use the guided demo in the meantime.");
      marketingService.track("demo_lead_submitted", { persona: persona || "unspecified" });
      setSubmitted(true);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "We could not submit the request. Please try again.");
    } finally { setSubmitting(false); }
  }

  const copy = persona ? personaCopy[persona] : null;

  return <main className="min-h-screen bg-slate-950 text-white">
    <header className="border-b border-white/10"><div className="mx-auto flex max-w-6xl items-center justify-between px-5 py-5"><Link href="/" className="flex items-center gap-2 font-black"><BarChart3 className="size-5"/>FinCruiz</Link><Link href="/demo" className="text-sm font-semibold text-slate-300 hover:text-white">Open guided demo</Link></div></header>
    <section className="mx-auto grid max-w-6xl gap-10 px-5 py-16 lg:grid-cols-[.85fr_1.15fr] lg:py-24">
      <div>
        <Link href="/" className="inline-flex items-center gap-2 text-sm text-slate-400 hover:text-white"><ArrowLeft className="size-4"/>Back to FinCruiz</Link>
        <p className="mt-10 text-xs font-black uppercase tracking-[.2em] text-indigo-300">Book a demo</p>
        <h1 className="mt-3 text-4xl font-black tracking-tight sm:text-5xl">{copy?.title ?? "See FinCruiz on your business questions."}</h1>
        <p className="mt-5 max-w-xl text-base leading-7 text-slate-300">{copy?.body ?? "Tell us what you want to improve. The demo can be shaped around management reporting, cash, branches, forecasting, integrations or decision modelling."}</p>
        <div className="mt-8 space-y-4 text-sm text-slate-300">
          {["Use your business priorities to shape the conversation.", "Review the implementation path from source data to management reporting.", "No customer testimonials or certifications will be claimed unless they are real and approved."].map((item) => <p key={item} className="flex gap-3"><CheckCircle2 className="mt-0.5 size-4 shrink-0 text-emerald-300"/>{item}</p>)}
        </div>
        <div className="mt-8 rounded-2xl border border-white/10 bg-white/[.04] p-4 text-xs leading-6 text-slate-400"><ShieldCheck className="mr-2 inline size-4 text-indigo-300"/>Your contact details are used to respond to this sales enquiry. See the public Privacy page for how FinCruiz describes data handling.</div>
      </div>

      <div className="rounded-[30px] border border-white/10 bg-white/[.06] p-6 sm:p-8">
        {submitted ? <div className="flex min-h-[480px] flex-col items-center justify-center text-center"><CheckCircle2 className="size-12 text-emerald-300"/><h2 className="mt-5 text-2xl font-black">Demo request received</h2><p className="mt-3 max-w-md text-slate-300">Your enquiry has been captured for follow-up. You can explore the synthetic guided demo now without waiting.</p><Link href="/demo" className="mt-6 inline-flex items-center gap-2 rounded-xl bg-white px-5 py-3 font-black text-slate-950">Open guided demo<ArrowRight className="size-4"/></Link></div> : <form onSubmit={submit} className="space-y-4">
          <div><label className="text-xs font-bold text-slate-300">I’m evaluating FinCruiz as</label><div className="mt-2 grid grid-cols-3 gap-2">{(["owner","finance","advisor"] as const).map((value) => <button key={value} type="button" onClick={() => setPersona(value)} className={`rounded-xl border px-3 py-2 text-xs font-bold ${persona===value?"border-indigo-300 bg-indigo-400/15":"border-white/10 bg-white/[.03]"}`}>{value==="owner"?"Owner / CEO":value==="finance"?"Finance":"Advisor"}</button>)}</div></div>
          <Field label="Name" value={form.name} onChange={(value)=>setForm({...form,name:value})} required />
          <Field label="Work email" type="email" value={form.work_email} onChange={(value)=>setForm({...form,work_email:value})} required />
          <Field label="Company" value={form.company_name} onChange={(value)=>setForm({...form,company_name:value})} required />
          <div className="grid gap-4 sm:grid-cols-2"><Field label="Role" value={form.role} onChange={(value)=>setForm({...form,role:value})}/><Field label="Country" value={form.country} onChange={(value)=>setForm({...form,country:value})}/></div>
          <div><label className="text-xs font-bold text-slate-300">Finance / management team size</label><select value={form.team_size} onChange={(e)=>setForm({...form,team_size:e.target.value})} className="mt-2 h-11 w-full rounded-xl border border-white/10 bg-slate-950 px-3 text-sm"><option value="">Select</option><option>1–5</option><option>6–20</option><option>21–50</option><option>51+</option></select></div>
          <div><label className="text-xs font-bold text-slate-300">What should the demo focus on?</label><textarea value={form.message} onChange={(e)=>setForm({...form,message:e.target.value})} maxLength={1200} className="mt-2 min-h-28 w-full rounded-xl border border-white/10 bg-slate-950 p-3 text-sm outline-none" placeholder="e.g. multi-branch reporting, Xero integration, cash forecasting…"/></div>
          <input tabIndex={-1} autoComplete="off" value={form.website} onChange={(e)=>setForm({...form,website:e.target.value})} className="hidden" aria-hidden="true" />
          {error ? <p className="rounded-xl bg-red-400/10 p-3 text-sm text-red-200">{error}</p> : null}
          <button disabled={submitting} className="flex h-12 w-full items-center justify-center gap-2 rounded-xl bg-indigo-500 font-black disabled:opacity-50">{submitting?<Loader2 className="size-4 animate-spin"/>:null}Request demo<ArrowRight className="size-4"/></button>
          <p className="text-center text-[11px] leading-5 text-slate-500">Submitting this form does not create a customer account or start billing.</p>
        </form>}
      </div>
    </section>
  </main>;
}

function Field({label,value,onChange,type="text",required=false}:{label:string;value:string;onChange:(value:string)=>void;type?:string;required?:boolean}) { return <div><label className="text-xs font-bold text-slate-300">{label}</label><input type={type} required={required} value={value} onChange={(e)=>onChange(e.target.value)} className="mt-2 h-11 w-full rounded-xl border border-white/10 bg-slate-950 px-3 text-sm outline-none focus:border-indigo-400"/></div>; }
