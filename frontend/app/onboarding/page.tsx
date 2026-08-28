"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";
import { useRouter } from "next/navigation";
import {
  ArrowLeft, ArrowRight, Building2, Check, Database, FileSpreadsheet,
  Landmark, Loader2, LogOut, PlayCircle, Rocket, ShieldCheck, Sparkles,
} from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { SessionSecurityGuard } from "@/components/session-security-guard";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { getApiErrorMessage } from "@/lib/api";
import { authService } from "@/services/auth-service";
import { companyService } from "@/services/company-service";
import { MARKET_CONFIGS, marketForCountry } from "@/lib/market-config";
import {
  clearOnboardingDraft, loadOnboardingDraft, saveOnboardingDraft,
  type CommercialOnboardingDraft,
} from "@/lib/commercial-onboarding";

const industries=["Professional services","Retail / eCommerce","Manufacturing","Transport / logistics","Construction","Healthcare","Technology / SaaS","Hospitality","Wholesale / distribution","Other"];
const models=["Service business","Product / retail","Subscription","Marketplace","Project-based","Mixed model"];
const steps=["Your business","How you operate","Choose your data route","Build your workspace"];

const emptyDraft: CommercialOnboardingDraft = {
  step:0, legalName:"", tradingName:"", countryCode:"AU", currencyCode:"AUD", fye:"6",
  industry:"", businessModel:"", employees:"", revenue:"", website:"", registration:"", dataRoute:"",
};

export default function OnboardingPage(){
  const router=useRouter();
  const [draft,setDraft]=useState<CommercialOnboardingDraft>(emptyDraft);
  const [checking,setChecking]=useState(true);
  const [saving,setSaving]=useState(false);
  const [error,setError]=useState("");

  useEffect(()=>{(async()=>{
    if(!authService.hasAccessToken()){router.replace("/login");return;}
    try{await companyService.getCurrentCompany();router.replace("/dashboard");return;}catch{}
    const stored=loadOnboardingDraft();
    let next=stored ?? emptyDraft;
    try{
      const u=await authService.getCurrentUser();
      const d=(u.user_metadata?.company_details??{}) as Record<string,unknown>;
      if(!stored){
        next={...next,
          legalName:d.legal_name?String(d.legal_name):next.legalName,
          tradingName:d.trading_name?String(d.trading_name):next.tradingName,
          industry:d.industry?String(d.industry):next.industry,
          businessModel:d.business_model?String(d.business_model):next.businessModel,
        };
      }
    }catch{}
    setDraft(next); setChecking(false);
  })()},[router]);

  useEffect(()=>{if(!checking) saveOnboardingDraft(draft)},[draft,checking]);

  const step=draft.step;
  const update=(values:Partial<CommercialOnboardingDraft>)=>setDraft(v=>({...v,...values}));
  const canContinue=useMemo(()=>{
    if(step===0) return draft.legalName.trim().length>=2;
    if(step===1) return Boolean(draft.industry&&draft.businessModel);
    if(step===2) return Boolean(draft.dataRoute);
    return true;
  },[step,draft]);

  function next(){if(canContinue) update({step:Math.min(3,step+1)});}
  function back(){update({step:Math.max(0,step-1)});}

  async function submit(e:FormEvent){
    e.preventDefault();
    if(step<3){next();return;}
    setSaving(true);setError("");
    try{
      await companyService.createCompany({
        legal_name:draft.legalName.trim(), trading_name:draft.tradingName.trim()||null,
        abn:draft.registration.trim()||null, country_code:draft.countryCode.toUpperCase(),
        currency_code:draft.currencyCode.toUpperCase(), financial_year_end_month:Number(draft.fye),
        industry:draft.industry||null, business_model:draft.businessModel||null,
        employee_count:draft.employees?Number(draft.employees):null,
        annual_revenue:draft.revenue?Number(draft.revenue):null, logo_path:null,
        website_url:draft.website.trim()||null,
      });
      const route=draft.dataRoute;
      clearOnboardingDraft();
      if(route==="csv") router.replace("/dashboard/uploads?welcome=1");
      else if(route==="xero") router.replace("/dashboard/integrations?welcome=1");
      else router.replace("/dashboard?welcome=1");
    }catch(err){setError(getApiErrorMessage(err));}finally{setSaving(false)}
  }

  if(checking)return <main className="flex min-h-screen items-center justify-center"><Loader2 className="mr-2 size-5 animate-spin"/>Preparing your setup…</main>;

  const market=marketForCountry(draft.countryCode);
  return <main className="min-h-screen bg-[radial-gradient(circle_at_top_left,hsl(var(--primary)/.10),transparent_35%),linear-gradient(to_bottom,hsl(var(--background)),hsl(var(--muted)/.25))] px-5 py-8 sm:py-12">
    <SessionSecurityGuard/>
    <div className="mx-auto max-w-6xl">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2 font-semibold"><div className="flex size-10 items-center justify-center rounded-xl bg-primary text-primary-foreground"><Sparkles className="size-4"/></div><div>FinCruiz<p className="text-xs font-normal text-muted-foreground">Build your finance intelligence workspace</p></div></div>
        <Button variant="ghost" onClick={()=>{void authService.logoutEverywhere("signed-out");router.replace("/login?reason=signed-out")}}><LogOut className="size-4"/>Sign out</Button>
      </div>

      <div className="mt-8 rounded-2xl border bg-card/80 p-4 shadow-sm backdrop-blur">
        <div className="grid gap-3 md:grid-cols-4">{steps.map((label,i)=><div key={label} className={`flex items-center gap-3 rounded-xl p-3 ${i===step?"bg-primary/5 ring-1 ring-primary/20":""}`}>
          <span className={`flex size-8 shrink-0 items-center justify-center rounded-full text-sm font-bold ${i<step?"bg-emerald-100 text-emerald-800":i===step?"bg-primary text-primary-foreground":"bg-muted text-muted-foreground"}`}>{i<step?<Check className="size-4"/>:i+1}</span>
          <div><p className={`text-sm font-semibold ${i>step?"text-muted-foreground":""}`}>{label}</p><p className="text-xs text-muted-foreground">{i<step?"Complete":i===step?"In progress":"Next"}</p></div>
        </div>)}</div>
      </div>

      <div className="mt-7 grid gap-7 lg:grid-cols-[.72fr_1.28fr]">
        <aside className="space-y-5">
          <div><p className="text-sm font-semibold text-primary">Guided setup · about 5 minutes</p><h1 className="mt-2 text-4xl font-bold tracking-tight">Get to your first management insight, not another setup screen.</h1><p className="mt-4 leading-7 text-muted-foreground">FinCruiz uses your business profile and finance data to prepare the workspace. You can refine mappings, branches and assumptions after the first build.</p></div>
          <div className="rounded-2xl border bg-card p-5"><p className="font-semibold">What happens after setup?</p><div className="mt-4 space-y-3 text-sm text-muted-foreground">
            <p className="flex gap-2"><Database className="mt-0.5 size-4 shrink-0 text-primary"/>Detect accounts, history and branches.</p>
            <p className="flex gap-2"><ShieldCheck className="mt-0.5 size-4 shrink-0 text-emerald-600"/>Run financial truth and structural checks.</p>
            <p className="flex gap-2"><Sparkles className="mt-0.5 size-4 shrink-0 text-primary"/>Prepare management insights and questions you can ask.</p>
          </div></div>
          <p className="text-xs text-muted-foreground">Your progress is saved in this browser. If you leave before creating the workspace, you can resume where you stopped.</p>
        </aside>

        <Card className="overflow-hidden shadow-lg"><CardContent className="p-0">
          <div className="border-b bg-gradient-to-r from-primary/10 via-primary/5 to-transparent p-6">
            <p className="text-xs font-semibold uppercase tracking-[.18em] text-muted-foreground">Step {step+1} of 4</p>
            <h2 className="mt-2 text-2xl font-bold">{step===0?"Tell us which business we are building for":step===1?"Give FinCruiz enough context to speak your language":step===2?"How would you like to start?":"Ready to create your workspace"}</h2>
            <p className="mt-2 text-sm text-muted-foreground">{step===0?"We only require the legal/business name.":step===1?"This improves reporting language, benchmarking and AI context.":step===2?"Both real-data routes feed the same financial intelligence engine.":"We will take you directly to the right next action."}</p>
          </div>

          <form onSubmit={submit} className="p-6">
            {error?<Alert variant="destructive" className="mb-5"><AlertDescription>{error}</AlertDescription></Alert>:null}

            {step===0?<div className="grid gap-5 sm:grid-cols-2">
              <div className="sm:col-span-2"><Label>Business / legal name</Label><Input className="mt-2" value={draft.legalName} onChange={e=>update({legalName:e.target.value})} placeholder="Example Pty Ltd" required/></div>
              <div><Label>Trading name</Label><Input className="mt-2" value={draft.tradingName} onChange={e=>update({tradingName:e.target.value})} placeholder="Optional"/></div>
              <div><Label>{market.registrationLabel}</Label><Input className="mt-2" value={draft.registration} onChange={e=>update({registration:e.target.value})} placeholder="Optional"/></div>
              <div><Label>Business country</Label><select className="mt-2 w-full rounded-xl border bg-background px-3 py-2 text-sm" value={draft.countryCode} onChange={e=>{const code=e.target.value;const m=marketForCountry(code);update({countryCode:code,currencyCode:m.currencyCode,fye:String(m.defaultFyeMonth)})}}>{Object.values(MARKET_CONFIGS).map(m=><option key={m.countryCode} value={m.countryCode}>{m.countryName}</option>)}</select></div>
              <div><Label>Reporting currency</Label><Input className="mt-2" value={draft.currencyCode} readOnly/></div>
            </div>:null}

            {step===1?<div className="space-y-6">
              <div><Label>Industry</Label><div className="mt-3 grid gap-2 sm:grid-cols-2">{industries.map(x=><button type="button" key={x} onClick={()=>update({industry:x})} className={`rounded-xl border p-3 text-left text-sm transition ${draft.industry===x?"border-primary bg-primary/5 font-semibold ring-1 ring-primary/20":"hover:bg-muted/50"}`}>{x}</button>)}</div></div>
              <div><Label>How does the business mainly earn revenue?</Label><div className="mt-3 flex flex-wrap gap-2">{models.map(x=><button type="button" key={x} onClick={()=>update({businessModel:x})} className={`rounded-full border px-3 py-2 text-sm ${draft.businessModel===x?"border-primary bg-primary text-primary-foreground":"hover:bg-muted"}`}>{x}</button>)}</div></div>
              <div className="grid gap-4 sm:grid-cols-2"><div><Label>Approx. employees</Label><Input className="mt-2" type="number" min="0" value={draft.employees} onChange={e=>update({employees:e.target.value})} placeholder="Optional"/></div><div><Label>Approx. annual revenue</Label><Input className="mt-2" type="number" min="0" value={draft.revenue} onChange={e=>update({revenue:e.target.value})} placeholder={`Optional · ${draft.currencyCode}`}/></div></div>
              <div className="grid gap-4 sm:grid-cols-2"><div><Label>Website</Label><Input className="mt-2" value={draft.website} onChange={e=>update({website:e.target.value})} placeholder="Optional"/></div><div><Label>{market.financialYearLabel} ends</Label><select value={draft.fye} onChange={e=>update({fye:e.target.value})} className="mt-2 w-full rounded-xl border bg-background px-3 py-2 text-sm">{Array.from({length:12},(_,i)=><option key={i+1} value={i+1}>{new Date(2026,i,1).toLocaleString("en",{month:"long"})}</option>)}</select></div></div>
            </div>:null}

            {step===2?<div className="grid gap-4">
              {[
                ["csv",FileSpreadsheet,"Upload General Ledger","Best for the fastest real-company setup","Upload a CSV export. FinCruiz validates it in the background, detects branches and then guides account mapping."],
                ["xero",Landmark,"Connect Xero","Keep accounting data connected","Authorise Xero and continue through the same review, validation and intelligence journey."],
                ["demo",PlayCircle,"Explore with demo data","See the experience before using company data","Create the workspace and explore FinCruiz first. You can connect real data from the Integration Hub later."],
              ].map(([key,Icon,title,badge,description]:any)=><button type="button" key={key} onClick={()=>update({dataRoute:key})} className={`group rounded-2xl border p-5 text-left transition ${draft.dataRoute===key?"border-primary bg-primary/5 ring-2 ring-primary/15":"hover:-translate-y-0.5 hover:shadow-md"}`}>
                <div className="flex items-start gap-4"><div className={`flex size-11 shrink-0 items-center justify-center rounded-xl ${draft.dataRoute===key?"bg-primary text-primary-foreground":"bg-muted"}`}><Icon className="size-5"/></div><div className="flex-1"><div className="flex flex-wrap items-center gap-2"><p className="font-semibold">{title}</p><span className="rounded-full bg-muted px-2 py-1 text-[11px] text-muted-foreground">{badge}</span></div><p className="mt-2 text-sm leading-6 text-muted-foreground">{description}</p></div>{draft.dataRoute===key?<Check className="size-5 text-primary"/>:null}</div>
              </button>)}
            </div>:null}

            {step===3?<div className="space-y-5">
              <div className="rounded-2xl border bg-muted/20 p-5"><div className="flex items-center gap-3"><div className="flex size-11 items-center justify-center rounded-xl bg-emerald-100 text-emerald-700"><Check className="size-5"/></div><div><p className="font-semibold">{draft.legalName}</p><p className="text-sm text-muted-foreground">{market.countryName} · {draft.currencyCode} · {draft.industry}</p></div></div></div>
              <div className="grid gap-3 sm:grid-cols-3">{[
                [Building2,"Business context","Ready"],
                [draft.dataRoute==="xero"?Landmark:draft.dataRoute==="csv"?FileSpreadsheet:PlayCircle,"Starting route",draft.dataRoute==="xero"?"Xero":draft.dataRoute==="csv"?"GL upload":"Demo"],
                [Sparkles,"Next milestone",draft.dataRoute==="demo"?"Explore insights":"Build intelligence"],
              ].map(([Icon,title,value]:any)=><div key={title} className="rounded-xl border p-4"><Icon className="size-5 text-primary"/><p className="mt-3 text-xs text-muted-foreground">{title}</p><p className="mt-1 font-semibold">{value}</p></div>)}</div>
              <Alert><Rocket className="size-4"/><AlertDescription>{draft.dataRoute==="csv"?"After creation we will take you straight to the secure GL importer.":draft.dataRoute==="xero"?"After creation we will take you directly to the Integration Hub to connect Xero.":"After creation you will enter the dashboard and can explore before connecting company data."}</AlertDescription></Alert>
            </div>:null}

            <div className="mt-8 flex items-center justify-between border-t pt-5">
              <Button type="button" variant="ghost" disabled={step===0||saving} onClick={back}><ArrowLeft className="size-4"/>Back</Button>
              <Button type="submit" disabled={!canContinue||saving}>{saving?<Loader2 className="size-4 animate-spin"/>:step===3?<Rocket className="size-4"/>:<ArrowRight className="size-4"/>}{step===3?"Create workspace & continue":"Continue"}</Button>
            </div>
          </form>
        </CardContent></Card>
      </div>
    </div>
  </main>
}
