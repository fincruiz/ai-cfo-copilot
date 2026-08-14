"use client";

import Link from "next/link";
import { useEffect, useMemo, useState } from "react";
import {
  ArrowLeft, ArrowRight, BadgeCheck, Bot, BrainCircuit, BriefcaseBusiness, CircleDollarSign,
  Cloud, Gauge, LineChart, LockKeyhole, Pause, Play, RotateCcw, ShieldCheck, Sparkles,
  Target, TrendingUp, UploadCloud, WandSparkles, WalletCards,
} from "lucide-react";

type Scene = { kicker:string; title:string; body:string; question:string; answer:string; action:string; accent:string };
const scenes: Scene[] = [
  { kicker:"1 · Raw data", title:"You start with a spreadsheet, not a finance degree.", body:"Drop in the General Ledger you already export from your accounting system. FinCruiz checks the structure before anything becomes a report.", question:"Is my file good enough?", answer:"I found the date, account, debit and credit fields. Debits equal credits and 98% of accounts can be suggested automatically.", action:"Review 4 account mappings", accent:"from-sky-500/30 to-indigo-500/10" },
  { kicker:"2 · Understand", title:"Numbers turn into a story you can actually use.", body:"Instead of asking you to interpret a trial balance, FinCruiz translates movements into plain-English management signals.", question:"What changed this month?", answer:"Revenue grew 9.4%, but receivables grew faster. Profit improved, while cash conversion weakened. Collections deserve attention before increasing spend.", action:"Open the collection-risk view", accent:"from-indigo-500/30 to-violet-500/10" },
  { kicker:"3 · Decide", title:"Ask the question you would normally ask a CFO.", body:"The AI CFO uses calculated company metrics first, then adds live industry and economic context when it is useful and clearly sourced.", question:"What should I do in the next 30 days?", answer:"Protect gross margin, focus on the five largest overdue customers, and test a slower-collections downside case before approving discretionary spend.", action:"Run the 90-day downside scenario", accent:"from-violet-500/30 to-fuchsia-500/10" },
  { kicker:"4 · Stay in control", title:"Your data remains your decision.", body:"Reset only the module you want, reset the whole workspace, or permanently delete the profile. Destructive actions always ask for confirmation.", question:"What if I upload the wrong file?", answer:"Reset only that module. Your other finance data stays intact. I can also take you directly to the right page to replace it.", action:"Reset AR only · keep everything else", accent:"from-emerald-500/25 to-cyan-500/10" },
];

const revenue=[72,76,73,81,86,89,96,101,105,112,118,126];
const cash=[66,64,69,67,62,65,70,68,74,78,75,84];
const months=["Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug"];

export default function DemoPage(){
  const [scene,setScene]=useState(0); const [playing,setPlaying]=useState(true); const [question,setQuestion]=useState("What should management focus on?");
  const [scenario,setScenario]=useState<"base"|"upside"|"downside">("base");
  useEffect(()=>{ if(!playing) return; const id=window.setInterval(()=>setScene(v=>(v+1)%scenes.length),6500); return()=>window.clearInterval(id);},[playing]);
  const s=scenes[scene];
  const scenarioData=useMemo(()=>scenario==="upside"?{sales:"+14.2%",cash:"$2.9m",profit:"$2.16m",note:"Faster growth + 2pt margin improvement"}:scenario==="downside"?{sales:"-4.8%",cash:"$1.28m",profit:"$1.17m",note:"Sales softness + collections 12 days slower"}:{sales:"+8.9%",cash:"$2.21m",profit:"$1.82m",note:"Current run-rate and collection assumptions"},[scenario]);
  const answers:Record<string,string>={
    "What should management focus on?":"Collections are the clearest near-term lever. Sales are rising, but overdue receivables are absorbing more cash. Start with the largest five overdue customers.",
    "Can I understand this without finance knowledge?":"Yes. FinCruiz keeps the accounting checks underneath, but translates the result into plain language: what changed, why it matters, and what action to consider next.",
    "What happens if the economy slows?":"Use the downside scenario. FinCruiz can combine your current run-rate with sourced economic context, then show how the assumption changes cash, margin and profit.",
  };
  return <main className="min-h-screen overflow-hidden bg-[#07101f] text-white">
    <div className="fixed inset-0 pointer-events-none demo-aurora"/>
    <header className="relative z-30 mx-auto flex max-w-7xl items-center justify-between px-5 py-5 lg:px-8">
      <Link href="/" className="flex items-center gap-2 text-sm text-slate-300 hover:text-white"><ArrowLeft className="size-4"/>FinCruiz</Link>
      <div className="flex items-center gap-2"><span className="hidden rounded-full border border-emerald-300/20 bg-emerald-300/10 px-3 py-1.5 text-xs text-emerald-200 sm:inline-flex"><ShieldCheck className="mr-1.5 size-3.5"/>Synthetic data · no login</span><Link href="/signup" className="rounded-xl bg-white px-4 py-2.5 text-sm font-bold text-slate-950">Start with my business</Link></div>
    </header>

    <section className="relative z-10 mx-auto max-w-7xl px-5 pb-16 pt-8 lg:px-8">
      <div className="grid items-center gap-10 lg:grid-cols-[.85fr_1.15fr]">
        <div>
          <div className="inline-flex items-center gap-2 rounded-full border border-sky-300/15 bg-sky-300/10 px-3 py-1.5 text-xs font-semibold text-sky-100"><Sparkles className="size-3.5"/>2-minute interactive product story</div>
          <h1 className="mt-5 text-4xl font-black tracking-[-.045em] sm:text-6xl">Finance that talks like a <span className="bg-gradient-to-r from-sky-300 via-indigo-300 to-violet-300 bg-clip-text text-transparent">business partner.</span></h1>
          <p className="mt-5 max-w-xl text-base leading-8 text-slate-300 sm:text-lg">You do not need to know finance terminology. Watch a sample company move from spreadsheet → insight → decision, then click around the live simulation yourself.</p>
          <div className="mt-7 flex flex-wrap gap-3"><button onClick={()=>setPlaying(v=>!v)} className="inline-flex items-center gap-2 rounded-xl bg-indigo-500 px-5 py-3 text-sm font-bold hover:bg-indigo-400">{playing?<Pause className="size-4"/>:<Play className="size-4"/>}{playing?"Pause story":"Play story"}</button><button onClick={()=>{setScene(0);setPlaying(true)}} className="inline-flex items-center gap-2 rounded-xl border border-white/15 bg-white/5 px-5 py-3 text-sm font-semibold hover:bg-white/10"><RotateCcw className="size-4"/>Restart</button></div>
          <div className="mt-8 grid grid-cols-4 gap-2">{scenes.map((item,index)=><button key={item.kicker} onClick={()=>{setScene(index);setPlaying(false)}} className={`h-1.5 rounded-full ${index===scene?"bg-indigo-300":"bg-white/10"}`} aria-label={`Show scene ${index+1}`}/>)}</div>
        </div>

        <div className={`relative overflow-hidden rounded-[34px] border border-white/10 bg-gradient-to-br ${s.accent} p-1 shadow-[0_40px_120px_rgba(20,30,80,.45)]`}>
          <div className="relative min-h-[520px] overflow-hidden rounded-[31px] bg-slate-950/80 p-5 backdrop-blur sm:p-7">
            <div className="absolute -right-12 -top-12 size-44 rounded-full bg-indigo-400/10 blur-3xl"/><div className="flex items-center justify-between"><span className="text-xs font-bold uppercase tracking-[.18em] text-indigo-200">{s.kicker}</span><span className="flex items-center gap-2 text-xs text-slate-400"><span className="size-2 animate-pulse rounded-full bg-emerald-400"/>Live simulation</span></div>
            <div key={scene} className="mt-8 animate-scene-in">
              <h2 className="max-w-2xl text-3xl font-bold tracking-tight sm:text-4xl">{s.title}</h2><p className="mt-4 max-w-2xl leading-7 text-slate-300">{s.body}</p>
              <div className="mt-8 grid gap-4 sm:grid-cols-[.8fr_1.2fr]">
                <div className="rounded-2xl border border-white/10 bg-white/[.04] p-5"><p className="text-xs uppercase tracking-wider text-slate-500">You ask</p><p className="mt-3 text-lg font-semibold">“{s.question}”</p><div className="mt-6 flex gap-2"><span className="demo-dot"/><span className="demo-dot [animation-delay:180ms]"/><span className="demo-dot [animation-delay:360ms]"/></div></div>
                <div className="rounded-2xl border border-indigo-300/15 bg-indigo-300/[.07] p-5"><div className="flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-indigo-200"><Bot className="size-4"/>AI CFO</div><p className="mt-3 text-sm leading-7 text-slate-100">{s.answer}</p><div className="mt-4 rounded-xl bg-white/[.05] px-3 py-2 text-xs text-sky-200"><b>Suggested next step:</b> {s.action}</div></div>
              </div>
              <div className="mt-5 flex items-center gap-3 rounded-2xl border border-emerald-300/10 bg-emerald-300/[.05] p-4 text-sm text-slate-300"><BadgeCheck className="size-5 shrink-0 text-emerald-300"/><span><b className="text-white">Evidence before explanation.</b> Calculated finance checks are separated from AI narrative and external context.</span></div>
            </div>
          </div>
        </div>
      </div>

      <div className="mt-20 text-center"><p className="text-xs font-semibold uppercase tracking-[.2em] text-indigo-300">Now take control</p><h2 className="mt-3 text-3xl font-bold sm:text-4xl">Try the parts customers use every week.</h2></div>
      <div className="mt-8 grid gap-5 lg:grid-cols-[1.2fr_.8fr]">
        <div className="rounded-[30px] border border-white/10 bg-white/[.05] p-5 sm:p-6">
          <div className="flex flex-wrap items-center justify-between gap-4"><div><p className="text-sm font-semibold">Demo Company · Executive view</p><p className="mt-1 text-xs text-slate-400">Synthetic 12-month history</p></div><span className="rounded-full bg-emerald-400/10 px-3 py-1 text-xs text-emerald-200">Finance confidence 96/100</span></div>
          <div className="mt-5 grid grid-cols-2 gap-3 sm:grid-cols-4">{[["Revenue","$14.0m",TrendingUp],["Profit","$1.82m",CircleDollarSign],["Cash","$2.21m",WalletCards],["Confidence","96%",Gauge]].map(([label,value,Icon]:any)=><div key={label} className="rounded-2xl border border-white/10 bg-slate-950/50 p-4"><Icon className="size-4 text-indigo-300"/><p className="mt-4 text-2xl font-bold">{value}</p><p className="mt-1 text-xs text-slate-400">{label}</p></div>)}</div>
          <div className="mt-5 rounded-2xl border border-white/10 bg-slate-950/45 p-5"><div className="flex justify-between"><span className="text-sm font-semibold">Business momentum</span><span className="text-xs text-slate-500">Revenue · cash</span></div><div className="mt-6 flex h-44 items-end gap-2">{revenue.map((v,i)=><div key={months[i]} className="flex flex-1 flex-col items-center gap-2"><div className="relative flex h-32 w-full items-end justify-center"><div className="w-[68%] rounded-t-md bg-indigo-400/65" style={{height:`${v/126*100}%`}}/><div className="absolute bottom-0 w-[24%] rounded-t bg-emerald-300" style={{height:`${cash[i]/90*80}%`}}/></div><span className="text-[9px] text-slate-500">{months[i]}</span></div>)}</div></div>
        </div>
        <div className="rounded-[30px] border border-indigo-300/15 bg-indigo-300/[.06] p-6"><div className="flex items-center gap-3"><div className="flex size-11 items-center justify-center rounded-2xl bg-indigo-400/15"><BrainCircuit className="size-5 text-indigo-200"/></div><div><p className="font-bold">Ask without finance jargon</p><p className="text-xs text-slate-400">Choose a question</p></div></div><div className="mt-5 space-y-2">{Object.keys(answers).map(q=><button key={q} onClick={()=>setQuestion(q)} className={`w-full rounded-xl border px-4 py-3 text-left text-sm ${q===question?"border-indigo-300/35 bg-indigo-300/10":"border-white/10 bg-white/[.03] hover:bg-white/[.07]"}`}>{q}</button>)}</div><div className="mt-4 rounded-2xl bg-slate-950/60 p-4 text-sm leading-7 text-slate-200"><Sparkles className="mb-3 size-4 text-indigo-300"/>{answers[question]}</div></div>
      </div>

      <div className="mt-5 rounded-[30px] border border-white/10 bg-white/[.04] p-6"><div className="flex flex-wrap items-end justify-between gap-4"><div><p className="font-bold">Scenario simulator</p><p className="mt-1 text-sm text-slate-400">See how a business assumption changes the outlook.</p></div><div className="flex gap-2">{(["downside","base","upside"] as const).map(x=><button key={x} onClick={()=>setScenario(x)} className={`rounded-xl px-4 py-2 text-xs font-bold capitalize ${scenario===x?"bg-white text-slate-950":"border border-white/10 bg-white/5 text-slate-300"}`}>{x}</button>)}</div></div><div className="mt-6 grid gap-3 sm:grid-cols-4"><div className="rounded-xl bg-slate-950/45 p-4"><p className="text-xs text-slate-500">Sales growth</p><p className="mt-2 text-2xl font-bold">{scenarioData.sales}</p></div><div className="rounded-xl bg-slate-950/45 p-4"><p className="text-xs text-slate-500">Forecast cash</p><p className="mt-2 text-2xl font-bold">{scenarioData.cash}</p></div><div className="rounded-xl bg-slate-950/45 p-4"><p className="text-xs text-slate-500">Forecast profit</p><p className="mt-2 text-2xl font-bold">{scenarioData.profit}</p></div><div className="rounded-xl bg-slate-950/45 p-4"><p className="text-xs text-slate-500">Assumption</p><p className="mt-2 text-sm leading-6">{scenarioData.note}</p></div></div></div>

      <div className="mt-16 grid gap-4 md:grid-cols-3">{[
        [UploadCloud,"Bring what you already have","Start with a CSV export. The guided uploader checks the file and tells you what is missing."],
        [WandSparkles,"Get the next step","The product and AI guide explain what to do next instead of leaving you inside a complex finance menu."],
        [LockKeyhole,"Stay in control","Reset one module, reset everything, or delete the account. Every destructive action uses a clear confirmation."],
      ].map(([Icon,title,text]:any)=><div key={title} className="rounded-2xl border border-white/10 bg-white/[.04] p-5"><Icon className="size-5 text-indigo-300"/><p className="mt-4 font-bold">{title}</p><p className="mt-2 text-sm leading-6 text-slate-400">{text}</p></div>)}</div>

      <div className="mt-16 rounded-[34px] border border-indigo-300/15 bg-gradient-to-r from-indigo-500/15 to-sky-500/10 p-8 text-center sm:p-12"><Target className="mx-auto size-7 text-indigo-300"/><h2 className="mt-4 text-3xl font-bold">Seen enough? Use the same experience with your business.</h2><p className="mx-auto mt-3 max-w-2xl text-slate-300">Create a workspace, load demo data first if you prefer, and replace it only when you are comfortable.</p><div className="mt-6 flex justify-center gap-3"><Link href="/signup" className="inline-flex items-center gap-2 rounded-xl bg-white px-5 py-3 text-sm font-bold text-slate-950">Create free workspace<ArrowRight className="size-4"/></Link><Link href="/login" className="rounded-xl border border-white/15 px-5 py-3 text-sm font-semibold">Sign in</Link></div></div>
    </section>
  </main>
}
