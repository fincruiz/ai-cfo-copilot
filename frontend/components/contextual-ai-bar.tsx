"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";
import { usePathname } from "next/navigation";
import { ArrowUp, BrainCircuit, ChevronDown, ChevronUp, Sparkles } from "lucide-react";
import { readWorkspaceScope } from "@/lib/workspace-scope";
import { usageService } from "@/services/usage-service";

const STORAGE_KEY="fincruiz_contextual_ai_collapsed";
const PAGE_CONTEXT=[
 {match:(p:string)=>p==="/dashboard",label:"this dashboard",prompts:["What should management focus on?","Why is profit moving differently from cash?"]},
 {match:(p:string)=>p.includes("working-capital"),label:"working capital",prompts:["Which customers need collections attention first?","What is driving the cash conversion cycle?"]},
 {match:(p:string)=>p.includes("native-planning")||p.includes("planning"),label:"this budget",prompts:["Which budget assumptions look unrealistic?","How should I phase this budget using historical seasonality?"]},
 {match:(p:string)=>p.includes("forecast"),label:"this forecast",prompts:["What could make us miss this forecast?","What happens to cash if revenue falls 10%?"]},
 {match:(p:string)=>p.includes("decision-simulator"),label:"this decision",prompts:["What is the biggest risk in this scenario?","Which assumption has the largest cash impact?"]},
 {match:(p:string)=>p.includes("bi")||p.includes("analytics"),label:"these trends",prompts:["What is the most important trend here?","Show me the strongest and weakest movement."]},
 {match:(p:string)=>p.includes("reports")||p.includes("kpis"),label:"these financials",prompts:["Explain these results in plain English.","Which number should management investigate first?"]},
 {match:(p:string)=>p.includes("branches"),label:"branch performance",prompts:["Which branch needs management attention?","Compare branch profitability and growth."]},
 {match:()=>true,label:"this page",prompts:["What should I pay attention to here?","What can FinCruiz help me do on this page?"]},
];

export function ContextualAIBar(){
 const pathname=usePathname(); const[question,setQuestion]=useState(""); const[collapsed,setCollapsed]=useState(true);
 const context=useMemo(()=>PAGE_CONTEXT.find(item=>item.match(pathname))!,[pathname]);
 useEffect(()=>{const saved=window.localStorage.getItem(STORAGE_KEY);setCollapsed(saved===null?true:saved==="true")},[]);
 function toggle(){setCollapsed(v=>{const next=!v;window.localStorage.setItem(STORAGE_KEY,String(next));usageService.track("contextual_ai_toggled",{state:next?"collapsed":"expanded"});return next})}
 function launch(value:string){const cleaned=value.trim();if(!cleaned)return;const scope=readWorkspaceScope();usageService.track("contextual_ai_opened",{area:pathname.split("/")[2]||"dashboard",scope:scope.mode});window.dispatchEvent(new CustomEvent("fincruiz:open-ai",{detail:{question:cleaned}}));setQuestion("")}
 function submit(e:FormEvent){e.preventDefault();launch(question)}
 return <div className="sticky bottom-0 z-20 mt-6 pb-2 pt-2 pointer-events-none">
   <div className={`pointer-events-auto ml-auto rounded-[20px] border border-indigo-200/70 bg-background/96 shadow-[0_14px_44px_rgba(15,23,42,.12)] backdrop-blur-xl transition-all dark:border-indigo-500/20 ${collapsed?"w-fit max-w-full":"mx-auto max-w-5xl p-2.5"}`}>
    {collapsed?<button type="button" onClick={toggle} className="flex items-center gap-3 rounded-[18px] px-3 py-2.5 text-left hover:bg-muted/60"><span className="flex size-9 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-600 to-sky-500 text-white"><BrainCircuit className="size-4"/></span><span><span className="block text-xs font-bold">Ask FinCruiz</span><span className="block max-w-44 truncate text-[10px] text-muted-foreground">About {context.label}</span></span><ChevronUp className="size-4 text-muted-foreground"/></button>:<>
      <div className="flex items-center justify-between gap-3 px-1 pb-2"><div className="flex items-center gap-2"><BrainCircuit className="size-4 text-indigo-600"/><span className="text-xs font-bold uppercase tracking-[.12em] text-indigo-600">Ask FinCruiz</span><span className="hidden text-xs text-muted-foreground sm:inline">· About {context.label}</span></div><button type="button" onClick={toggle} className="flex items-center gap-1 rounded-lg px-2 py-1 text-xs text-muted-foreground hover:bg-muted"><ChevronDown className="size-4"/>Collapse</button></div>
      <form onSubmit={submit} className="flex items-center gap-2"><input value={question} onChange={e=>setQuestion(e.target.value)} placeholder={`Ask FinCruiz about ${context.label}…`} className="min-w-0 flex-1 rounded-xl bg-muted/45 px-3 py-2.5 text-sm outline-none placeholder:text-muted-foreground"/><button type="submit" disabled={!question.trim()} className="flex size-10 shrink-0 items-center justify-center rounded-xl bg-primary text-primary-foreground disabled:opacity-35"><ArrowUp className="size-4"/></button></form>
      <div className="mt-2 hidden gap-1.5 overflow-x-auto pb-0.5 sm:flex">{context.prompts.map(prompt=><button key={prompt} type="button" onClick={()=>launch(prompt)} className="whitespace-nowrap rounded-full bg-muted/70 px-3 py-1 text-[10px] font-medium text-muted-foreground hover:bg-indigo-50 hover:text-indigo-700"><Sparkles className="mr-1 inline size-2.5"/>{prompt}</button>)}</div>
    </>}
   </div>
 </div>
}
