"use client";
import { useEffect, useState } from "react";
import { Save, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { companyService } from "@/services/company-service";
export default function SettingsPage(){
 const [s,setS]=useState<any>(null);const [saving,setSaving]=useState(false);
 useEffect(()=>{companyService.getPreferences().then(setS)},[]);
 if(!s)return <div className="flex min-h-[400px] items-center justify-center"><Loader2 className="size-5 animate-spin"/></div>;
 async function save(){setSaving(true);try{setS(await companyService.updatePreferences(s))}finally{setSaving(false)}}
 return <div className="mx-auto max-w-4xl space-y-6"><div><p className="text-sm text-muted-foreground">Administration</p><h1 className="text-3xl font-semibold">Workspace Settings</h1></div><Card><CardHeader><CardTitle>Reporting preferences</CardTitle><CardDescription>Control defaults across reports and analytics.</CardDescription></CardHeader><CardContent className="space-y-5">
 {[
 ["theme_preference","Theme",["system","light","dark"]],
 ["reporting_frequency","Reporting frequency",["monthly","quarterly","annual"]],
 ["default_report_view","Default report view",["consolidated","branch"]],
 ["number_format","Number format",["international","indian"]],
 ].map(([key,label,options]:any)=><div key={key} className="grid gap-2 sm:grid-cols-[220px_1fr] sm:items-center"><label>{label}</label><select className="h-10 rounded-md border bg-background px-3" value={s[key]} onChange={e=>setS({...s,[key]:e.target.value})}>{options.map((o:string)=><option key={o}>{o}</option>)}</select></div>)}
 <div className="grid gap-2 sm:grid-cols-[220px_1fr] sm:items-center"><label>Variance warning %</label><Input type="number" value={s.variance_warning_percent} onChange={e=>setS({...s,variance_warning_percent:Number(e.target.value)})}/></div>
 {["show_ai_assistant","email_notifications"].map(key=><label key={key} className="flex items-center gap-3"><input type="checkbox" checked={Boolean(s[key])} onChange={e=>setS({...s,[key]:e.target.checked})}/>{key==="show_ai_assistant"?"Show AI CFO assistant":"Email notifications"}</label>)}
 <Button onClick={()=>void save()} disabled={saving}>{saving?<Loader2 className="size-4 animate-spin"/>:<Save className="size-4"/>}Save settings</Button>
 </CardContent></Card></div>
}
