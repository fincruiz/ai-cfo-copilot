"use client";
import { useEffect, useState } from "react";
import { Save, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { companyService } from "@/services/company-service";

export default function ProfilePage(){
 const [form,setForm]=useState<any>(null); const [saving,setSaving]=useState(false); const [message,setMessage]=useState("");
 useEffect(()=>{companyService.getCurrentCompany().then(setForm)},[]);
 if(!form)return <div className="flex min-h-[400px] items-center justify-center"><Loader2 className="size-5 animate-spin"/></div>;
 async function save(){setSaving(true);try{setForm(await companyService.updateCompany(form));setMessage("Company profile updated.")}finally{setSaving(false)}}
 return <div className="mx-auto max-w-4xl space-y-6"><div><p className="text-sm text-muted-foreground">Administration</p><h1 className="text-3xl font-semibold">Company Profile</h1></div>{message?<p className="rounded-xl bg-emerald-50 p-3 text-emerald-800">{message}</p>:null}<Card><CardHeader><CardTitle>Business details</CardTitle><CardDescription>These details appear in reports and future board packs.</CardDescription></CardHeader><CardContent className="grid gap-5 md:grid-cols-2">{[["legal_name","Legal name"],["trading_name","Trading name"],["abn","Tax ID / ABN"],["industry","Industry"],["business_model","Business model"],["website_url","Website"],["currency_code","Currency"],["country_code","Country code"]].map(([k,l])=><div key={k} className="space-y-2"><Label>{l}</Label><Input value={form[k]??""} onChange={e=>setForm({...form,[k]:e.target.value})}/></div>)}<div className="md:col-span-2"><Button onClick={()=>void save()} disabled={saving}>{saving?<Loader2 className="size-4 animate-spin"/>:<Save className="size-4"/>}Save profile</Button></div></CardContent></Card></div>
}
