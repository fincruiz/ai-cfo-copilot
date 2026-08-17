"use client";

import { ChangeEvent, FormEvent, useEffect, useState } from "react";
import Link from "next/link";
import { CheckCircle2, Clock3, Loader2, RefreshCcw, UploadCloud, XCircle } from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button, buttonVariants } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { getApiErrorMessage } from "@/lib/api";
import { financeService } from "@/services/finance-service";
import { ModuleResetButton } from "@/components/module-reset-button";
import type { IngestionJob } from "@/types/finance";


const fieldAliases: Record<string, string[]> = {
  date: ["date", "transaction date", "transaction_date", "posting date", "journal date"],
  account: ["account code", "account_code", "gl code", "ledger code", "nominal code", "account number"],
  accountName: ["account name", "account_name", "ledger name", "description"],
  debit: ["debit", "debit amount", "dr"],
  credit: ["credit", "credit amount", "cr"],
};

function detectHeader(headers: string[]) {
  const normalized = headers.map((h) => h.trim().toLowerCase());
  return Object.fromEntries(Object.entries(fieldAliases).map(([key, aliases]) => [key, aliases.some((alias) => normalized.includes(alias))]));
}

export default function UploadsPage() {
  const [file, setFile] = useState<File | null>(null);
  const [sourceSystem, setSourceSystem] = useState("Manual upload");
  const [isUploading, setIsUploading] = useState(false);
  const [error, setError] = useState("");
  const [job, setJob] = useState<IngestionJob | null>(null);
  const [jobs, setJobs] = useState<IngestionJob[]>([]);
  const [preview, setPreview] = useState<string[][]>([]);
  const [welcomeFlow, setWelcomeFlow] = useState(false);
  const detected = detectHeader(preview[0] ?? []);
  const detectedCount = Object.values(detected).filter(Boolean).length;

  useEffect(() => { setWelcomeFlow(new URLSearchParams(window.location.search).get("welcome") === "1"); }, []);

  useEffect(() => {
    let cancelled = false;
    async function refreshJobs() {
      try { const items = await financeService.getIngestionJobs(); if (!cancelled) { setJobs(items); if (job) { const next = items.find((item) => item.id === job.id); if (next) setJob(next); } } } catch {}
    }
    void refreshJobs();
    const timer = window.setInterval(refreshJobs, 2000);
    return () => { cancelled = true; window.clearInterval(timer); };
  }, [job?.id]);

  async function selectFile(event: ChangeEvent<HTMLInputElement>) {
    const selected = event.target.files?.[0] ?? null;
    setFile(selected);
    if (selected) { const text = await selected.slice(0, 64 * 1024).text(); setPreview(text.split(/\r?\n/).filter(Boolean).slice(0,4).map(line => line.split(",").slice(0,8))); } else setPreview([]);
    setJob(null);
    setError("");
  }

  async function submit(event: FormEvent) {
    event.preventDefault();
    if (!file) {
      setError("Choose a CSV file first.");
      return;
    }

    setIsUploading(true);
    setError("");
    setJob(null);
    try {
      const created = await financeService.stageGeneralLedger(file, sourceSystem);
      setJob(created);
      setJobs((current) => [created, ...current.filter((item) => item.id !== created.id)]);
    } catch (uploadError) {
      setError(getApiErrorMessage(uploadError));
    } finally {
      setIsUploading(false);
    }
  }

  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div className="flex items-start justify-between gap-4">
        <div>
        <p className="text-sm font-medium text-muted-foreground">Finance data</p>
        <h1 className="mt-1 text-3xl font-semibold tracking-tight">Upload general ledger</h1>
        <p className="mt-2 text-muted-foreground">Upload a UTF-8 CSV. FinCruiz validates every row and creates ledger transactions.</p>
        </div>
        <ModuleResetButton scope="general_ledger" label="Reset GL data" description="This removes only the loaded General Ledger and its upload records. Your company profile, AR/AP data and settings remain." />
      </div>

      <div className="grid gap-6 lg:grid-cols-[1fr_360px]">
        <Card>
          <CardHeader>
            <CardTitle>Step 1 · Choose your General Ledger</CardTitle>
            <CardDescription>Large CSVs are streamed to staging storage first, then validated and imported in the background. Required fields include transaction date, account code, debit and credit.</CardDescription>
          </CardHeader>
          <CardContent>
            <form onSubmit={submit} className="space-y-5">
              {error ? <Alert variant="destructive"><AlertDescription>{error}</AlertDescription></Alert> : null}

              <Label htmlFor="gl-file" className="flex min-h-48 cursor-pointer flex-col items-center justify-center rounded-xl border border-dashed bg-muted/20 p-8 text-center hover:bg-muted/35">
                <UploadCloud className="mb-4 size-10 text-muted-foreground" />
                <span className="font-medium">{file ? file.name : "Choose a CSV file"}</span>
                <span className="mt-1 text-sm text-muted-foreground">Click to browse</span>
                <Input id="gl-file" type="file" accept=".csv,text/csv" className="sr-only" onChange={selectFile} />
              </Label>

              {preview.length ? <div className="space-y-2 rounded-xl border bg-muted/20 p-4"><div className="flex items-center justify-between"><p className="font-medium">Smart file preview</p><span className="text-xs text-muted-foreground">Step 2 of 3 · confirm detected columns</span></div><div className="overflow-x-auto"><table className="w-full text-xs"><tbody>{preview.map((row,i)=><tr key={i} className={i===0?"font-semibold":"text-muted-foreground"}>{row.map((cell,j)=><td key={j} className="border-b px-2 py-2 whitespace-nowrap">{cell || "—"}</td>)}</tr>)}</tbody></table></div><p className="text-xs text-muted-foreground">FinCruiz will run full validation on upload and show any rows that need attention before you continue to mapping.</p></div>:null}
              {preview.length ? <div className="rounded-xl border p-4"><div className="flex items-center justify-between gap-3"><div><p className="font-medium">Detected finance fields</p><p className="text-xs text-muted-foreground">A quick client-side check before the full backend validation.</p></div><span className={`rounded-full px-3 py-1 text-xs font-semibold ${detectedCount >= 4 ? "bg-emerald-100 text-emerald-800" : "bg-amber-100 text-amber-800"}`}>{detectedCount}/5 detected</span></div><div className="mt-4 grid gap-2 sm:grid-cols-5">{[["Date",detected.date],["Account",detected.account],["Account name",detected.accountName],["Debit",detected.debit],["Credit",detected.credit]].map(([label,ok])=><div key={String(label)} className={`rounded-lg border px-3 py-2 text-xs ${ok ? "bg-emerald-50 text-emerald-800" : "bg-amber-50 text-amber-800"}`}>{ok ? "✓" : "!"} {label}</div>)}</div>{detectedCount < 4 ? <p className="mt-3 text-xs text-amber-700">Some core fields were not recognised. You can still upload; the backend will perform the authoritative validation and explain what needs fixing.</p> : <p className="mt-3 text-xs text-emerald-700">The file looks structurally ready for full validation.</p>}</div> : null}

              <div className="space-y-2">
                <Label htmlFor="source-system">Source system</Label>
                <Input id="source-system" value={sourceSystem} onChange={(event) => setSourceSystem(event.target.value)} placeholder="Xero, MYOB, SAP, manual export..." />
              </div>

              <Button type="submit" disabled={!file || isUploading} className="w-full sm:w-auto">
                {isUploading ? <Loader2 className="size-4 animate-spin" /> : <UploadCloud className="size-4" />}
                {isUploading ? "Streaming to secure staging..." : "Step 3 · Stage and process"}
              </Button>
            </form>
          </CardContent>
        </Card>

        <Card>
          <CardHeader><CardTitle>Expected columns</CardTitle><CardDescription>FinCruiz also recognises common aliases.</CardDescription></CardHeader>
          <CardContent className="space-y-3 text-sm">
            {["transaction_date", "source_account_code", "source_account_name", "description", "debit", "credit", "currency_code"].map((column) => (
              <div key={column} className="flex items-center gap-2"><CheckCircle2 className="size-4 text-emerald-600" /><code>{column}</code></div>
            ))}
          </CardContent>
        </Card>
      </div>

      {job ? (
        <Card>
          <CardHeader><CardTitle className="flex items-center gap-2">{job.status === "completed" ? <CheckCircle2 className="size-5 text-emerald-600"/> : job.status === "failed" || job.status === "validation_failed" ? <XCircle className="size-5 text-destructive"/> : <Clock3 className="size-5 text-primary"/>}Background import</CardTitle><CardDescription>{job.original_file_name} · {job.phase.replaceAll("_", " ")}</CardDescription></CardHeader>
          <CardContent className="space-y-4">
            <div className="h-2 overflow-hidden rounded-full bg-muted"><div className="h-full bg-primary transition-all" style={{width:`${job.progress_percent}%`}}/></div>
            <div className="grid gap-3 sm:grid-cols-4">{[["Progress",`${job.progress_percent}%`],["Rows",job.total_rows?.toLocaleString() ?? "—"],["Inserted",job.inserted_rows.toLocaleString()],["Status",job.status.replaceAll("_"," ")]].map(([label,value])=><div key={String(label)} className="rounded-xl border p-3"><p className="text-xs text-muted-foreground">{label}</p><p className="mt-1 font-semibold capitalize">{value}</p></div>)}</div>
            {job.error_message ? <Alert variant="destructive"><AlertDescription>{job.error_message}</AlertDescription></Alert> : null}
            {job.status === "completed" ? <div className="flex flex-wrap gap-3">{welcomeFlow ? <Link href="/dashboard/getting-started" className={buttonVariants()}>Continue guided setup</Link> : <Link href="/dashboard/mapping" className={buttonVariants()}>Review account mappings</Link>}<Link href="/dashboard/reports" className={buttonVariants({variant:"outline"})}>View reports</Link></div> : null}
            {job.status === "failed" ? <Button variant="outline" onClick={async()=>setJob(await financeService.retryIngestionJob(job.id))}><RefreshCcw className="size-4"/>Retry job</Button> : null}
          </CardContent>
        </Card>
      ) : null}

      {jobs.length ? <Card><CardHeader><CardTitle>Recent imports</CardTitle><CardDescription>Uploads continue processing even if you navigate elsewhere in FinCruiz.</CardDescription></CardHeader><CardContent className="space-y-2">{jobs.slice(0,8).map((item)=><button type="button" key={item.id} onClick={()=>setJob(item)} className="flex w-full items-center justify-between rounded-xl border p-3 text-left hover:bg-muted/30"><div><p className="font-medium">{item.original_file_name}</p><p className="text-xs text-muted-foreground">{item.phase.replaceAll("_"," ")} · {new Date(item.created_at).toLocaleString()}</p></div><span className="text-sm font-semibold">{item.progress_percent}%</span></button>)}</CardContent></Card> : null}
    </div>
  );
}
