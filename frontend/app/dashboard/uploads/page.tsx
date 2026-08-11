"use client";

import { ChangeEvent, FormEvent, useState } from "react";
import Link from "next/link";
import { CheckCircle2, FileSpreadsheet, Loader2, UploadCloud, XCircle } from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button, buttonVariants } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { getApiErrorMessage } from "@/lib/api";
import { financeService } from "@/services/finance-service";
import type { GLUploadResult } from "@/types/finance";

export default function UploadsPage() {
  const [file, setFile] = useState<File | null>(null);
  const [sourceSystem, setSourceSystem] = useState("Manual upload");
  const [isUploading, setIsUploading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<GLUploadResult | null>(null);

  function selectFile(event: ChangeEvent<HTMLInputElement>) {
    setFile(event.target.files?.[0] ?? null);
    setResult(null);
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
    setResult(null);
    try {
      setResult(await financeService.uploadGeneralLedger(file, sourceSystem));
    } catch (uploadError) {
      setError(getApiErrorMessage(uploadError));
    } finally {
      setIsUploading(false);
    }
  }

  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div>
        <p className="text-sm font-medium text-muted-foreground">Finance data</p>
        <h1 className="mt-1 text-3xl font-semibold tracking-tight">Upload general ledger</h1>
        <p className="mt-2 text-muted-foreground">Upload a UTF-8 CSV. FinCruiz validates every row and creates ledger transactions.</p>
      </div>

      <div className="grid gap-6 lg:grid-cols-[1fr_360px]">
        <Card>
          <CardHeader>
            <CardTitle>General ledger CSV</CardTitle>
            <CardDescription>Maximum file size 10 MB. Required fields include transaction date, account code, debit and credit.</CardDescription>
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

              <div className="space-y-2">
                <Label htmlFor="source-system">Source system</Label>
                <Input id="source-system" value={sourceSystem} onChange={(event) => setSourceSystem(event.target.value)} placeholder="Xero, MYOB, SAP, manual export..." />
              </div>

              <Button type="submit" disabled={!file || isUploading} className="w-full sm:w-auto">
                {isUploading ? <Loader2 className="size-4 animate-spin" /> : <UploadCloud className="size-4" />}
                {isUploading ? "Uploading and validating..." : "Upload and validate"}
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

      {result ? (
        <Card>
          <CardHeader>
            <div className="flex items-center gap-3">
              {result.validation.invalid_rows === 0 ? <CheckCircle2 className="size-6 text-emerald-600" /> : <XCircle className="size-6 text-amber-600" />}
              <div><CardTitle>Validation complete</CardTitle><CardDescription>{result.upload.original_file_name}</CardDescription></div>
            </div>
          </CardHeader>
          <CardContent className="space-y-6">
            <div className="grid gap-4 sm:grid-cols-4">
              {[
                ["Total rows", result.validation.total_rows],
                ["Valid rows", result.validation.valid_rows],
                ["Invalid rows", result.validation.invalid_rows],
                ["Transactions inserted", result.inserted_transaction_count ?? result.validation.valid_rows],
              ].map(([label, value]) => <div key={String(label)} className="rounded-lg border p-4"><p className="text-sm text-muted-foreground">{label}</p><p className="mt-1 text-2xl font-semibold">{value}</p></div>)}
            </div>

            {result.validation.issues.length ? (
              <div className="space-y-2">
                <h3 className="font-medium">Validation issues</h3>
                {result.validation.issues.slice(0, 20).map((issue, index) => (
                  <div key={`${issue.row_number}-${index}`} className="rounded-lg border p-3 text-sm">
                    <span className="font-medium">{issue.severity.toUpperCase()}</span> · {issue.row_number ? `Row ${issue.row_number} · ` : ""}{issue.message}
                  </div>
                ))}
              </div>
            ) : null}

            <div className="flex flex-wrap gap-3">
              <Link href="/dashboard/mapping" className={buttonVariants()}>Review account mappings</Link>
              <Link href="/dashboard/reports" className={buttonVariants({ variant: "outline" })}>View trial balance</Link>
            </div>
          </CardContent>
        </Card>
      ) : null}
    </div>
  );
}
