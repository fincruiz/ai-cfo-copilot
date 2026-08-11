"use client";

import { useState } from "react";
import {
  CheckCircle2,
  FileSpreadsheet,
  Loader2,
  UploadCloud,
  Users,
  WalletCards,
} from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { getApiErrorMessage } from "@/lib/api";
import { analyticsService } from "@/services/analytics-service";
import type { FinanceImportResult } from "@/types/analytics";

type ImportKind = "coa" | "ar" | "ap";

const cards = [
  {
    kind: "coa" as const,
    title: "Chart of accounts",
    description: "Import confirmed report mappings and reduce manual mapping work.",
    icon: FileSpreadsheet,
    columns: "Account Code, Account Name, Reporting Group, Reporting Subgroup, Statement, Sign Convention",
  },
  {
    kind: "ar" as const,
    title: "AR invoice ageing",
    description: "Analyse customer exposure, overdue invoices and collection cycles.",
    icon: Users,
    columns: "Party Name, Outstanding Amount, Document Number, Document Date, Due Date, Branch, Age Bucket",
  },
  {
    kind: "ap" as const,
    title: "AP invoice ageing",
    description: "Analyse vendor exposure, overdue liabilities and payment cycles.",
    icon: WalletCards,
    columns: "Party Name, Outstanding Amount, Document Number, Document Date, Due Date, Branch, Age Bucket",
  },
];

export default function ImportCenterPage() {
  const [files, setFiles] = useState<Record<ImportKind, File | null>>({
    coa: null,
    ar: null,
    ap: null,
  });
  const [loading, setLoading] = useState<ImportKind | null>(null);
  const [results, setResults] = useState<Partial<Record<ImportKind, FinanceImportResult>>>({});
  const [error, setError] = useState("");

  async function upload(kind: ImportKind) {
    const file = files[kind];
    if (!file) return;

    setLoading(kind);
    setError("");

    try {
      const result =
        kind === "coa"
          ? await analyticsService.uploadCoa(file)
          : kind === "ar"
            ? await analyticsService.uploadArAgeing(file)
            : await analyticsService.uploadApAgeing(file);

      setResults((current) => ({ ...current, [kind]: result }));
    } catch (uploadError) {
      setError(getApiErrorMessage(uploadError));
    } finally {
      setLoading(null);
    }
  }

  function downloadTemplate(kind: ImportKind) {
    const templates: Record<ImportKind, string> = {
      coa:
        "account_code,account_name,reporting_group,reporting_subgroup,statement,sign_convention,display_order\n4000,Sales Revenue,Revenue,Sales,Income Statement,credit,10\n5000,Cost of Goods Sold,Cost of Sales,COGS,Income Statement,debit,20\n1000,Bank Account,Current Assets,Cash and Cash Equivalents,Balance Sheet,debit,100\n",
      ar:
        "party_name,outstanding_amount,document_number,document_date,due_date,branch,age_bucket,currency_code\nCustomer A,12000,INV-001,2026-06-01,2026-07-01,MEL,1-30,AUD\n",
      ap:
        "party_name,outstanding_amount,document_number,document_date,due_date,branch,age_bucket,currency_code\nSupplier A,9000,BILL-001,2026-06-01,2026-07-01,SYD,1-30,AUD\n",
    };

    const blob = new Blob([templates[kind]], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${kind}_template.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  return (
    <div className="mx-auto max-w-7xl space-y-7">
      <div className="animate-rise">
        <p className="text-sm font-medium text-muted-foreground">Data & setup</p>
        <h1 className="mt-1 text-3xl font-semibold tracking-tight">Finance Import Centre</h1>
        <p className="mt-2 max-w-3xl text-muted-foreground">
          Upload the supporting finance packs used for mapping, customer collections,
          vendor payments and working-capital analytics.
        </p>
      </div>

      {error ? (
        <Alert variant="destructive">
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      ) : null}

      <div className="grid gap-5 lg:grid-cols-3">
        {cards.map(({ kind, title, description, icon: Icon, columns }, index) => {
          const result = results[kind];

          return (
            <Card
              key={kind}
              className="animate-card-in overflow-hidden transition duration-300 hover:-translate-y-1 hover:shadow-xl"
              style={{ animationDelay: `${index * 90}ms` }}
            >
              <CardHeader>
                <div className="flex size-12 items-center justify-center rounded-2xl bg-slate-950 text-white">
                  <Icon className="size-5" />
                </div>
                <CardTitle className="pt-3">{title}</CardTitle>
                <CardDescription className="leading-6">{description}</CardDescription>
              </CardHeader>

              <CardContent className="space-y-4">
                <div className="rounded-xl bg-muted/40 p-3 text-xs leading-5 text-muted-foreground">
                  Expected columns: {columns}
                </div>

                <Input
                  type="file"
                  accept=".csv,text/csv"
                  onChange={(event) =>
                    setFiles((current) => ({
                      ...current,
                      [kind]: event.target.files?.[0] ?? null,
                    }))
                  }
                />

                <div className="flex gap-2">
                  <Button
                    className="flex-1"
                    disabled={!files[kind] || loading !== null}
                    onClick={() => void upload(kind)}
                  >
                    {loading === kind ? (
                      <Loader2 className="size-4 animate-spin" />
                    ) : (
                      <UploadCloud className="size-4" />
                    )}
                    Upload
                  </Button>
                  <Button variant="outline" onClick={() => downloadTemplate(kind)}>
                    Template
                  </Button>
                </div>

                {result ? (
                  <div className="rounded-xl border border-emerald-200 bg-emerald-50 p-4 text-sm text-emerald-900 dark:border-emerald-900 dark:bg-emerald-950/30 dark:text-emerald-200">
                    <div className="flex items-center gap-2 font-semibold">
                      <CheckCircle2 className="size-4" />
                      Import completed
                    </div>
                    <p className="mt-2">
                      {result.inserted_rows} rows saved · {result.invalid_rows} invalid
                    </p>
                  </div>
                ) : null}
              </CardContent>
            </Card>
          );
        })}
      </div>

      <Card>
        <CardHeader>
          <CardTitle>What these imports unlock</CardTitle>
          <CardDescription>
            The legacy finance system used these packs to enrich reporting and management analysis.
          </CardDescription>
        </CardHeader>
        <CardContent className="grid gap-4 md:grid-cols-3">
          {[
            ["COA", "Persistent mappings, faster GL onboarding and fewer classification corrections."],
            ["AR ageing", "Customer exposure, overdue concentration, collection cycle and credit risk."],
            ["AP ageing", "Vendor exposure, overdue obligations, payment cycle and liquidity planning."],
          ].map(([title, text]) => (
            <div key={title} className="rounded-2xl border bg-muted/20 p-5">
              <p className="font-bold">{title}</p>
              <p className="mt-2 text-sm leading-6 text-muted-foreground">{text}</p>
            </div>
          ))}
        </CardContent>
      </Card>
    </div>
  );
}
