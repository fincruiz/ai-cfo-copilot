"use client";

import { ChangeEvent, FormEvent, useMemo, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import {
  ArrowLeft,
  ArrowRight,
  BarChart3,
  Building2,
  Check,
  CheckCircle2,
  CloudCog,
  FileSpreadsheet,
  ImagePlus,
  Layers3,
  Loader2,
  Sparkles,
} from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { getApiErrorMessage } from "@/lib/api";
import { authService } from "@/services/auth-service";

const industries = [
  "Manufacturing",
  "Wholesale / Distribution",
  "Retail",
  "Professional Services",
  "Construction",
  "Logistics",
  "Hospitality",
  "Healthcare",
  "Technology",
  "Other",
];

const countries = [
  ["AU", "Australia", "AUD"],
  ["IN", "India", "INR"],
  ["US", "United States", "USD"],
  ["GB", "United Kingdom", "GBP"],
  ["CA", "Canada", "CAD"],
  ["NZ", "New Zealand", "NZD"],
  ["SG", "Singapore", "SGD"],
  ["AE", "United Arab Emirates", "AED"],
  ["OT", "Other", "USD"],
];

const currencies = [
  ["AUD", "Australian Dollar"],
  ["INR", "Indian Rupee"],
  ["USD", "US Dollar"],
  ["GBP", "British Pound"],
  ["CAD", "Canadian Dollar"],
  ["NZD", "New Zealand Dollar"],
  ["SGD", "Singapore Dollar"],
  ["AED", "UAE Dirham"],
  ["EUR", "Euro"],
  ["JPY", "Japanese Yen"],
];

const modules = [
  "Financial Statements",
  "KPI Dashboard",
  "AI CFO",
  "Forecasting",
  "Working Capital",
  "Benchmarking",
  "Board Reports",
  "Board Pack",
  "PowerPoint Export",
];

const reportingStructures = [
  {
    value: "Consolidated Only",
    title: "Consolidated only",
    text: "One company-level view without branch segmentation.",
    icon: Building2,
  },
  {
    value: "Branch / Business Unit Reporting",
    title: "Branch or business unit",
    text: "Operational reporting focused on individual branches.",
    icon: Layers3,
  },
  {
    value: "Consolidated + Branch Reporting",
    title: "Both consolidated and branch",
    text: "Company-wide consolidation plus branch-level analysis.",
    icon: Sparkles,
  },
];

export default function SignupPage() {
  const router = useRouter();
  const [step, setStep] = useState(1);
  const [done, setDone] = useState(false);
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const [logoPreview, setLogoPreview] = useState("");

  const [form, setForm] = useState({
    full_name: "",
    email: "",
    password: "",
    confirm_password: "",
    legal_name: "",
    trading_name: "",
    tax_id: "",
    industry: "Technology",
    country_code: "AU",
    currency_code: "AUD",
    financial_year_end_month: "6",
    business_model: "",
    employee_count: "",
    annual_revenue: "",
    website_url: "",
    reporting_frequency: "Monthly",
    reporting_structure: "Consolidated + Branch Reporting",
    preferred_data_source: "Manual Excel / CSV",
    enabled_modules: modules,
  });

  function set<K extends keyof typeof form>(key: K, value: (typeof form)[K]) {
    setForm((current) => ({ ...current, [key]: value }));
  }

  const countryName = useMemo(
    () => countries.find(([code]) => code === form.country_code)?.[1] ?? form.country_code,
    [form.country_code],
  );

  function handleLogo(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!["image/png", "image/jpeg", "image/webp"].includes(file.type)) {
      setError("Logo must be PNG, JPG or WebP.");
      return;
    }

    if (file.size > 2 * 1024 * 1024) {
      setError("Logo must be smaller than 2 MB.");
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result ?? "");
      setLogoPreview(result);
      window.localStorage.setItem("fincruiz_pending_logo", result);
      window.localStorage.setItem("fincruiz_pending_logo_name", file.name);
    };
    reader.readAsDataURL(file);
  }

  function next(event: FormEvent) {
    event.preventDefault();
    setError("");

    if (step === 1 && form.password !== form.confirm_password) {
      setError("Passwords do not match.");
      return;
    }

    if (step === 2 && !form.legal_name.trim()) {
      setError("Company legal name is required.");
      return;
    }

    setStep((value) => Math.min(4, value + 1));
  }

  async function submit() {
    setSaving(true);
    setError("");

    try {
      const result = await authService.signup({
        email: form.email.trim(),
        password: form.password,
        full_name: form.full_name.trim(),
        company_details: {
          legal_name: form.legal_name.trim(),
          trading_name: form.trading_name.trim() || null,
          abn: form.tax_id.trim() || null,
          country_code: form.country_code,
          country_name: countryName,
          currency_code: form.currency_code,
          financial_year_end_month: Number(form.financial_year_end_month),
          industry: form.industry,
          business_model: form.business_model.trim() || null,
          employee_count: form.employee_count ? Number(form.employee_count) : null,
          annual_revenue: form.annual_revenue ? Number(form.annual_revenue) : null,
          website_url: form.website_url.trim() || null,
        },
        reporting_preferences: {
          frequency: form.reporting_frequency,
          structure: form.reporting_structure,
        },
        enabled_modules: form.enabled_modules,
        preferred_data_source: form.preferred_data_source,
        logo_selected: Boolean(logoPreview),
      });

      if (result.confirmation_required) {
        setDone(true);
      } else {
        router.replace("/onboarding");
      }
    } catch (signupError) {
      setError(getApiErrorMessage(signupError));
    } finally {
      setSaving(false);
    }
  }

  if (done) {
    return (
      <main className="flex min-h-screen items-center justify-center bg-[radial-gradient(circle_at_top_left,rgba(99,102,241,0.15),transparent_35%)] px-6">
        <Card className="w-full max-w-xl animate-rise rounded-3xl shadow-2xl">
          <CardHeader className="text-center">
            <CheckCircle2 className="mx-auto mb-4 size-14 text-emerald-600" />
            <CardTitle className="text-3xl">Confirm your email</CardTitle>
            <CardDescription className="text-base leading-7">
              We sent a confirmation link to <strong>{form.email}</strong>. Confirm it, then sign in to finish creating your FinCruiz workspace.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <Button size="lg" className="h-14 w-full text-base" onClick={() => router.push("/login")}>
              Go to sign in
            </Button>
          </CardContent>
        </Card>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-[radial-gradient(circle_at_10%_10%,rgba(99,102,241,0.12),transparent_28%),radial-gradient(circle_at_90%_16%,rgba(14,165,233,0.11),transparent_26%)] px-5 py-8">
      <div className="mx-auto max-w-6xl animate-rise">
        <div className="mb-8 flex items-center justify-between">
          <Link href="/" className="flex items-center gap-3">
            <div className="flex size-12 items-center justify-center rounded-2xl bg-primary text-primary-foreground shadow-lg">
              <BarChart3 className="size-5" />
            </div>
            <div>
              <p className="text-lg font-bold">FinCruiz</p>
              <p className="text-xs text-muted-foreground">Workspace registration</p>
            </div>
          </Link>

          <Link
            href="/login"
            className="rounded-xl border bg-white px-5 py-3 text-sm font-semibold shadow-sm transition hover:-translate-y-0.5 hover:shadow-md"
          >
            Already registered? Sign in
          </Link>
        </div>

        <div className="mb-7 grid grid-cols-2 gap-3 md:grid-cols-4">
          {["Account", "Company", "Reporting", "Review"].map((label, index) => {
            const active = step === index + 1;
            const complete = step > index + 1;

            return (
              <button
                key={label}
                type="button"
                onClick={() => complete && setStep(index + 1)}
                className={[
                  "flex min-h-14 items-center gap-3 rounded-2xl border px-4 text-left text-sm font-semibold transition duration-300",
                  active
                    ? "scale-[1.02] bg-primary text-primary-foreground shadow-xl"
                    : complete
                      ? "bg-emerald-50 text-emerald-800"
                      : "bg-white/80",
                ].join(" ")}
              >
                <span className="flex size-7 items-center justify-center rounded-full bg-white/20">
                  {complete ? <Check className="size-4" /> : index + 1}
                </span>
                {label}
              </button>
            );
          })}
        </div>

        <Card className="overflow-hidden rounded-[28px] border-white/60 bg-white/90 shadow-2xl backdrop-blur">
          <CardHeader className="border-b bg-gradient-to-r from-slate-50 to-indigo-50/50 px-7 py-7">
            <CardTitle className="text-2xl">
              {["Create your account", "Tell us about the company", "Configure reporting", "Review and create workspace"][step - 1]}
            </CardTitle>
            <CardDescription className="text-base">
              These details preconfigure reporting, KPIs, forecasting and your workspace.
            </CardDescription>
          </CardHeader>

          <CardContent className="p-7">
            {error ? (
              <Alert variant="destructive" className="mb-6">
                <AlertDescription>{error}</AlertDescription>
              </Alert>
            ) : null}

            <div key={step} className="animate-step-in">
              {step === 1 ? (
                <form className="grid gap-6 md:grid-cols-2" onSubmit={next}>
                  <Field label="Full name">
                    <Input className="h-12" value={form.full_name} onChange={(e) => set("full_name", e.target.value)} required />
                  </Field>
                  <Field label="Work email">
                    <Input className="h-12" type="email" value={form.email} onChange={(e) => set("email", e.target.value)} required />
                  </Field>
                  <Field label="Password">
                    <Input className="h-12" type="password" minLength={8} value={form.password} onChange={(e) => set("password", e.target.value)} required />
                  </Field>
                  <Field label="Confirm password">
                    <Input className="h-12" type="password" minLength={8} value={form.confirm_password} onChange={(e) => set("confirm_password", e.target.value)} required />
                  </Field>
                  <div className="md:col-span-2 flex justify-end">
                    <Button type="submit" size="lg" className="h-14 min-w-48 text-base">
                      Continue
                      <ArrowRight className="size-5" />
                    </Button>
                  </div>
                </form>
              ) : null}

              {step === 2 ? (
                <form className="grid gap-6 md:grid-cols-2" onSubmit={next}>
                  <Field label="Legal company name">
                    <Input className="h-12" value={form.legal_name} onChange={(e) => set("legal_name", e.target.value)} required />
                  </Field>
                  <Field label="Trading name">
                    <Input className="h-12" value={form.trading_name} onChange={(e) => set("trading_name", e.target.value)} />
                  </Field>
                  <Field label="Industry">
                    <Select value={form.industry} onChange={(value) => set("industry", value)} options={industries.map((value) => [value, value])} />
                  </Field>
                  <Field label="Country">
                    <Select
                      value={form.country_code}
                      onChange={(value) => {
                        const row = countries.find(([code]) => code === value);
                        set("country_code", value);
                        if (row) set("currency_code", row[2]);
                      }}
                      options={countries.map(([code, name]) => [code, name])}
                    />
                  </Field>
                  <Field label="Currency">
                    <Select
                      value={form.currency_code}
                      onChange={(value) => set("currency_code", value)}
                      options={currencies.map(([code, name]) => [code, `${code} — ${name}`])}
                    />
                  </Field>
                  <Field label="Tax ID / ABN / GSTIN">
                    <Input className="h-12" value={form.tax_id} onChange={(e) => set("tax_id", e.target.value)} />
                  </Field>
                  <Field label="Business model">
                    <Input className="h-12" value={form.business_model} onChange={(e) => set("business_model", e.target.value)} placeholder="SaaS, retail, services..." />
                  </Field>
                  <Field label="Financial year end month">
                    <Select
                      value={form.financial_year_end_month}
                      onChange={(value) => set("financial_year_end_month", value)}
                      options={[
                        ["1", "January"], ["2", "February"], ["3", "March"], ["4", "April"],
                        ["5", "May"], ["6", "June"], ["7", "July"], ["8", "August"],
                        ["9", "September"], ["10", "October"], ["11", "November"], ["12", "December"],
                      ]}
                    />
                  </Field>
                  <Field label="Employee count">
                    <Input className="h-12" type="number" min={0} value={form.employee_count} onChange={(e) => set("employee_count", e.target.value)} />
                  </Field>
                  <Field label="Annual revenue">
                    <Input className="h-12" type="number" min={0} value={form.annual_revenue} onChange={(e) => set("annual_revenue", e.target.value)} />
                  </Field>
                  <Field label="Website">
                    <Input className="h-12" type="url" value={form.website_url} onChange={(e) => set("website_url", e.target.value)} />
                  </Field>

                  <div className="space-y-2">
                    <Label>Company logo</Label>
                    <label className="flex min-h-28 cursor-pointer items-center gap-4 rounded-2xl border border-dashed bg-slate-50 p-4 transition hover:border-indigo-400 hover:bg-indigo-50/50">
                      {logoPreview ? (
                        <img src={logoPreview} alt="Company logo preview" className="size-16 rounded-xl border bg-white object-contain p-1" />
                      ) : (
                        <div className="flex size-14 items-center justify-center rounded-xl bg-white shadow-sm">
                          <ImagePlus className="size-6 text-indigo-600" />
                        </div>
                      )}
                      <div>
                        <p className="font-semibold">{logoPreview ? "Logo selected" : "Upload company logo"}</p>
                        <p className="mt-1 text-xs text-muted-foreground">PNG, JPG or WebP · Maximum 2 MB</p>
                      </div>
                      <input type="file" accept="image/png,image/jpeg,image/webp" className="hidden" onChange={handleLogo} />
                    </label>
                  </div>

                  <div className="md:col-span-2 flex justify-between">
                    <Button type="button" size="lg" variant="outline" onClick={() => setStep(1)}>
                      <ArrowLeft className="size-4" />
                      Back
                    </Button>
                    <Button type="submit" size="lg" className="h-14 min-w-48 text-base">
                      Continue
                      <ArrowRight className="size-5" />
                    </Button>
                  </div>
                </form>
              ) : null}

              {step === 3 ? (
                <form className="space-y-8" onSubmit={next}>
                  <div className="grid gap-6 md:grid-cols-2">
                    <Field label="Reporting frequency">
                      <Select value={form.reporting_frequency} onChange={(value) => set("reporting_frequency", value)} options={["Monthly", "Quarterly", "Annual"].map((value) => [value, value])} />
                    </Field>

                    <Field label="Preferred data source">
                      <Select
                        value={form.preferred_data_source}
                        onChange={(value) => set("preferred_data_source", value)}
                        options={[
                          ["Manual Excel / CSV", "Manual Excel / CSV"],
                          ["API Connection — Coming Soon", "Direct API connection — Coming Soon"],
                        ]}
                      />
                    </Field>
                  </div>

                  <div>
                    <Label className="text-base">Reporting structure</Label>
                    <p className="mt-1 text-sm text-muted-foreground">
                      You can change this later. Both views are recommended for businesses with branches or business units.
                    </p>
                    <div className="mt-4 grid gap-4 lg:grid-cols-3">
                      {reportingStructures.map(({ value, title, text, icon: Icon }) => {
                        const selected = form.reporting_structure === value;

                        return (
                          <button
                            key={value}
                            type="button"
                            onClick={() => set("reporting_structure", value)}
                            className={[
                              "relative rounded-2xl border p-5 text-left transition duration-300 hover:-translate-y-1 hover:shadow-lg",
                              selected ? "border-indigo-500 bg-indigo-50 ring-2 ring-indigo-100" : "bg-white",
                            ].join(" ")}
                          >
                            {selected ? (
                              <span className="absolute right-4 top-4 flex size-7 items-center justify-center rounded-full bg-indigo-600 text-white">
                                <Check className="size-4" />
                              </span>
                            ) : null}
                            <div className="flex size-11 items-center justify-center rounded-xl bg-slate-950 text-white">
                              <Icon className="size-5" />
                            </div>
                            <p className="mt-5 font-bold">{title}</p>
                            <p className="mt-2 text-sm leading-6 text-muted-foreground">{text}</p>
                          </button>
                        );
                      })}
                    </div>
                  </div>

                  <div className="rounded-2xl border bg-slate-50 p-5">
                    <div className="flex items-center gap-3">
                      <div className="flex size-11 items-center justify-center rounded-xl bg-white shadow-sm">
                        <CloudCog className="size-5 text-indigo-600" />
                      </div>
                      <div>
                        <p className="font-bold">Direct accounting-system connections</p>
                        <p className="text-sm text-muted-foreground">
                          Xero, QuickBooks, Business Central, NetSuite, SAP and other API connectors are coming soon.
                        </p>
                      </div>
                      <span className="ml-auto rounded-full bg-indigo-100 px-3 py-1 text-xs font-bold text-indigo-700">COMING SOON</span>
                    </div>
                  </div>

                  <div>
                    <Label className="text-base">Enabled modules</Label>
                    <div className="mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
                      {modules.map((module) => (
                        <label
                          key={module}
                          className="flex cursor-pointer items-center gap-3 rounded-xl border bg-white p-4 transition hover:border-indigo-300 hover:shadow-sm"
                        >
                          <input
                            type="checkbox"
                            checked={form.enabled_modules.includes(module)}
                            onChange={(event) =>
                              set(
                                "enabled_modules",
                                event.target.checked
                                  ? [...form.enabled_modules, module]
                                  : form.enabled_modules.filter((item) => item !== module),
                              )
                            }
                          />
                          {module}
                        </label>
                      ))}
                    </div>
                  </div>

                  <div className="flex justify-between">
                    <Button type="button" size="lg" variant="outline" onClick={() => setStep(2)}>
                      <ArrowLeft className="size-4" />
                      Back
                    </Button>
                    <Button type="submit" size="lg" className="h-14 min-w-48 text-base">
                      Review
                      <ArrowRight className="size-5" />
                    </Button>
                  </div>
                </form>
              ) : null}

              {step === 4 ? (
                <div className="space-y-7">
                  <div className="grid gap-4 md:grid-cols-2">
                    <Summary title="Account" lines={[form.full_name, form.email]} />
                    <Summary title="Company" lines={[form.legal_name, `${countryName} · ${form.currency_code}`, form.industry]} />
                    <Summary title="Reporting" lines={[form.reporting_frequency, form.reporting_structure, form.preferred_data_source]} />
                    <Summary title="Modules" lines={form.enabled_modules} />
                  </div>

                  <div className="rounded-2xl bg-slate-950 p-6 text-white">
                    <div className="flex items-start gap-4">
                      <FileSpreadsheet className="mt-1 size-6 text-indigo-300" />
                      <div>
                        <p className="font-bold">Ready to create your finance workspace</p>
                        <p className="mt-2 text-sm leading-6 text-slate-300">
                          Your preferences will configure the first reporting experience. You can refine mappings, branches, modules and report settings after setup.
                        </p>
                      </div>
                    </div>
                  </div>

                  <div className="flex flex-col gap-3 sm:flex-row sm:justify-between">
                    <Button size="lg" variant="outline" onClick={() => setStep(3)}>
                      <ArrowLeft className="size-4" />
                      Back
                    </Button>
                    <Button
                      size="lg"
                      className="h-16 min-w-64 text-lg font-bold shadow-xl"
                      onClick={() => void submit()}
                      disabled={saving}
                    >
                      {saving ? <Loader2 className="size-5 animate-spin" /> : <Sparkles className="size-5" />}
                      Create my FinCruiz account
                    </Button>
                  </div>
                </div>
              ) : null}
            </div>
          </CardContent>
        </Card>
      </div>
    </main>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="space-y-2">
      <Label>{label}</Label>
      {children}
    </div>
  );
}

function Select({
  value,
  onChange,
  options,
}: {
  value: string;
  onChange: (value: string) => void;
  options: string[][];
}) {
  return (
    <select
      className="h-12 w-full rounded-md border bg-background px-3"
      value={value}
      onChange={(event) => onChange(event.target.value)}
    >
      {options.map(([optionValue, label]) => (
        <option key={optionValue} value={optionValue}>
          {label}
        </option>
      ))}
    </select>
  );
}

function Summary({ title, lines }: { title: string; lines: string[] }) {
  return (
    <div className="rounded-2xl border bg-muted/20 p-5">
      <p className="font-bold">{title}</p>
      {lines.filter(Boolean).map((line) => (
        <p key={line} className="mt-2 text-sm text-muted-foreground">
          {line}
        </p>
      ))}
    </div>
  );
}
