"use client";

import {
  FormEvent,
  useEffect,
  useState,
} from "react";
import { useRouter } from "next/navigation";
import {
  Building2,
  ImagePlus,
  Loader2,
  LogOut,
} from "lucide-react";

import {
  Alert,
  AlertDescription,
} from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { getApiErrorMessage } from "@/lib/api";
import { authService } from "@/services/auth-service";
import { companyService } from "@/services/company-service";

export default function OnboardingPage() {
  const router = useRouter();

  const [legalName, setLegalName] =
    useState("");
  const [tradingName, setTradingName] =
    useState("");
  const [abn, setAbn] = useState("");
  const [countryCode, setCountryCode] =
    useState("AU");
  const [currencyCode, setCurrencyCode] =
    useState("AUD");
  const [
    financialYearEndMonth,
    setFinancialYearEndMonth,
  ] = useState("6");
  const [industry, setIndustry] =
    useState("");
  const [businessModel, setBusinessModel] =
    useState("");
  const [employeeCount, setEmployeeCount] =
    useState("");
  const [annualRevenue, setAnnualRevenue] =
    useState("");
  const [websiteUrl, setWebsiteUrl] =
    useState("");
  const [logoPreview, setLogoPreview] = useState("");
  const [logoFile, setLogoFile] = useState<File | null>(null);

  const [errorMessage, setErrorMessage] =
    useState("");
  const [isSubmitting, setIsSubmitting] =
    useState(false);
  const [isChecking, setIsChecking] =
    useState(true);

  useEffect(() => {
    async function checkExistingCompany() {
      if (!authService.hasAccessToken()) {
        router.replace("/login");
        return;
      }

      try {
        const pendingLogo = window.localStorage.getItem("fincruiz_pending_logo");
        if (pendingLogo) {
          setLogoPreview(pendingLogo);
          try {
            const response = await fetch(pendingLogo);
            const blob = await response.blob();
            const name = window.localStorage.getItem("fincruiz_pending_logo_name") || "company-logo.png";
            setLogoFile(new File([blob], name, { type: blob.type || "image/png" }));
          } catch {
            // The user may select the logo again before submitting.
          }
        }

        await companyService.getCurrentCompany();
        router.replace("/dashboard");
      } catch {
        try {
          const user = await authService.getCurrentUser();
          const details = (user.user_metadata?.company_details ?? {}) as Record<string, unknown>;
          if (details.legal_name) setLegalName(String(details.legal_name));
          if (details.trading_name) setTradingName(String(details.trading_name));
          if (details.abn) setAbn(String(details.abn));
          if (details.country_code) setCountryCode(String(details.country_code));
          if (details.currency_code) setCurrencyCode(String(details.currency_code));
          if (details.financial_year_end_month) setFinancialYearEndMonth(String(details.financial_year_end_month));
          if (details.industry) setIndustry(String(details.industry));
          if (details.business_model) setBusinessModel(String(details.business_model));
          if (details.employee_count != null) setEmployeeCount(String(details.employee_count));
          if (details.annual_revenue != null) setAnnualRevenue(String(details.annual_revenue));
          if (details.website_url) setWebsiteUrl(String(details.website_url));
        } catch {
          // The form remains editable when metadata is unavailable.
        }
        setIsChecking(false);
      }
    }

    void checkExistingCompany();
  }, [router]);

  async function handleSubmit(
    event: FormEvent<HTMLFormElement>,
  ) {
    event.preventDefault();

    setErrorMessage("");
    setIsSubmitting(true);

    try {
      await companyService.createCompany({
        legal_name: legalName.trim(),
        trading_name:
          tradingName.trim() || null,
        abn: abn.trim() || null,
        country_code:
          countryCode.trim().toUpperCase(),
        currency_code:
          currencyCode.trim().toUpperCase(),
        financial_year_end_month:
          Number(financialYearEndMonth),
        industry:
          industry.trim() || null,
        business_model:
          businessModel.trim() || null,
        employee_count:
          employeeCount === ""
            ? null
            : Number(employeeCount),
        annual_revenue:
          annualRevenue === ""
            ? null
            : Number(annualRevenue),
        logo_path: null,
        website_url:
          websiteUrl.trim() || null,
      });

      if (logoFile) {
        await companyService.uploadLogo(logoFile);
        window.localStorage.removeItem("fincruiz_pending_logo");
        window.localStorage.removeItem("fincruiz_pending_logo_name");
      }

      router.replace("/dashboard");
    } catch (error: unknown) {
      setErrorMessage(
        getApiErrorMessage(error),
      );
    } finally {
      setIsSubmitting(false);
    }
  }

  function handleLogout() {
    authService.logout();
    router.replace("/login");
  }

  if (isChecking) {
    return (
      <main className="flex min-h-screen items-center justify-center">
        <div className="flex items-center gap-3 text-muted-foreground">
          <Loader2 className="size-5 animate-spin" />
          Checking your workspace...
        </div>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-muted/30 px-6 py-10">
      <div className="mx-auto max-w-5xl">
        <div className="mb-8 flex items-center justify-between">
          <div>
            <p className="text-sm font-medium text-muted-foreground">
              FinCruiz onboarding
            </p>

            <h1 className="mt-2 text-3xl font-semibold tracking-tight">
              Set up your company
            </h1>

            <p className="mt-2 text-muted-foreground">
              Add the business details used for
              reporting, KPIs and AI insights.
            </p>
          </div>

          <Button
            type="button"
            variant="outline"
            onClick={handleLogout}
          >
            <LogOut />
            Sign out
          </Button>
        </div>

        <form onSubmit={handleSubmit}>
          <div className="grid gap-6 lg:grid-cols-[1fr_320px]">
            <Card>
              <CardHeader>
                <CardTitle>
                  Company details
                </CardTitle>

                <CardDescription>
                  You can update these details
                  later in settings.
                </CardDescription>
              </CardHeader>

              <CardContent className="space-y-6">
                {errorMessage ? (
                  <Alert variant="destructive">
                    <AlertDescription>
                      {errorMessage}
                    </AlertDescription>
                  </Alert>
                ) : null}

                <div className="grid gap-5 md:grid-cols-2">
                  <div className="space-y-2 md:col-span-2">
                    <Label htmlFor="legal-name">
                      Legal name
                    </Label>

                    <Input
                      id="legal-name"
                      value={legalName}
                      onChange={(event) =>
                        setLegalName(
                          event.target.value,
                        )
                      }
                      placeholder="Example Technologies Pty Ltd"
                      required
                      minLength={2}
                      maxLength={255}
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="trading-name">
                      Trading name
                    </Label>

                    <Input
                      id="trading-name"
                      value={tradingName}
                      onChange={(event) =>
                        setTradingName(
                          event.target.value,
                        )
                      }
                      placeholder="Example"
                      maxLength={255}
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="abn">
                      ABN
                    </Label>

                    <Input
                      id="abn"
                      value={abn}
                      onChange={(event) =>
                        setAbn(
                          event.target.value,
                        )
                      }
                      placeholder="Optional"
                      maxLength={20}
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="country-code">
                      Country code
                    </Label>

                    <Input
                      id="country-code"
                      value={countryCode}
                      onChange={(event) =>
                        setCountryCode(
                          event.target.value,
                        )
                      }
                      maxLength={2}
                      required
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="currency-code">
                      Currency code
                    </Label>

                    <Input
                      id="currency-code"
                      value={currencyCode}
                      onChange={(event) =>
                        setCurrencyCode(
                          event.target.value,
                        )
                      }
                      maxLength={3}
                      required
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="financial-year-end">
                      Financial year end month
                    </Label>

                    <Input
                      id="financial-year-end"
                      type="number"
                      min={1}
                      max={12}
                      value={financialYearEndMonth}
                      onChange={(event) =>
                        setFinancialYearEndMonth(
                          event.target.value,
                        )
                      }
                      required
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="industry">
                      Industry
                    </Label>

                    <Input
                      id="industry"
                      value={industry}
                      onChange={(event) =>
                        setIndustry(
                          event.target.value,
                        )
                      }
                      placeholder="Software"
                      maxLength={255}
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="business-model">
                      Business model
                    </Label>

                    <Input
                      id="business-model"
                      value={businessModel}
                      onChange={(event) =>
                        setBusinessModel(
                          event.target.value,
                        )
                      }
                      placeholder="SaaS"
                      maxLength={255}
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="employee-count">
                      Employee count
                    </Label>

                    <Input
                      id="employee-count"
                      type="number"
                      min={0}
                      value={employeeCount}
                      onChange={(event) =>
                        setEmployeeCount(
                          event.target.value,
                        )
                      }
                      placeholder="0"
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="annual-revenue">
                      Annual revenue
                    </Label>

                    <Input
                      id="annual-revenue"
                      type="number"
                      min={0}
                      step="0.01"
                      value={annualRevenue}
                      onChange={(event) =>
                        setAnnualRevenue(
                          event.target.value,
                        )
                      }
                      placeholder="0.00"
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2 md:col-span-2">
                    <Label htmlFor="website-url">
                      Website
                    </Label>

                    <Input
                      id="website-url"
                      type="url"
                      value={websiteUrl}
                      onChange={(event) =>
                        setWebsiteUrl(
                          event.target.value,
                        )
                      }
                      placeholder="https://example.com"
                      disabled={isSubmitting}
                    />
                  </div>

                  <div className="space-y-2 md:col-span-2">
                    <Label>Company logo</Label>
                    <label className="flex cursor-pointer items-center gap-4 rounded-2xl border border-dashed bg-muted/30 p-4 transition hover:border-indigo-400 hover:bg-indigo-50/50">
                      {logoPreview ? (
                        <img
                          src={logoPreview}
                          alt="Company logo preview"
                          className="size-16 rounded-xl border bg-white object-contain p-1"
                        />
                      ) : (
                        <div className="flex size-14 items-center justify-center rounded-xl bg-background shadow-sm">
                          <ImagePlus className="size-6 text-indigo-600" />
                        </div>
                      )}
                      <div>
                        <p className="font-medium">
                          {logoPreview ? "Logo ready to upload" : "Select company logo"}
                        </p>
                        <p className="mt-1 text-xs text-muted-foreground">
                          PNG, JPG or WebP · Maximum 2 MB
                        </p>
                      </div>
                      <input
                        type="file"
                        accept="image/png,image/jpeg,image/webp"
                        className="hidden"
                        onChange={(event) => {
                          const file = event.target.files?.[0];
                          if (!file) return;
                          if (file.size > 2 * 1024 * 1024) {
                            setErrorMessage("Logo must be smaller than 2 MB.");
                            return;
                          }
                          setLogoFile(file);
                          setLogoPreview(URL.createObjectURL(file));
                        }}
                      />
                    </label>
                  </div>
                </div>
              </CardContent>
            </Card>

            <div className="space-y-6">
              <Card>
                <CardHeader>
                  <div className="flex size-11 items-center justify-center rounded-xl bg-primary text-primary-foreground">
                    <Building2 className="size-5" />
                  </div>

                  <CardTitle className="pt-3">
                    Your first workspace
                  </CardTitle>

                  <CardDescription>
                    You will become the owner of
                    this company workspace.
                  </CardDescription>
                </CardHeader>

                <CardContent className="space-y-4">
                  <div className="rounded-lg border bg-muted/40 p-4 text-sm">
                    <p className="font-medium">
                      Defaults
                    </p>

                    <div className="mt-3 space-y-2 text-muted-foreground">
                      <p>
                        Country:{" "}
                        {countryCode || "—"}
                      </p>
                      <p>
                        Currency:{" "}
                        {currencyCode || "—"}
                      </p>
                      <p>
                        Financial year end: Month{" "}
                        {financialYearEndMonth ||
                          "—"}
                      </p>
                    </div>
                  </div>

                  <Button
                    type="submit"
                    size="lg"
                    className="w-full"
                    disabled={
                      isSubmitting ||
                      legalName.trim().length < 2
                    }
                  >
                    {isSubmitting ? (
                      <>
                        <Loader2 className="animate-spin" />
                        Creating workspace...
                      </>
                    ) : (
                      <>
                        <Building2 />
                        Create company
                      </>
                    )}
                  </Button>
                </CardContent>
              </Card>
            </div>
          </div>
        </form>
      </div>
    </main>
  );
}