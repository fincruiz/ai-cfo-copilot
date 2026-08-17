export type CommercialOnboardingDraft = {
  step: number;
  legalName: string;
  tradingName: string;
  countryCode: string;
  currencyCode: string;
  fye: string;
  industry: string;
  businessModel: string;
  employees: string;
  revenue: string;
  website: string;
  registration: string;
  dataRoute: "csv" | "xero" | "demo" | "";
};

const KEY = "fincruiz.commercial-onboarding.v1";

export function loadOnboardingDraft(): CommercialOnboardingDraft | null {
  if (typeof window === "undefined") return null;
  try {
    const raw = window.localStorage.getItem(KEY);
    return raw ? (JSON.parse(raw) as CommercialOnboardingDraft) : null;
  } catch {
    return null;
  }
}

export function saveOnboardingDraft(value: CommercialOnboardingDraft) {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(KEY, JSON.stringify(value));
}

export function clearOnboardingDraft() {
  if (typeof window === "undefined") return;
  window.localStorage.removeItem(KEY);
}
