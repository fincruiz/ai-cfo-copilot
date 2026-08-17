import { clearOnboardingDraft, loadOnboardingDraft, saveOnboardingDraft } from "@/lib/commercial-onboarding";

describe("commercial onboarding draft", () => {
  beforeEach(() => window.localStorage.clear());

  it("persists and restores the guided setup", () => {
    const draft = {
      step: 2, legalName: "Acme", tradingName: "", countryCode: "AU", currencyCode: "AUD",
      fye: "6", industry: "Technology / SaaS", businessModel: "Subscription",
      employees: "", revenue: "", website: "", registration: "", dataRoute: "csv" as const,
    };
    saveOnboardingDraft(draft);
    expect(loadOnboardingDraft()).toEqual(draft);
    clearOnboardingDraft();
    expect(loadOnboardingDraft()).toBeNull();
  });
});
