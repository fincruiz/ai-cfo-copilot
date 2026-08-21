"use client";

import Link from "next/link";
import { useEffect, useMemo, useRef, useState } from "react";
import { usePathname, useRouter } from "next/navigation";
import {
  BarChart3,
  BrainCircuit,
  Building2,
  ChevronDown,
  CreditCard,
  FileBarChart,
  FileInput,
  FileText,
  Gauge,
  Handshake,
  History,
  LayoutDashboard,
  LifeBuoy,
  LogOut,
  Menu,
  PanelLeftClose,
  PanelLeftOpen,
  PlugZap,
  Presentation,
  Search,
  Settings,
  ShieldCheck,
  SlidersHorizontal,
  Sparkles,
  TrendingUp,
  Upload,
  UserRound,
  WandSparkles,
  X,
  MessageSquareText,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import { ThemeToggle } from "@/components/theme-toggle";
import { AICFOFloating } from "@/components/ai-cfo-floating";
import { FeatureExplorer, type Capability } from "@/components/feature-explorer";
import { authService } from "@/services/auth-service";
import { companyService } from "@/services/company-service";
import { marketService, type MarketProfile } from "@/services/market-service";
import { usageService } from "@/services/usage-service";
import { WorkspaceScopeSelector } from "@/components/workspace-scope-selector";
import { ReportingPeriodIndicator } from "@/components/reporting-period-indicator";
import { ContextualAIBar } from "@/components/contextual-ai-bar";
import { BetaFeedbackButton } from "@/components/beta-feedback-button";

type NavItem = {
  label: string;
  href: string;
  icon: typeof LayoutDashboard;
  description: string;
  keywords?: string;
};

type NavGroup = {
  label: string;
  icon: typeof LayoutDashboard;
  description: string;
  items: NavItem[];
};

const homeItem: NavItem = {
  label: "Home",
  href: "/dashboard",
  icon: LayoutDashboard,
  description: "Executive command centre for performance, priorities, confidence and next actions.",
  keywords: "dashboard executive management command centre",
};

const navigationGroups: NavGroup[] = [
  {
    label: "Performance",
    icon: Gauge,
    description: "Understand actual performance and the drivers behind it.",
    items: [
      { label: "Financials", href: "/dashboard/reports", icon: FileBarChart, description: "P&L, Balance Sheet, trial balance and ledger evidence.", keywords: "p&l pnl balance sheet statements" },
      { label: "KPIs & trends", href: "/dashboard/analytics", icon: BarChart3, description: "Monthly movements, ratios, trends and variance patterns.", keywords: "kpis ratios trends variance" },
      { label: "Working capital", href: "/dashboard/working-capital", icon: Handshake, description: "Receivables, payables, overdue exposure and cash conversion.", keywords: "ar ap receivables payables collections cash" },
      { label: "Branches", href: "/dashboard/branches", icon: Building2, description: "Compare branches and business units without losing the consolidated view.", keywords: "branch division business unit" },
      { label: "Benchmarking", href: "/dashboard/benchmarking", icon: BarChart3, description: "Compare performance with relevant industry and economic context.", keywords: "benchmark industry economic" },
    ],
  },
  {
    label: "Plan",
    icon: TrendingUp,
    description: "Build budgets, forecasts and integrated scenarios.",
    items: [
      { label: "Budget", href: "/dashboard/native-planning", icon: SlidersHorizontal, description: "Build a management budget directly in FinCruiz.", keywords: "budget plan target" },
      { label: "Forecast", href: "/dashboard/forecasting", icon: TrendingUp, description: "Project performance from actuals and management assumptions.", keywords: "forecast projection outlook" },
      { label: "Three-way model", href: "/dashboard/three-way-forecast", icon: TrendingUp, description: "Connect P&L, Balance Sheet and Cash Flow in one forecast.", keywords: "three way cash balance sheet pnl integrated" },
      { label: "Scenarios", href: "/dashboard/decision-simulator", icon: Sparkles, description: "Model hiring, pricing, growth, capex and working-capital decisions.", keywords: "what if decision scenario hire pricing capex" },
      { label: "Budget vs actual", href: "/dashboard/planning", icon: SlidersHorizontal, description: "Compare plans, budgets and scenarios against actual performance.", keywords: "budget actual variance scenario" },
    ],
  },
  {
    label: "Decisions",
    icon: BrainCircuit,
    description: "Move from financial signal to management action.",
    items: [
      { label: "Ask FinCruiz", href: "/dashboard/intelligence", icon: BrainCircuit, description: "Ask management questions and investigate evidence-backed answers.", keywords: "ai cfo insights intelligence question" },
      { label: "Visual analysis", href: "/dashboard/bi", icon: BarChart3, description: "Explore management charts and conversational BI.", keywords: "bi charts graphs visualization" },
      { label: "Power of One", href: "/dashboard/power-of-one", icon: SlidersHorizontal, description: "Test the impact of small changes in price, volume, costs and working capital.", keywords: "sensitivity driver scenario" },
    ],
  },
  {
    label: "Reports",
    icon: FileText,
    description: "Turn the same governed numbers into management-ready output.",
    items: [
      { label: "Management reports", href: "/dashboard/board-reports", icon: FileText, description: "Management and board-ready reporting views.", keywords: "management report board commentary" },
      { label: "Board packs", href: "/dashboard/board-packs", icon: FileBarChart, description: "Review generated board-pack records and outputs.", keywords: "board pack" },
      { label: "Build a board pack", href: "/dashboard/board-pack-builder", icon: Presentation, description: "Create a board pack with outlook, risks, priorities and decisions.", keywords: "board presentation" },
      { label: "PowerPoint", href: "/dashboard/powerpoint", icon: Presentation, description: "Create presentation-ready finance and management output.", keywords: "ppt pptx slides export" },
    ],
  },
  {
    label: "Data",
    icon: PlugZap,
    description: "Connect, validate and govern the data behind every answer.",
    items: [
      { label: "Integrations", href: "/dashboard/integrations", icon: PlugZap, description: "Connect Xero, Zoho Books and Tally to the canonical finance model.", keywords: "xero zoho tally erp connection" },
      { label: "Upload", href: "/dashboard/uploads", icon: Upload, description: "Upload the General Ledger and validate it before activation.", keywords: "gl general ledger csv" },
      { label: "Import centre", href: "/dashboard/import-center", icon: FileInput, description: "Load supporting finance datasets such as AR, AP and Chart of Accounts.", keywords: "ar ap coa import" },
      { label: "Account mapping", href: "/dashboard/mapping", icon: WandSparkles, description: "Map source accounts into consistent reporting groups.", keywords: "coa mapping accounts" },
      { label: "Reliability", href: "/dashboard/data-reliability", icon: ShieldCheck, description: "Review balance, mappings, lineage, branches and ingestion health.", keywords: "reliability assurance reconciliation finance truth" },
    ],
  },
  {
    label: "Admin",
    icon: Settings,
    description: "Workspace access, governance, billing and support.",
    items: [
      { label: "Company profile", href: "/dashboard/profile", icon: UserRound, description: "Maintain company details used across FinCruiz.", keywords: "company profile" },
      { label: "Access & permissions", href: "/dashboard/access", icon: ShieldCheck, description: "Manage workspace members, roles and access rights.", keywords: "roles users permissions security" },
      { label: "Plan & billing", href: "/dashboard/subscription", icon: CreditCard, description: "Review plans, trial status, entitlements and billing.", keywords: "subscription pricing trial billing" },
      { label: "Data & privacy", href: "/dashboard/settings", icon: Settings, description: "Control resets, privacy and permanent deletion actions.", keywords: "privacy reset delete data" },
      { label: "Audit trail", href: "/dashboard/audit", icon: History, description: "Review important workspace actions and configuration changes.", keywords: "audit history activity" },
      { label: "Support", href: "/dashboard/support", icon: LifeBuoy, description: "Check platform and workspace diagnostics or contact support.", keywords: "support help health diagnostics" },
      { label: "Beta feedback", href: "/dashboard/beta-feedback", icon: MessageSquareText, description: "Review tester feedback and launch issues.", keywords: "beta feedback testing bugs" },
    ],
  },
];

const hiddenCapabilities: Capability[] = [
  { group: "Performance", label: "KPIs", href: "/dashboard/kpis", description: "Financial ratios and performance indicators with interpretation.", keywords: "kpi ratio liquidity margin" },
  { group: "Plan", label: "Planning workspace", href: "/dashboard/planning", description: "Compare budgets and scenarios with actual performance.", keywords: "budget scenario actual" },
  { group: "Plan", label: "Native planning", href: "/dashboard/native-planning", description: "Build budgets directly inside FinCruiz.", keywords: "native budget planning" },
  { group: "Performance", label: "Analytics", href: "/dashboard/analytics", description: "Explore trends, branches and variance patterns.", keywords: "analytics trend variance" },
  { group: "Decisions", label: "Decision simulator", href: "/dashboard/decision-simulator", description: "Test management decisions through the integrated finance model.", keywords: "scenario decision what if" },
  { group: "Getting started", label: "Getting started", href: "/dashboard/getting-started", description: "Complete the essential steps for a trustworthy management workspace.", keywords: "onboarding setup start" },
];

const capabilities: Capability[] = [
  { group: "Home", label: homeItem.label, description: homeItem.description, href: homeItem.href, keywords: homeItem.keywords },
  ...navigationGroups.flatMap((group) => group.items.map((item) => ({ group: group.label, label: item.label, description: item.description, href: item.href, keywords: item.keywords }))),
  ...hiddenCapabilities,
];

export default function DashboardLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  const pathname = usePathname();
  const router = useRouter();
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [explorerOpen, setExplorerOpen] = useState(false);
  const [isAuthorizing, setIsAuthorizing] = useState(true);
  const [companyRole, setCompanyRole] = useState("");
  const [marketProfile, setMarketProfile] = useState<MarketProfile | null>(null);
  const navScrollRef = useRef<HTMLElement | null>(null);

  const activeGroup = useMemo(
    () => navigationGroups.find((group) => group.items.some((item) => pathname === item.href || pathname.startsWith(`${item.href}/`)))?.label ?? "",
    [pathname],
  );
  const activeItem = useMemo(
    () => navigationGroups.flatMap((group) => group.items.map((item) => ({ ...item, group: group.label }))).find((item) => pathname === item.href || pathname.startsWith(`${item.href}/`)),
    [pathname],
  );
  const [openGroup, setOpenGroup] = useState("");

  useEffect(() => {
    setCollapsed(window.localStorage.getItem("fincruiz_sidebar_collapsed") === "true");
    setOpenGroup(activeGroup);
    let cancelled = false;
    async function authorizeDashboard() {
      if (!authService.hasAccessToken()) { router.replace("/login"); return; }
      try {
        await authService.getCurrentUser();
        await authService.getCurrentCompany();
        companyService.getAccess().then((access) => { if (!cancelled) setCompanyRole(access.role); }).catch(() => undefined);
        marketService.current().then((market) => { if (!cancelled) setMarketProfile(market); }).catch(() => undefined);
        if (!cancelled) setIsAuthorizing(false);
      } catch (error: unknown) {
        if (cancelled) return;
        const apiError = error as { response?: { status?: number; data?: { error_code?: string } } };
        if (apiError.response?.status === 404 && apiError.response?.data?.error_code === "COMPANY_MEMBERSHIP_NOT_FOUND") { router.replace("/onboarding"); return; }
        authService.logout(); router.replace("/login");
      }
    }
    void authorizeDashboard();
    return () => { cancelled = true; };
  }, [router, activeGroup]);

  useEffect(() => {
    if (!isAuthorizing) usageService.track("page_viewed", { area: pathname.startsWith("/dashboard/") ? pathname.split("/")[2] || "dashboard" : "dashboard" });
  }, [pathname, isAuthorizing]);

  useEffect(() => {
    const node = navScrollRef.current;
    if (!node) return;
    node.scrollTop = Number(window.sessionStorage.getItem("fincruiz_sidebar_scroll") || "0");
    const save = () => window.sessionStorage.setItem("fincruiz_sidebar_scroll", String(node.scrollTop));
    node.addEventListener("scroll", save, { passive: true });
    return () => node.removeEventListener("scroll", save);
  }, [collapsed]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "k") {
        event.preventDefault();
        openExplorer("keyboard");
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  });

  function toggleSidebar() {
    setCollapsed((current) => {
      const next = !current;
      window.localStorage.setItem("fincruiz_sidebar_collapsed", String(next));
      return next;
    });
  }

  function openExplorer(source: string) {
    usageService.track("explore_opened", { source });
    setExplorerOpen(true);
    setMobileOpen(false);
  }

  function handleLogout() {
    authService.logout();
    router.replace("/login");
  }

  if (isAuthorizing) {
    return <div className="flex min-h-screen items-center justify-center bg-background"><div className="text-center"><div className="mx-auto size-8 animate-spin rounded-full border-2 border-muted border-t-primary"/><p className="mt-3 text-sm text-muted-foreground">Securing your workspace…</p></div></div>;
  }

  const SidebarContent = ({ mobile = false }: { mobile?: boolean }) => {
    const expanded = !collapsed || mobile;
    return (
      <>
        <div className="flex h-[72px] items-center px-4">
          <Link href="/dashboard" onClick={() => setMobileOpen(false)} className="flex min-w-0 items-center gap-3">
            <span className="fincruiz-brand-mark"><BarChart3 className="size-[18px]" /></span>
            {expanded ? <div className="min-w-0"><p className="truncate text-[15px] font-bold tracking-[-.02em]">FinCruiz</p><p className="truncate text-[10px] font-semibold uppercase tracking-[.14em] text-muted-foreground">Finance operating system</p></div> : null}
          </Link>
          {mobile ? <button type="button" onClick={() => setMobileOpen(false)} className="ml-auto flex size-9 items-center justify-center rounded-xl hover:bg-muted"><X className="size-4"/></button> : null}
        </div>

        <div className="px-3 pb-3">
          <button type="button" onClick={() => openExplorer("sidebar_search")} className={`fincruiz-command-trigger ${expanded ? "w-full justify-start" : "mx-auto size-11 justify-center px-0"}`} title="Explore FinCruiz · Ctrl K">
            <Search className="size-4 shrink-0"/>
            {expanded ? <><span className="min-w-0 flex-1 truncate text-left">Search FinCruiz</span><kbd className="hidden rounded-md border bg-background px-1.5 py-0.5 text-[10px] font-semibold text-muted-foreground xl:inline">⌘K</kbd></> : null}
          </button>
        </div>

        <nav ref={navScrollRef} className="fincruiz-scroll-stable min-h-0 flex-1 overflow-y-auto overscroll-contain px-3 pb-4">
          <Link href="/dashboard" onClick={() => setMobileOpen(false)} className={`fincruiz-nav-item ${pathname === "/dashboard" ? "fincruiz-nav-item-active" : ""} ${expanded ? "" : "justify-center px-0"}`} title={homeItem.description}>
            <LayoutDashboard className="size-[17px] shrink-0"/>{expanded ? <span>Home</span> : null}
          </Link>

          <div className="my-3 h-px bg-border/70" />

          {navigationGroups.map((group) => {
            const GroupIcon = group.icon;
            const active = activeGroup === group.label;
            const isOpen = openGroup === group.label;
            return <div key={group.label} className="mb-1">
              <button
                type="button"
                onClick={() => {
                  if (!expanded) { setCollapsed(false); setOpenGroup(group.label); return; }
                  setOpenGroup((current) => current === group.label ? "" : group.label);
                }}
                className={`fincruiz-nav-group ${active ? "fincruiz-nav-group-active" : ""} ${expanded ? "" : "justify-center px-0"}`}
                title={`${group.label} — ${group.description}`}
              >
                <GroupIcon className="size-[17px] shrink-0"/>
                {expanded ? <><span className="min-w-0 flex-1 truncate text-left">{group.label}</span><ChevronDown className={`size-3.5 shrink-0 transition-transform ${isOpen ? "rotate-180" : ""}`}/></> : null}
              </button>

              {expanded && isOpen ? <div className="ml-[17px] mt-1 border-l border-border/70 pl-3">
                {group.items.map((item) => {
                  const Icon = item.icon;
                  const itemActive = pathname === item.href || pathname.startsWith(`${item.href}/`);
                  return <Link key={`${group.label}-${item.label}`} href={item.href} onClick={() => { usageService.track("navigation_feature_opened", { feature: item.label, group: group.label }); setMobileOpen(false); }} className={`fincruiz-nav-subitem ${itemActive ? "fincruiz-nav-subitem-active" : ""}`} title={item.description}>
                    <Icon className="size-3.5 shrink-0"/><span className="truncate">{item.label}</span>
                  </Link>;
                })}
              </div> : null}
            </div>;
          })}
        </nav>

        <div className="border-t border-border/70 p-3">
          {expanded && marketProfile ? <Link href="/pricing" className="mb-2 flex items-center justify-between rounded-xl px-3 py-2 text-xs text-muted-foreground hover:bg-muted"><span>{marketProfile.country_name}</span><span className="font-semibold">{marketProfile.currency_code}</span></Link> : null}
          <Button type="button" variant="ghost" className={expanded ? "w-full justify-start" : "w-full justify-center px-0"} onClick={handleLogout}><LogOut className="size-4"/>{expanded ? "Sign out" : null}</Button>
        </div>
      </>
    );
  };

  const sectionLabel = pathname === "/dashboard" ? "Home" : activeGroup || "Workspace";
  const pageLabel = pathname === "/dashboard" ? "Executive command centre" : activeItem?.label ?? "FinCruiz";

  return <div className="h-dvh overflow-hidden bg-[var(--workspace-background)]">
    <aside className={`fincruiz-sidebar fixed inset-y-0 left-0 z-30 hidden h-dvh transition-all duration-300 lg:flex lg:flex-col ${collapsed ? "w-[76px]" : "w-[244px]"}`}><SidebarContent/></aside>

    {mobileOpen ? <div className="fixed inset-0 z-[90] bg-slate-950/45 backdrop-blur-sm lg:hidden" onMouseDown={(event) => event.target === event.currentTarget && setMobileOpen(false)}><aside className="flex h-full w-[min(88vw,300px)] flex-col bg-background shadow-2xl"><SidebarContent mobile/></aside></div> : null}

    <div className={`${collapsed ? "lg:pl-[76px]" : "lg:pl-[244px]"} flex h-dvh min-h-0 flex-col overflow-hidden transition-[padding] duration-300`}>
      <header className="fincruiz-topbar z-20 flex h-[72px] shrink-0 items-center justify-between gap-3 px-4 sm:px-6">
        <div className="flex min-w-0 items-center gap-3">
          <button type="button" onClick={() => setMobileOpen(true)} className="flex size-10 items-center justify-center rounded-xl border bg-background text-muted-foreground hover:bg-muted lg:hidden"><Menu className="size-4"/></button>
          <button type="button" onClick={toggleSidebar} className="hidden size-9 items-center justify-center rounded-xl text-muted-foreground hover:bg-muted lg:flex" title={collapsed ? "Expand sidebar" : "Collapse sidebar"}>{collapsed ? <PanelLeftOpen className="size-4"/> : <PanelLeftClose className="size-4"/>}</button>
          <div className="min-w-0">
            <p className="truncate text-[11px] font-semibold uppercase tracking-[.14em] text-muted-foreground">{sectionLabel}</p>
            <p className="truncate text-sm font-semibold tracking-[-.01em]">{pageLabel}</p>
          </div>
        </div>

        <button type="button" onClick={() => openExplorer("top_bar")} className="fincruiz-command-trigger hidden min-w-[260px] max-w-[360px] flex-1 md:flex" aria-label="Explore FinCruiz">
          <Search className="size-4"/><span className="min-w-0 flex-1 truncate text-left">Search, navigate or find a workflow</span><kbd className="rounded-md border bg-background px-1.5 py-0.5 text-[10px] font-semibold text-muted-foreground">⌘K</kbd>
        </button>

        <div className="flex shrink-0 items-center gap-2">
          <ReportingPeriodIndicator/><WorkspaceScopeSelector/>
          <div className="hidden rounded-xl border bg-background px-3 py-2 text-[10px] font-semibold capitalize text-muted-foreground 2xl:block">{companyRole ? companyRole.replaceAll("_", " ") : "workspace"}</div>
          <ThemeToggle/>
        </div>
      </header>

      <main className="fincruiz-scroll-stable min-h-0 flex-1 overflow-y-auto overscroll-contain px-4 py-5 sm:px-6 lg:px-8 lg:py-7"><div className="min-h-full">{children}<ContextualAIBar/></div></main>
    </div>

    <BetaFeedbackButton/>
    <AICFOFloating/>
    {explorerOpen ? <FeatureExplorer capabilities={capabilities} onClose={() => setExplorerOpen(false)}/> : null}
  </div>;
}
