"use client";

import Link from "next/link";
import { useEffect, useMemo, useState } from "react";
import { usePathname, useRouter } from "next/navigation";
import {
  BarChart3, BrainCircuit, PlugZap, Building2, ChevronDown, ChevronLeft, ChevronRight,
  FileBarChart, FileInput, FileText, Gauge, Handshake, History, LayoutDashboard, LogOut,
  PanelLeftClose, PanelLeftOpen, Presentation, Settings, ShieldCheck, SlidersHorizontal,
  TrendingUp, Upload, UserRound, WandSparkles, Search, Sparkles, Menu, X, Compass, Minimize2,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import { HelpTip } from "@/components/ui/help-tip";
import { ThemeToggle } from "@/components/theme-toggle";
import { AICFOFloating } from "@/components/ai-cfo-floating";
import { FeatureExplorer, type Capability } from "@/components/feature-explorer";
import { authService } from "@/services/auth-service";
import { companyService } from "@/services/company-service";
import { usageService } from "@/services/usage-service";

type NavItem = { label: string; href: string; icon: typeof LayoutDashboard; description: string; keywords?: string };
type NavGroup = { label: string; items: NavItem[]; defaultOpen?: boolean };

const navigationGroups: NavGroup[] = [
  { label: "Overview", defaultOpen: true, items: [
    { label: "Home", href: "/dashboard", icon: LayoutDashboard, description: "Your management-first view of performance, priorities, risks and next actions.", keywords: "dashboard executive management" },
    { label: "Intelligence Center", href: "/dashboard/intelligence", icon: BrainCircuit, description: "See what FinCruiz understands, the signals it is watching and the priorities it recommends.", keywords: "organizational brain insights ai" },
  ]},
  { label: "Finance & performance", defaultOpen: true, items: [
    { label: "Financial reports", href: "/dashboard/reports", icon: FileBarChart, description: "Profit & Loss, Balance Sheet, trial balance and core financial statements.", keywords: "p&l pnl balance sheet statements" },
    { label: "KPIs", href: "/dashboard/kpis", icon: Gauge, description: "Financial ratios and performance indicators with interpretation.", keywords: "ratios metrics margin liquidity" },
    { label: "Analytics", href: "/dashboard/analytics", icon: BarChart3, description: "Explore monthly movements, branch drivers, trends and variance patterns.", keywords: "trends monthly variance" },
    { label: "Working capital", href: "/dashboard/working-capital", icon: Handshake, description: "Understand receivables, payables, overdue exposure and cash tied up in operations.", keywords: "ar ap receivables payables cash collections" },
    { label: "Industry benchmarking", href: "/dashboard/benchmarking", icon: BarChart3, description: "Compare company performance with relevant external industry and economic context.", keywords: "benchmark market industry economic" },
  ]},
  { label: "Planning & forecasting", items: [
    { label: "Forecasting", href: "/dashboard/forecasting", icon: TrendingUp, description: "Project future revenue and financial performance from historical trends and assumptions.", keywords: "forecast projection" },
    { label: "Three-Way Forecast", href: "/dashboard/three-way-forecast", icon: TrendingUp, description: "Model P&L, Balance Sheet and Cash Flow together to answer decisions such as hiring, pricing, capex and cash impact.", keywords: "what if hire employees cash scenario integrated forecast" },
    { label: "Decision Simulator", href: "/dashboard/decision-simulator", icon: Sparkles, description: "Test hiring, pricing, revenue, working-capital and capex decisions through the integrated three-way financial model.", keywords: "what if decision scenario hire pricing cash simulator" },
    { label: "Power of One", href: "/dashboard/power-of-one", icon: SlidersHorizontal, description: "Test the impact of small changes in price, volume, margin, working capital and costs.", keywords: "sensitivity driver scenario" },
    { label: "Native Budget Builder", href: "/dashboard/native-planning", icon: SlidersHorizontal, description: "Build budgets directly in FinCruiz without relying on an external spreadsheet.", keywords: "budget plan" },
    { label: "Budgets & scenarios", href: "/dashboard/planning", icon: SlidersHorizontal, description: "Compare plans, budgets and scenarios against actual performance.", keywords: "budget scenario actual variance" },
  ]},
  { label: "Management reporting", items: [
    { label: "Board reports", href: "/dashboard/board-reports", icon: FileText, description: "Management and board-ready reporting views built from current company data.", keywords: "management report board" },
    { label: "Board packs", href: "/dashboard/board-packs", icon: FileBarChart, description: "Review generated board-pack records and reporting outputs.", keywords: "board pack" },
    { label: "Board Pack Builder", href: "/dashboard/board-pack-builder", icon: Presentation, description: "Create board packs with financials, outlook, risks, priorities and decisions required.", keywords: "ppt presentation board pack" },
    { label: "PowerPoint export", href: "/dashboard/powerpoint", icon: Presentation, description: "Export presentation-ready financial and management reporting.", keywords: "pptx export slides" },
  ]},
  { label: "Data & organization", items: [
    { label: "Integration Hub", href: "/dashboard/integrations", icon: PlugZap, description: "Connect systems such as Xero, Zoho Books and Tally to FinCruiz.", keywords: "xero zoho tally erp connection" },
    { label: "Upload data", href: "/dashboard/uploads", icon: Upload, description: "Upload the General Ledger and validate its structure before analysis.", keywords: "gl general ledger csv" },
    { label: "Import Centre", href: "/dashboard/import-center", icon: FileInput, description: "Load supporting finance datasets such as AR, AP and Chart of Accounts.", keywords: "ar ap coa import" },
    { label: "Account mapping", href: "/dashboard/mapping", icon: WandSparkles, description: "Map source accounts into consistent reporting groups used by FinCruiz.", keywords: "coa mapping accounts" },
    { label: "Branches", href: "/dashboard/branches", icon: Building2, description: "Manage branches or business units and analyse performance across them.", keywords: "branch division business unit" },
  ]},
  { label: "Governance", items: [
    { label: "Profile", href: "/dashboard/profile", icon: UserRound, description: "Maintain company details used across reports and intelligence.", keywords: "company profile" },
    { label: "Data & Privacy", href: "/dashboard/settings", icon: Settings, description: "Control demo data, module resets, full data reset and permanent account deletion.", keywords: "privacy reset delete data" },
    { label: "Access & permissions", href: "/dashboard/access", icon: ShieldCheck, description: "Manage workspace members, roles and access rights.", keywords: "roles users security permissions" },
    { label: "Audit trail", href: "/dashboard/audit", icon: History, description: "Review important workspace actions such as uploads, resets and configuration changes.", keywords: "audit history activity" },
  ]},
];

const capabilities: Capability[] = navigationGroups.flatMap((group) => group.items.map((item) => ({ group: group.label, label: item.label, description: item.description, href: item.href, keywords: item.keywords })));

export default function DashboardLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  const pathname = usePathname();
  const router = useRouter();
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [explorerOpen, setExplorerOpen] = useState(false);
  const [exploreExpanded, setExploreExpanded] = useState(true);
  const [isAuthorizing, setIsAuthorizing] = useState(true);
  const [companyRole, setCompanyRole] = useState("");
  const activeGroup = useMemo(() => navigationGroups.find((group) => group.items.some((item) => item.href === "/dashboard" ? pathname === "/dashboard" : pathname === item.href || pathname.startsWith(`${item.href}/`)))?.label, [pathname]);
  const [openGroups, setOpenGroups] = useState<Record<string, boolean>>(() => Object.fromEntries(navigationGroups.map((g) => [g.label, Boolean(g.defaultOpen)])));

  useEffect(() => {
    setCollapsed(window.localStorage.getItem("fincruiz_sidebar_collapsed") === "true");
    setExploreExpanded(window.localStorage.getItem("fincruiz_explore_compact") !== "true");
    if (activeGroup) setOpenGroups((current) => ({ ...current, [activeGroup]: true }));
    let cancelled = false;
    async function authorizeDashboard() {
      if (!authService.hasAccessToken()) { router.replace("/login"); return; }
      try {
        await authService.getCurrentUser();
        await authService.getCurrentCompany();
        companyService.getAccess().then((access) => { if (!cancelled) setCompanyRole(access.role); }).catch(() => undefined);
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

  function toggleSidebar() {
    setCollapsed((current) => { const next = !current; window.localStorage.setItem("fincruiz_sidebar_collapsed", String(next)); return next; });
  }
  function toggleExploreSize() {
    setExploreExpanded((current) => { const next = !current; window.localStorage.setItem("fincruiz_explore_compact", String(!next)); usageService.track("explore_banner_resized", { expanded: next }); return next; });
  }
  function openExplorer(source: string) { usageService.track("explore_opened", { source }); setExplorerOpen(true); setMobileOpen(false); }
  function handleLogout() { authService.logout(); router.replace("/login"); }

  if (isAuthorizing) return <div className="flex min-h-screen items-center justify-center bg-background"><div className="text-center"><div className="mx-auto size-8 animate-spin rounded-full border-2 border-muted border-t-primary"/><p className="mt-3 text-sm text-muted-foreground">Securing your workspace…</p></div></div>;

  const SidebarContent = ({ mobile = false }: { mobile?: boolean }) => (
    <>
      <div className="flex h-16 items-center border-b px-4">
        <div className="flex size-10 shrink-0 items-center justify-center rounded-xl bg-primary text-primary-foreground"><BarChart3 className="size-5" /></div>
        {!collapsed || mobile ? <div className="ml-3 min-w-0"><p className="truncate font-semibold tracking-tight">FinCruiz</p><p className="truncate text-xs text-muted-foreground">Business intelligence brain</p></div> : null}
        {mobile ? <button type="button" onClick={() => setMobileOpen(false)} className="ml-auto flex size-9 items-center justify-center rounded-lg hover:bg-muted"><X className="size-4"/></button> : <button type="button" onClick={toggleSidebar} className={["ml-auto flex size-9 items-center justify-center rounded-lg border bg-background text-muted-foreground transition hover:bg-muted hover:text-foreground", collapsed ? "absolute -right-4 top-4 shadow-md" : ""].join(" ")} title={collapsed ? "Expand sidebar" : "Collapse sidebar"}>{collapsed ? <ChevronRight className="size-4"/> : <ChevronLeft className="size-4"/>}</button>}
      </div>

      {(!collapsed || mobile) ? (
        <div className="border-b p-3">
          <div className={`relative overflow-hidden rounded-2xl border transition-all duration-300 ${exploreExpanded ? "bg-gradient-to-br from-primary/[.09] via-background to-sky-500/[.07] p-4 shadow-sm" : "bg-muted/30 p-2"}`}>
            {exploreExpanded ? <><div className="flex items-start gap-3"><div className="flex size-10 shrink-0 items-center justify-center rounded-xl bg-primary text-primary-foreground"><Compass className="size-4"/></div><div className="min-w-0 flex-1"><p className="font-semibold">Explore FinCruiz</p><p className="mt-1 text-xs leading-5 text-muted-foreground">See every capability or search by what you want to achieve.</p></div><button type="button" onClick={toggleExploreSize} className="flex size-7 items-center justify-center rounded-lg text-muted-foreground hover:bg-background" title="Keep Explore compact"><Minimize2 className="size-3.5"/></button></div><button type="button" onClick={() => openExplorer("sidebar_banner")} className="mt-4 flex w-full items-center justify-center gap-2 rounded-xl bg-primary px-3 py-2.5 text-sm font-semibold text-primary-foreground shadow-sm hover:opacity-90"><Search className="size-4"/>Find a capability</button><p className="mt-2 text-center text-[10px] text-muted-foreground">You can keep this compact anytime.</p></> : <div className="flex w-full items-center gap-2 rounded-xl"><button type="button" onClick={() => openExplorer("sidebar_compact")} className="flex min-w-0 flex-1 items-center gap-3 rounded-xl px-2 py-2 text-left hover:bg-muted"><Search className="size-4 shrink-0 text-primary"/><div className="min-w-0 flex-1"><p className="text-sm font-semibold">Explore FinCruiz</p><p className="truncate text-xs text-muted-foreground">Search every capability</p></div><Sparkles className="size-4 shrink-0 text-muted-foreground"/></button><button type="button" onClick={toggleExploreSize} className="rounded-lg px-2 py-2 text-[10px] font-medium text-muted-foreground hover:bg-muted hover:text-foreground">Expand</button></div>}
          </div>
        </div>
      ) : <div className="border-b p-3"><button type="button" onClick={() => openExplorer("collapsed_sidebar")} title="Explore all FinCruiz capabilities" className="mx-auto flex size-10 items-center justify-center rounded-xl border hover:bg-muted"><Search className="size-4"/></button></div>}

      <nav className="fincruiz-scroll-stable flex-1 overflow-y-auto px-3 py-3">
        {navigationGroups.map((group) => {
          const isOpen = Boolean(openGroups[group.label]);
          return <div key={group.label} className="mb-2">
            {(!collapsed || mobile) ? <button type="button" onClick={() => setOpenGroups((current) => ({ ...current, [group.label]: !isOpen }))} className="flex w-full items-center justify-between rounded-lg px-3 py-2 text-left text-[11px] font-semibold uppercase tracking-[.14em] text-muted-foreground hover:bg-muted/60"><span className="truncate">{group.label}</span><ChevronDown className={`size-3.5 shrink-0 transition ${isOpen ? "rotate-180" : ""}`}/></button> : <div className="mx-auto my-3 h-px w-8 bg-border"/>}
            {(collapsed && !mobile) || isOpen ? <div className="space-y-1">{group.items.map((item) => {
              const Icon = item.icon;
              const active = item.href === "/dashboard" ? pathname === "/dashboard" : pathname === item.href || pathname.startsWith(`${item.href}/`);
              return <Link key={`${group.label}-${item.label}`} href={item.href} onClick={() => { usageService.track("navigation_feature_opened", { feature: item.label, group: group.label }); setMobileOpen(false); }} className={["group/nav relative flex items-center rounded-xl text-sm font-medium transition-colors", collapsed && !mobile ? "justify-center px-2 py-2.5" : "gap-3 px-3 py-2.5", active ? "bg-primary text-primary-foreground shadow-sm" : "text-muted-foreground hover:bg-muted hover:text-foreground"].join(" ")}>
                <Icon className="size-4 shrink-0"/>{(!collapsed || mobile) ? <><span className="min-w-0 flex-1 truncate">{item.label}</span><HelpTip title={item.label} text={item.description} side="right"/></> : null}
              </Link>;
            })}</div> : null}
          </div>;
        })}
      </nav>
      <div className="border-t p-3"><Button type="button" variant="ghost" className={collapsed && !mobile ? "w-full justify-center px-0" : "w-full justify-start"} onClick={handleLogout}><LogOut className="size-4"/>{(!collapsed || mobile) ? "Sign out" : null}</Button></div>
    </>
  );

  return <div className="min-h-screen bg-muted/25">
    <aside className={["fixed inset-y-0 left-0 z-30 hidden border-r bg-background transition-all duration-300 lg:flex lg:flex-col", collapsed ? "w-20" : "w-72"].join(" ")}><SidebarContent/></aside>
    {mobileOpen ? <div className="fixed inset-0 z-[90] bg-slate-950/45 backdrop-blur-sm lg:hidden" onMouseDown={(e) => e.target === e.currentTarget && setMobileOpen(false)}><aside className="flex h-full w-[min(88vw,320px)] flex-col bg-background shadow-2xl"><SidebarContent mobile/></aside></div> : null}
    <div className={collapsed ? "transition-all duration-300 lg:pl-20" : "transition-all duration-300 lg:pl-72"}>
      <header className="sticky top-0 z-20 flex h-16 items-center justify-between border-b bg-background/90 px-4 backdrop-blur-xl sm:px-6">
        <div className="flex items-center gap-3"><button type="button" onClick={() => setMobileOpen(true)} className="flex size-9 items-center justify-center rounded-lg border text-muted-foreground hover:bg-muted lg:hidden"><Menu className="size-4"/></button><button type="button" onClick={toggleSidebar} className="hidden size-9 items-center justify-center rounded-lg border text-muted-foreground hover:bg-muted lg:flex" title={collapsed ? "Expand sidebar" : "Collapse sidebar"}>{collapsed ? <PanelLeftOpen className="size-4"/> : <PanelLeftClose className="size-4"/>}</button><div><p className="text-sm font-medium">FinCruiz Workspace</p><p className="text-xs capitalize text-muted-foreground">{companyRole ? `${companyRole.replaceAll("_", " ")} · ` : ""}Management intelligence</p></div></div>
        <div className="flex items-center gap-2"><button type="button" onClick={() => openExplorer("top_bar")} className="group hidden min-w-[190px] items-center gap-3 rounded-2xl border border-indigo-200/80 bg-gradient-to-r from-indigo-50 via-background to-sky-50 px-4 py-2 text-left shadow-sm transition hover:-translate-y-0.5 hover:border-indigo-300 hover:shadow-md dark:border-indigo-500/20 dark:from-indigo-950/30 dark:via-background dark:to-sky-950/20 sm:flex"><span className="flex size-9 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-600 to-sky-500 text-white"><Compass className="size-4"/></span><span className="min-w-0 flex-1"><span className="block text-sm font-semibold leading-4">Explore FinCruiz</span><span className="mt-1 block text-[10px] leading-3 text-muted-foreground">Find every capability</span></span><ChevronRight className="size-4 shrink-0 text-muted-foreground transition group-hover:translate-x-0.5"/></button><button type="button" onClick={() => openExplorer("top_bar_mobile")} className="flex size-10 items-center justify-center rounded-xl border bg-background sm:hidden" aria-label="Explore FinCruiz"><Compass className="size-4"/></button><ThemeToggle/></div>
      </header>
      <main className="p-4 sm:p-6 lg:p-8">{children}</main>
    </div>
    <AICFOFloating/>
    {explorerOpen ? <FeatureExplorer capabilities={capabilities} onClose={() => setExplorerOpen(false)}/> : null}
  </div>;
}
