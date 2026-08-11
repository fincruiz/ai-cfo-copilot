"use client";

import Link from "next/link";
import { useEffect, useState } from "react";
import { usePathname, useRouter } from "next/navigation";
import {
  BarChart3,
  Bot,
  Building2,
  ChevronLeft,
  ChevronRight,
  FileBarChart,
  FileInput,
  FileText,
  Gauge,
  Handshake,
  LayoutDashboard,
  LogOut,
  PanelLeftClose,
  PanelLeftOpen,
  Presentation,
  Settings,
  SlidersHorizontal,
  TrendingUp,
  Upload,
  UserRound,
  WandSparkles,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import { ThemeToggle } from "@/components/theme-toggle";
import { AICFOFloating } from "@/components/ai-cfo-floating";
import { authService } from "@/services/auth-service";

const navigationGroups = [
  {
    label: "Executive",
    items: [
      { label: "Dashboard", href: "/dashboard", icon: LayoutDashboard },
      { label: "AI CFO Assistant", href: "/dashboard", icon: Bot, disabled: true },
    ],
  },
  {
    label: "Data & setup",
    items: [
      { label: "Upload GL", href: "/dashboard/uploads", icon: Upload },
      { label: "Import Centre", href: "/dashboard/import-center", icon: FileInput },
      { label: "Account mapping", href: "/dashboard/mapping", icon: WandSparkles },
      { label: "Branches", href: "/dashboard/branches", icon: Building2 },
    ],
  },
  {
    label: "Reporting & analytics",
    items: [
      { label: "Financial reports", href: "/dashboard/reports", icon: FileBarChart },
      { label: "KPIs", href: "/dashboard/kpis", icon: Gauge },
      { label: "Analytics", href: "/dashboard/analytics", icon: BarChart3 },
      { label: "Working capital", href: "/dashboard/working-capital", icon: Handshake },
    ],
  },
  {
    label: "Planning & intelligence",
    items: [
      { label: "Forecasting", href: "/dashboard/forecasting", icon: TrendingUp },
      { label: "Three-Way Forecast", href: "/dashboard/three-way-forecast", icon: TrendingUp },
      { label: "Power of One", href: "/dashboard/power-of-one", icon: SlidersHorizontal },
      { label: "Native Budget Builder", href: "/dashboard/native-planning", icon: SlidersHorizontal },
      { label: "Budgets & scenarios", href: "/dashboard/planning", icon: SlidersHorizontal },
    ],
  },
  {
    label: "Board & exports",
    items: [
      { label: "Board reports", href: "/dashboard/board-reports", icon: FileText },
      { label: "Board packs", href: "/dashboard/board-packs", icon: FileBarChart },
      { label: "Board Pack Builder", href: "/dashboard/board-pack-builder", icon: Presentation },
      { label: "PowerPoint export", href: "/dashboard/powerpoint", icon: Presentation },
    ],
  },
  {
    label: "Administration",
    items: [
      { label: "Profile", href: "/dashboard/profile", icon: UserRound },
      { label: "Settings", href: "/dashboard/settings", icon: Settings },
    ],
  },
];

export default function DashboardLayout({
  children,
}: Readonly<{ children: React.ReactNode }>) {
  const pathname = usePathname();
  const router = useRouter();
  const [collapsed, setCollapsed] = useState(false);

  useEffect(() => {
    const saved = window.localStorage.getItem("fincruiz_sidebar_collapsed");
    setCollapsed(saved === "true");
  }, []);

  function toggleSidebar() {
    setCollapsed((current) => {
      const next = !current;
      window.localStorage.setItem("fincruiz_sidebar_collapsed", String(next));
      return next;
    });
  }

  function handleLogout() {
    authService.logout();
    router.replace("/login");
  }

  return (
    <div className="min-h-screen bg-muted/30">
      <aside
        className={[
          "fixed inset-y-0 left-0 z-30 hidden border-r bg-background transition-all duration-300 lg:flex lg:flex-col",
          collapsed ? "w-20" : "w-72",
        ].join(" ")}
      >
        <div className="flex h-16 items-center border-b px-4">
          <div className="flex size-10 shrink-0 items-center justify-center rounded-xl bg-primary text-primary-foreground">
            <BarChart3 className="size-5" />
          </div>

          {!collapsed ? (
            <div className="ml-3 min-w-0">
              <p className="truncate font-semibold tracking-tight">FinCruiz</p>
              <p className="truncate text-xs text-muted-foreground">AI CFO Platform</p>
            </div>
          ) : null}

          <button
            type="button"
            onClick={toggleSidebar}
            className={[
              "ml-auto flex size-9 items-center justify-center rounded-lg border bg-background text-muted-foreground transition hover:bg-muted hover:text-foreground",
              collapsed ? "absolute -right-4 top-4 shadow-md" : "",
            ].join(" ")}
            title={collapsed ? "Expand sidebar" : "Collapse sidebar"}
          >
            {collapsed ? <ChevronRight className="size-4" /> : <ChevronLeft className="size-4" />}
          </button>
        </div>

        {!collapsed ? (
          <div className="border-b px-4 py-4">
            <div className="rounded-xl border bg-muted/30 p-3">
              <p className="text-xs font-medium uppercase tracking-wider text-muted-foreground">Workspace</p>
              <p className="mt-1 text-sm font-semibold">Finance intelligence</p>
              <p className="mt-1 text-xs text-muted-foreground">Live company dataset</p>
            </div>
          </div>
        ) : null}

        <nav className="flex-1 overflow-y-auto px-3 py-4">
          {navigationGroups.map((group) => (
            <div key={group.label} className="mb-5">
              {!collapsed ? (
                <p className="mb-2 px-3 text-[11px] font-semibold uppercase tracking-[0.16em] text-muted-foreground">
                  {group.label}
                </p>
              ) : (
                <div className="mx-auto mb-2 h-px w-8 bg-border" />
              )}

              <div className="space-y-1">
                {group.items.map((item) => {
                  const Icon = item.icon;
                  const active =
                    item.href === "/dashboard"
                      ? pathname === "/dashboard"
                      : pathname === item.href || pathname.startsWith(`${item.href}/`);

                  const classes = [
                    "group relative flex items-center rounded-lg text-sm font-medium transition-colors",
                    collapsed ? "justify-center px-2 py-2.5" : "gap-3 px-3 py-2.5",
                    active
                      ? "bg-primary text-primary-foreground"
                      : "text-muted-foreground hover:bg-muted hover:text-foreground",
                    item.disabled ? "cursor-not-allowed opacity-45" : "",
                  ].join(" ");

                  const content = (
                    <>
                      <Icon className="size-4 shrink-0" />
                      {!collapsed ? <span className="flex-1">{item.label}</span> : null}
                      {!collapsed && item.disabled ? (
                        <span className="text-[10px] uppercase">Soon</span>
                      ) : null}
                      {collapsed ? (
                        <span className="pointer-events-none absolute left-[calc(100%+12px)] z-50 hidden whitespace-nowrap rounded-md bg-slate-950 px-3 py-2 text-xs font-semibold text-white shadow-lg group-hover:block">
                          {item.label}
                          {item.disabled ? " · Soon" : ""}
                        </span>
                      ) : null}
                    </>
                  );

                  return item.disabled ? (
                    <div key={`${group.label}-${item.label}-${item.href}`} className={classes} title="Coming soon">
                      {content}
                    </div>
                  ) : (
                    <Link key={`${group.label}-${item.label}-${item.href}`} href={item.href} className={classes}>
                      {content}
                    </Link>
                  );
                })}
              </div>
            </div>
          ))}
        </nav>

        <div className="border-t p-3">
          <Button
            type="button"
            variant="ghost"
            className={collapsed ? "w-full justify-center px-0" : "w-full justify-start"}
            onClick={handleLogout}
            title="Sign out"
          >
            <LogOut className="size-4" />
            {!collapsed ? "Sign out" : null}
          </Button>
        </div>
      </aside>

      <div className={collapsed ? "transition-all duration-300 lg:pl-20" : "transition-all duration-300 lg:pl-72"}>
        <header className="sticky top-0 z-20 flex h-16 items-center justify-between border-b bg-background/95 px-6 backdrop-blur">
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={toggleSidebar}
              className="hidden size-9 items-center justify-center rounded-lg border text-muted-foreground hover:bg-muted lg:flex"
              title={collapsed ? "Expand sidebar" : "Collapse sidebar"}
            >
              {collapsed ? <PanelLeftOpen className="size-4" /> : <PanelLeftClose className="size-4" />}
            </button>
            <div>
              <p className="text-sm font-medium">FinCruiz Workspace</p>
              <p className="text-xs text-muted-foreground">Financial intelligence and reporting</p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <ThemeToggle />
            <Button type="button" variant="outline" size="sm" onClick={handleLogout} className="lg:hidden">
              <LogOut className="size-4" />
              Sign out
            </Button>
          </div>
        </header>

        <main className="p-6 lg:p-8">{children}</main>
      </div>
      <AICFOFloating />
    </div>
  );
}
