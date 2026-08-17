"use client";

import { useEffect, useState } from "react";
import { Database, FlaskConical, Loader2, Save, ShieldCheck, Trash2, UserX } from "lucide-react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { ConfirmDialog } from "@/components/ui/confirm-dialog";
import { getApiErrorMessage } from "@/lib/api";
import { authService } from "@/services/auth-service";
import { companyService } from "@/services/company-service";
import { workspaceService, type WorkspaceStatus } from "@/services/workspace-service";
import { SubscriptionHealthCard } from "@/components/subscription-health-card";
import { ProductAdoptionCard } from "@/components/product-adoption-card";

export default function SettingsPage() {
  const [settings, setSettings] = useState<any>(null);
  const [workspace, setWorkspace] = useState<WorkspaceStatus | null>(null);
  const [saving, setSaving] = useState(false);
  const [working, setWorking] = useState<"demo" | "reset" | "delete" | null>(null);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [confirmAction, setConfirmAction] = useState<"reset" | "delete" | null>(null);

  useEffect(() => {
    Promise.all([companyService.getPreferences(), workspaceService.getStatus()])
      .then(([preferences, status]) => {
        setSettings(preferences);
        setWorkspace(status);
      })
      .catch((loadError) => setError(getApiErrorMessage(loadError)));
  }, []);

  async function refreshWorkspace() {
    setWorkspace(await workspaceService.getStatus());
  }

  async function save() {
    setSaving(true);
    setError("");
    try {
      setSettings(await companyService.updatePreferences(settings));
      setMessage("Workspace settings saved.");
    } catch (saveError) {
      setError(getApiErrorMessage(saveError));
    } finally {
      setSaving(false);
    }
  }

  async function loadDemo() {
    setWorking("demo");
    setError("");
    setMessage("");
    try {
      await workspaceService.loadDemo(false);
      await refreshWorkspace();
      setMessage("Demo data loaded. You can now explore reports, KPIs, forecasting and AI CFO features.");
    } catch (demoError) {
      setError(getApiErrorMessage(demoError));
    } finally {
      setWorking(null);
    }
  }

  async function resetData() {
    setWorking("reset");
    setError("");
    setMessage("");
    try {
      await workspaceService.resetData();
      setConfirmAction(null);
      await refreshWorkspace();
      setMessage("All loaded financial data has been permanently removed. Your company profile and settings are still available.");
    } catch (resetError) {
      setError(getApiErrorMessage(resetError));
    } finally {
      setWorking(null);
    }
  }

  async function deleteProfile() {
    setWorking("delete");
    setError("");
    try {
      await workspaceService.deleteAccount();
      setConfirmAction(null);
      authService.logout();
      window.location.replace("/?account=deleted");
    } catch (deleteError) {
      setError(getApiErrorMessage(deleteError));
      setWorking(null);
    }
  }

  if (!settings || !workspace) {
    return <div className="flex min-h-[400px] items-center justify-center"><Loader2 className="size-5 animate-spin" /></div>;
  }

  return (
    <div className="mx-auto max-w-4xl space-y-6">
      <div>
        <p className="text-sm text-muted-foreground">Administration</p>
        <h1 className="text-3xl font-semibold">Workspace Settings</h1>
        <p className="mt-2 text-sm text-muted-foreground">Control reporting defaults, demo data and your privacy choices.</p>
      </div>

      {message ? <div className="rounded-xl border border-emerald-200 bg-emerald-50 p-4 text-sm text-emerald-900">{message}</div> : null}
      {error ? <div className="rounded-xl border border-destructive/30 bg-destructive/5 p-4 text-sm text-destructive">{error}</div> : null}

      <SubscriptionHealthCard />
      <ProductAdoptionCard />

      <Card>
        <CardHeader>
          <CardTitle>Reporting preferences</CardTitle>
          <CardDescription>Control defaults across reports and analytics.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-5">
          {[
            ["theme_preference", "Theme", ["system", "light", "dark"]],
            ["reporting_frequency", "Reporting frequency", ["monthly", "quarterly", "annual"]],
            ["default_report_view", "Default report view", ["consolidated", "branch"]],
            ["number_format", "Number format", ["international", "indian"]],
          ].map(([key, label, options]: any) => (
            <div key={key} className="grid gap-2 sm:grid-cols-[220px_1fr] sm:items-center">
              <label>{label}</label>
              <select className="h-10 rounded-md border bg-background px-3" value={settings[key]} onChange={(event) => setSettings({ ...settings, [key]: event.target.value })}>
                {options.map((option: string) => <option key={option}>{option}</option>)}
              </select>
            </div>
          ))}
          <div className="grid gap-2 sm:grid-cols-[220px_1fr] sm:items-center">
            <label>Variance warning %</label>
            <Input type="number" value={settings.variance_warning_percent} onChange={(event) => setSettings({ ...settings, variance_warning_percent: Number(event.target.value) })} />
          </div>
          {["show_ai_assistant", "email_notifications"].map((key) => (
            <label key={key} className="flex items-center gap-3">
              <input type="checkbox" checked={Boolean(settings[key])} onChange={(event) => setSettings({ ...settings, [key]: event.target.checked })} />
              {key === "show_ai_assistant" ? "Show AI CFO assistant" : "Email notifications"}
            </label>
          ))}
          <Button onClick={() => void save()} disabled={saving}>
            {saving ? <Loader2 className="size-4 animate-spin" /> : <Save className="size-4" />}Save settings
          </Button>
        </CardContent>
      </Card>

      <Card className="border-sky-200">
        <CardHeader>
          <div className="flex items-start gap-3">
            <div className="rounded-lg bg-sky-100 p-2 text-sky-700"><FlaskConical className="size-5" /></div>
            <div>
              <CardTitle>Explore with demo data</CardTitle>
              <CardDescription className="mt-1">Load a synthetic 12-month ledger so you can experience FinCruiz without uploading real company information.</CardDescription>
            </div>
          </div>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="grid gap-3 rounded-xl bg-muted/40 p-4 sm:grid-cols-3">
            <div><p className="text-xs text-muted-foreground">Financial data</p><p className="font-semibold">{workspace.has_financial_data ? "Loaded" : "Empty"}</p></div>
            <div><p className="text-xs text-muted-foreground">Demo mode</p><p className="font-semibold">{workspace.demo_data_active ? "Active" : "Off"}</p></div>
            <div><p className="text-xs text-muted-foreground">Transactions</p><p className="font-semibold">{workspace.transaction_count.toLocaleString()}</p></div>
          </div>
          <Button variant="outline" onClick={() => void loadDemo()} disabled={working !== null || workspace.has_financial_data}>
            {working === "demo" ? <Loader2 className="size-4 animate-spin" /> : <Database className="size-4" />}
            {workspace.demo_data_active ? "Demo data active" : workspace.has_financial_data ? "Reset current data first" : "Load demo workspace"}
          </Button>
          <p className="text-xs text-muted-foreground">Demo figures are synthetic and are clearly marked as demo data. Loading demo data is disabled when your workspace already contains financial information.</p>
        </CardContent>
      </Card>

      <Card className="border-amber-200">
        <CardHeader>
          <div className="flex items-start gap-3">
            <div className="rounded-lg bg-amber-100 p-2 text-amber-800"><ShieldCheck className="size-5" /></div>
            <div>
              <CardTitle>Reset loaded financial data</CardTitle>
              <CardDescription className="mt-1">Permanently remove uploaded ledgers, mappings, imports, plans, forecasts and generated board-pack records. Your company profile, preferences and login remain intact.</CardDescription>
            </div>
          </div>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="rounded-xl border bg-muted/20 p-4 text-sm">
            <p><span className="font-medium">Currently stored:</span> {workspace.upload_count} upload(s), {workspace.transaction_count.toLocaleString()} ledger transaction(s), {workspace.mapping_count} mapping(s).</p>
          </div>
          <Button variant="outline" onClick={() => setConfirmAction("reset")} disabled={working !== null || !workspace.has_financial_data}>
            <Trash2 className="size-4" />Reset loaded data
          </Button>
        </CardContent>
      </Card>

      <Card className="border-destructive/40">
        <CardHeader>
          <div className="flex items-start gap-3">
            <div className="rounded-lg bg-destructive/10 p-2 text-destructive"><UserX className="size-5" /></div>
            <div>
              <CardTitle>Delete profile</CardTitle>
              <CardDescription className="mt-1">Permanently delete your FinCruiz profile, login identity and any single-user workspace you own. This cannot be undone.</CardDescription>
            </div>
          </div>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="rounded-xl border border-destructive/20 bg-destructive/5 p-4 text-sm text-muted-foreground">
            For safety, an owner cannot delete their profile while their workspace still has other active members. Ownership must be transferred first.
          </div>
          <Button variant="destructive" onClick={() => setConfirmAction("delete")} disabled={working !== null}>
            {working === "delete" ? <Loader2 className="size-4 animate-spin" /> : <UserX className="size-4" />}Permanently delete profile
          </Button>
        </CardContent>
      </Card>
      <ConfirmDialog
        open={confirmAction === "reset"}
        title="Reset all loaded financial data?"
        description="This permanently removes uploaded ledgers, mappings, imports, plans, forecasts and generated finance records. Your login, company profile and preferences will stay intact."
        confirmLabel="Yes, reset my data"
        onCancel={() => setConfirmAction(null)}
        onConfirm={() => void resetData()}
        loading={working === "reset"}
        destructive
      />
      <ConfirmDialog
        open={confirmAction === "delete"}
        title="Permanently delete your profile?"
        description="This removes your FinCruiz login and your single-user workspace. This cannot be undone. Shared-workspace owners must transfer ownership first."
        confirmLabel="Yes, delete my profile"
        onCancel={() => setConfirmAction(null)}
        onConfirm={() => void deleteProfile()}
        loading={working === "delete"}
        destructive
      />
    </div>
  );
}
