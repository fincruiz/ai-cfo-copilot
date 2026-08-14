"use client";

import { useEffect, useState } from "react";
import { Building2, Cloud, RefreshCw, ShieldCheck, Unplug, Zap } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ConfirmDialog } from "@/components/ui/confirm-dialog";
import { getApiErrorMessage } from "@/lib/api";
import { integrationService } from "@/services/integration-service";
import type { IntegrationConnection, Provider } from "@/types/integrations";

const meta: Record<Provider, { name: string; subtitle: string; accent: string }> = {
  xero: { name: "Xero", subtitle: "Accounting, contacts, invoices and banking", accent: "bg-sky-500/10" },
  zoho: { name: "Zoho Books", subtitle: "Accounts, customers, invoices and bills", accent: "bg-orange-500/10" },
  tally: { name: "TallyPrime", subtitle: "Secure bridge from on-premise Tally data", accent: "bg-emerald-500/10" },
};

export default function IntegrationsPage() {
  const [items, setItems] = useState<IntegrationConnection[]>([]);
  const [busy, setBusy] = useState("");
  const [error, setError] = useState("");
  const [tallyToken, setTallyToken] = useState("");
  const [remove, setRemove] = useState<Provider | null>(null);
  const [tenantChoice, setTenantChoice] = useState<Record<string, string>>({});

  const load = async () => {
    try {
      setError("");
      setItems(await integrationService.list());
    } catch (e) {
      setError(getApiErrorMessage(e));
    }
  };

  useEffect(() => { void load(); }, []);

  const connectXero = async () => {
    try {
      setBusy("xero"); setError("");
      const response = await integrationService.start("xero");
      window.location.assign(response.authorization_url);
    } catch (e) {
      setError(getApiErrorMessage(e)); setBusy("");
    }
  };

  const connectZoho = async () => {
    try {
      setBusy("zoho"); setError("");
      const response = await integrationService.start("zoho");
      window.location.assign(response.authorization_url);
    } catch (e) {
      setError(getApiErrorMessage(e)); setBusy("");
    }
  };

  const sync = async (provider: "xero" | "zoho") => {
    try {
      setBusy(provider); setError("");
      await integrationService.sync(provider);
      await load();
    } catch (e) {
      setError(getApiErrorMessage(e));
    } finally {
      setBusy("");
    }
  };

  const createTallyBridge = async () => {
    try {
      setBusy("tally"); setError("");
      const response = await integrationService.createTallyToken();
      setTallyToken(response.bridge_token);
      await load();
    } catch (e) {
      setError(getApiErrorMessage(e));
    } finally {
      setBusy("");
    }
  };

  const chooseTenant = async (provider: "xero" | "zoho") => {
    const tenant = tenantChoice[provider];
    if (!tenant) return;
    try {
      setBusy(provider); setError("");
      await integrationService.selectTenant(provider, tenant);
      await load();
    } catch (e) {
      setError(getApiErrorMessage(e));
    } finally {
      setBusy("");
    }
  };

  const disconnectIntegration = async () => {
    if (!remove) return;
    const provider = remove;
    try {
      setBusy(provider); setError("");
      await integrationService.disconnect(provider);
      setRemove(null);
      await load();
    } catch (e) {
      setError(getApiErrorMessage(e));
    } finally {
      setBusy("");
    }
  };

  return (
    <div className="mx-auto max-w-7xl space-y-8 p-6 lg:p-10">
      <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
        <div>
          <div className="mb-3 inline-flex items-center gap-2 rounded-full border bg-background px-3 py-1 text-xs font-medium"><Zap className="size-3.5" />Organizational Brain · Data connections</div>
          <h1 className="text-3xl font-semibold tracking-tight">Connect the systems that run your business</h1>
          <p className="mt-2 max-w-3xl text-muted-foreground">FinCruiz brings finance and operational systems into one governed intelligence layer so management can ask one question across the company instead of reconciling separate dashboards.</p>
        </div>
        <div className="rounded-2xl border bg-background p-4 text-sm"><p className="font-medium">Privacy first</p><p className="mt-1 text-muted-foreground">Disconnecting a source can also remove its synchronized FinCruiz copy.</p></div>
      </div>

      {error ? <div className="rounded-xl border border-destructive/30 bg-destructive/5 p-4 text-sm text-destructive">{error}</div> : null}

      <div className="grid gap-5 lg:grid-cols-3">
        {(["xero", "zoho", "tally"] as Provider[]).map((provider) => {
          const item = items.find((connection) => connection.provider === provider);
          const connected = item?.status === "connected";
          const isBusy = busy === provider;
          const tenantOptions = provider === "xero"
            ? ((item?.metadata?.tenants as any[] | undefined) ?? [])
            : ((item?.metadata?.organizations as any[] | undefined) ?? []);

          return (
            <div key={provider} className="rounded-3xl border bg-background p-6 shadow-sm">
              <div className={`flex size-12 items-center justify-center rounded-2xl ${meta[provider].accent}`}>
                {provider === "tally" ? <Building2 className="size-5" /> : <Cloud className="size-5" />}
              </div>
              <div className="mt-5 flex items-center justify-between gap-3">
                <h2 className="text-xl font-semibold">{meta[provider].name}</h2>
                <span className={`rounded-full px-2.5 py-1 text-xs font-medium ${connected ? "bg-emerald-500/10 text-emerald-700" : "bg-muted text-muted-foreground"}`}>{item?.status?.replaceAll("_", " ") || "disconnected"}</span>
              </div>
              <p className="mt-2 text-sm text-muted-foreground">{meta[provider].subtitle}</p>

              {item?.external_tenant_name ? <div className="mt-4 rounded-xl bg-muted/50 p-3 text-sm"><span className="text-muted-foreground">Connected organisation</span><br /><strong>{item.external_tenant_name}</strong></div> : null}

              {item?.status === "selection_required" ? (
                <div className="mt-4 rounded-2xl border bg-muted/30 p-4">
                  <p className="text-sm font-medium">Choose the organisation FinCruiz should use</p>
                  <select className="mt-3 w-full rounded-xl border bg-background px-3 py-2 text-sm" value={tenantChoice[provider] || ""} onChange={(event) => setTenantChoice((current) => ({ ...current, [provider]: event.target.value }))}>
                    <option value="">Select organisation…</option>
                    {tenantOptions.map((organisation: any) => {
                      const id = String(organisation.tenantId || organisation.organization_id);
                      const name = organisation.tenantName || organisation.name || organisation.organisationName || "Organisation";
                      return <option key={id} value={id}>{name}</option>;
                    })}
                  </select>
                  <Button className="mt-3" size="sm" disabled={!tenantChoice[provider] || isBusy} onClick={() => void chooseTenant(provider as "xero" | "zoho")}>Use this organisation</Button>
                </div>
              ) : null}

              <div className="mt-6 flex flex-wrap gap-2">
                {provider === "xero" && !connected && item?.status !== "selection_required" ? (
                  <Button type="button" disabled={busy === "xero" || item?.configured === false} onClick={(event) => { event.preventDefault(); event.stopPropagation(); void connectXero(); }}>
                    {item?.configured === false ? "Configure server" : busy === "xero" ? "Connecting..." : "Connect Xero"}
                  </Button>
                ) : null}
                {provider === "zoho" && !connected && item?.status !== "selection_required" ? (
                  <Button type="button" disabled={busy === "zoho" || item?.configured === false} onClick={(event) => { event.preventDefault(); event.stopPropagation(); void connectZoho(); }}>
                    {item?.configured === false ? "Configure server" : busy === "zoho" ? "Connecting..." : "Connect Zoho Books"}
                  </Button>
                ) : null}
                {provider !== "tally" && connected ? <Button type="button" onClick={() => void sync(provider as "xero" | "zoho")} disabled={isBusy}><RefreshCw className={`mr-2 size-4 ${isBusy ? "animate-spin" : ""}`} />{isBusy ? "Syncing…" : "Sync now"}</Button> : null}
                {provider === "tally" && !connected ? <Button type="button" onClick={() => void createTallyBridge()} disabled={isBusy}>{isBusy ? "Creating..." : "Create secure bridge"}</Button> : null}
                {item && item.status !== "disconnected" ? <Button type="button" variant="outline" disabled={isBusy} onClick={() => setRemove(provider)}><Unplug className="mr-2 size-4" />Disconnect</Button> : null}
              </div>
              {item?.last_synced_at ? <p className="mt-4 text-xs text-muted-foreground">Last sync · {new Date(item.last_synced_at).toLocaleString()}</p> : null}
            </div>
          );
        })}
      </div>

      {tallyToken ? <div className="rounded-3xl border bg-background p-6"><div className="flex items-start gap-3"><ShieldCheck className="mt-0.5 size-5" /><div><h2 className="font-semibold">Tally bridge token — shown once</h2><p className="mt-1 text-sm text-muted-foreground">Use this token in the FinCruiz Tally bridge running on the same network as TallyPrime. Do not email or store it in a spreadsheet.</p><code className="mt-4 block break-all rounded-xl bg-muted p-4 text-xs">{tallyToken}</code></div></div></div> : null}

      <ConfirmDialog
        open={!!remove}
        title={`Disconnect ${remove ? meta[remove].name : "integration"}?`}
        description="This removes the connection and synchronized copy from FinCruiz. It does not delete anything from the source system."
        confirmLabel="Disconnect & delete synced data"
        destructive
        onCancel={() => setRemove(null)}
        onConfirm={() => void disconnectIntegration()}
        loading={!!remove && busy === remove}
      />
    </div>
  );
}
