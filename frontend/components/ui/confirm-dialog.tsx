"use client";

import { AlertTriangle, Loader2, X } from "lucide-react";
import { Button } from "@/components/ui/button";

interface ConfirmDialogProps {
  open: boolean;
  title: string;
  description: string;
  confirmLabel: string;
  onCancel: () => void;
  onConfirm: () => void;
  loading?: boolean;
  destructive?: boolean;
}

export function ConfirmDialog({
  open,
  title,
  description,
  confirmLabel,
  onCancel,
  onConfirm,
  loading = false,
  destructive = false,
}: ConfirmDialogProps) {
  if (!open) return null;

  return (
    <div className="fixed inset-0 z-[120] flex items-center justify-center bg-slate-950/55 p-4 backdrop-blur-sm">
      <div className="w-full max-w-md animate-card-in rounded-3xl border bg-background p-6 shadow-2xl">
        <div className="flex items-start gap-4">
          <div className={[
            "flex size-11 shrink-0 items-center justify-center rounded-2xl",
            destructive ? "bg-destructive/10 text-destructive" : "bg-amber-100 text-amber-800 dark:bg-amber-950/30 dark:text-amber-300",
          ].join(" ")}>
            <AlertTriangle className="size-5" />
          </div>
          <div className="min-w-0 flex-1">
            <h2 className="text-xl font-semibold">{title}</h2>
            <p className="mt-2 text-sm leading-6 text-muted-foreground">{description}</p>
          </div>
          <button type="button" onClick={onCancel} disabled={loading} className="rounded-xl p-2 text-muted-foreground transition hover:bg-muted">
            <X className="size-4" />
          </button>
        </div>

        <div className="mt-6 flex flex-col-reverse gap-2 sm:flex-row sm:justify-end">
          <Button variant="outline" onClick={onCancel} disabled={loading}>Cancel</Button>
          <Button variant={destructive ? "destructive" : "default"} onClick={onConfirm} disabled={loading}>
            {loading ? <Loader2 className="size-4 animate-spin" /> : null}
            {confirmLabel}
          </Button>
        </div>
      </div>
    </div>
  );
}
