"use client";

import { AlertTriangle, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ViewportModal } from "@/components/ui/viewport-modal";

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
  return (
    <ViewportModal
      open={open}
      onClose={() => {
        if (!loading) onCancel();
      }}
      title={title}
      description={description}
      maxWidthClass="max-w-md"
      footer={
        <div className="flex flex-col-reverse gap-2 sm:flex-row sm:justify-end">
          <Button variant="outline" onClick={onCancel} disabled={loading}>Cancel</Button>
          <Button variant={destructive ? "destructive" : "default"} onClick={onConfirm} disabled={loading}>
            {loading ? <Loader2 className="size-4 animate-spin" /> : null}
            {confirmLabel}
          </Button>
        </div>
      }
    >
      <div className="flex items-start gap-4 rounded-2xl border bg-muted/30 p-4">
        <div className={[
          "flex size-11 shrink-0 items-center justify-center rounded-2xl",
          destructive
            ? "bg-destructive/10 text-destructive"
            : "bg-amber-100 text-amber-800 dark:bg-amber-950/30 dark:text-amber-300",
        ].join(" ")}>
          <AlertTriangle className="size-5" />
        </div>
        <p className="pt-1 text-sm leading-6 text-muted-foreground">
          Review this action carefully before continuing.
        </p>
      </div>
    </ViewportModal>
  );
}
