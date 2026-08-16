"use client";

import { useEffect, type MouseEvent, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { X } from "lucide-react";

export function ViewportModal({
  open,
  onClose,
  title,
  description,
  children,
  footer,
  maxWidthClass = "max-w-2xl",
}: {
  open: boolean;
  onClose: () => void;
  title: string;
  description?: string;
  children: ReactNode;
  footer?: ReactNode;
  maxWidthClass?: string;
}) {
  useEffect(() => {
    if (!open) return;
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKeyDown);
    return () => {
      document.body.style.overflow = previousOverflow;
      window.removeEventListener("keydown", onKeyDown);
    };
  }, [open, onClose]);

  if (!open || typeof document === "undefined") return null;

  const dismissBackdrop = (event: MouseEvent<HTMLDivElement>) => {
    if (event.target === event.currentTarget) onClose();
  };

  return createPortal(
    <div
      className="fixed inset-0 z-[300] flex items-center justify-center overflow-y-auto bg-slate-950/50 p-4 backdrop-blur-sm sm:p-6"
      onMouseDown={dismissBackdrop}
      role="presentation"
    >
      <section
        role="dialog"
        aria-modal="true"
        aria-labelledby="viewport-modal-title"
        className={`my-auto flex max-h-[min(88vh,900px)] w-full ${maxWidthClass} flex-col overflow-hidden rounded-[28px] border bg-background shadow-2xl`}
      >
        <header className="flex shrink-0 items-start justify-between gap-4 border-b px-5 py-5 sm:px-6">
          <div className="min-w-0">
            <h2 id="viewport-modal-title" className="text-xl font-semibold tracking-tight">{title}</h2>
            {description ? <p className="mt-1 text-sm leading-6 text-muted-foreground">{description}</p> : null}
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close dialog"
            className="flex size-9 shrink-0 items-center justify-center rounded-xl border bg-background transition hover:bg-muted"
          >
            <X className="size-4" />
          </button>
        </header>
        <div className="min-h-0 flex-1 overflow-y-auto overscroll-contain px-5 py-5 sm:px-6">{children}</div>
        {footer ? <footer className="shrink-0 border-t bg-background px-5 py-4 sm:px-6">{footer}</footer> : null}
      </section>
    </div>,
    document.body,
  );
}
