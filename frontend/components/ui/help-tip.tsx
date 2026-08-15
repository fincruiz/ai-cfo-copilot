"use client";

import { Info } from "lucide-react";
import { useCallback, useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";

type HelpSide = "right" | "left" | "top" | "bottom";

export function HelpTip({ text, title, side = "right" }: { text: string; title?: string; side?: HelpSide }) {
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const [open, setOpen] = useState(false);
  const [mounted, setMounted] = useState(false);
  const [position, setPosition] = useState({ left: 16, top: 16 });

  useEffect(() => setMounted(true), []);

  const place = useCallback(() => {
    const trigger = triggerRef.current;
    if (!trigger) return;
    const rect = trigger.getBoundingClientRect();
    const width = Math.min(320, Math.max(240, window.innerWidth - 32));
    const estimatedHeight = title ? 120 : 96;
    const gap = 10;
    let left = rect.right + gap;
    let top = rect.top + rect.height / 2 - estimatedHeight / 2;

    if (side === "left") left = rect.left - width - gap;
    if (side === "top") {
      left = rect.left + rect.width / 2 - width / 2;
      top = rect.top - estimatedHeight - gap;
    }
    if (side === "bottom") {
      left = rect.left + rect.width / 2 - width / 2;
      top = rect.bottom + gap;
    }

    if (left + width > window.innerWidth - 12) left = window.innerWidth - width - 12;
    if (left < 12) left = 12;
    if (top + estimatedHeight > window.innerHeight - 12) top = window.innerHeight - estimatedHeight - 12;
    if (top < 12) top = 12;
    setPosition({ left, top });
  }, [side, title]);

  useEffect(() => {
    if (!open) return;
    place();
    const update = () => place();
    window.addEventListener("resize", update);
    window.addEventListener("scroll", update, true);
    return () => {
      window.removeEventListener("resize", update);
      window.removeEventListener("scroll", update, true);
    };
  }, [open, place]);

  const bubble = mounted && open ? createPortal(
    <div
      role="tooltip"
      className="pointer-events-none fixed z-[9999] w-[min(320px,calc(100vw-32px))] rounded-2xl border border-border/80 bg-popover px-4 py-3 text-left shadow-[0_18px_55px_rgba(15,23,42,.22)] backdrop-blur-xl"
      style={{ left: position.left, top: position.top }}
    >
      {title ? <p className="mb-1 text-sm font-semibold text-popover-foreground">{title}</p> : null}
      <p className="text-xs font-normal leading-5 text-muted-foreground">{text}</p>
    </div>,
    document.body,
  ) : null;

  return (
    <span className="inline-flex shrink-0">
      <button
        ref={triggerRef}
        type="button"
        aria-label={title ? `About ${title}` : "More information"}
        aria-expanded={open}
        className="flex size-5 items-center justify-center rounded-full text-muted-foreground/70 transition hover:bg-muted hover:text-foreground focus:outline-none focus:ring-2 focus:ring-ring"
        onMouseEnter={() => { place(); setOpen(true); }}
        onMouseLeave={() => setOpen(false)}
        onFocus={() => { place(); setOpen(true); }}
        onBlur={() => setOpen(false)}
        onClick={(event) => { event.preventDefault(); event.stopPropagation(); place(); setOpen((value) => !value); }}
      >
        <Info className="size-3.5" />
      </button>
      {bubble}
    </span>
  );
}
