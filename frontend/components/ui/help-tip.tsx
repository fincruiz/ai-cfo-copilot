"use client";

import { Info } from "lucide-react";

export function HelpTip({ text, side = "right" }: { text: string; side?: "right" | "left" | "top" }) {
  const position = side === "left"
    ? "right-[calc(100%+10px)] top-1/2 -translate-y-1/2"
    : side === "top"
      ? "bottom-[calc(100%+10px)] left-1/2 -translate-x-1/2"
      : "left-[calc(100%+10px)] top-1/2 -translate-y-1/2";

  return (
    <span className="group/help relative inline-flex shrink-0">
      <button
        type="button"
        aria-label="More information"
        className="flex size-5 items-center justify-center rounded-full text-muted-foreground/70 transition hover:bg-muted hover:text-foreground focus:outline-none focus:ring-2 focus:ring-ring"
        onClick={(event) => event.preventDefault()}
      >
        <Info className="size-3.5" />
      </button>
      <span className={`pointer-events-none absolute z-[100] hidden w-64 rounded-xl border bg-popover px-3 py-2 text-left text-xs font-normal leading-5 text-popover-foreground shadow-xl group-hover/help:block group-focus-within/help:block ${position}`}>
        {text}
      </span>
    </span>
  );
}
