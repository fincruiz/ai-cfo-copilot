"use client";

import { useEffect, useState } from "react";
import { Moon, Sun } from "lucide-react";

export function ThemeToggle() {
  const [dark, setDark] = useState(false);

  useEffect(() => {
    const saved = window.localStorage.getItem("fincruiz_theme");
    const shouldUseDark =
      saved === "dark" ||
      (!saved && window.matchMedia("(prefers-color-scheme: dark)").matches);
    document.documentElement.classList.toggle("dark", shouldUseDark);
    setDark(shouldUseDark);
  }, []);

  function toggle() {
    setDark((current) => {
      const next = !current;
      document.documentElement.classList.toggle("dark", next);
      window.localStorage.setItem("fincruiz_theme", next ? "dark" : "light");
      return next;
    });
  }

  return (
    <button
      type="button"
      onClick={toggle}
      className="flex size-10 items-center justify-center rounded-xl border bg-background shadow-sm transition hover:-translate-y-0.5 hover:bg-muted hover:shadow-md"
      title={dark ? "Use light mode" : "Use dark mode"}
    >
      {dark ? <Sun className="size-4" /> : <Moon className="size-4" />}
    </button>
  );
}
