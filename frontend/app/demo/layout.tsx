import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Interactive AI CFO Demo",
  description:
    "Explore FinCruiz with a guided synthetic-company demo. Ask management questions, inspect evidence, compare branches and test decisions without using customer data.",
};

export default function DemoLayout({ children }: { children: React.ReactNode }) {
  return children;
}
