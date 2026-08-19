import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: {
    default: "FinCruiz | AI CFO & Management Intelligence",
    template: "%s | FinCruiz",
  },
  description:
    "Turn finance data into evidence-backed management decisions with conversational BI, branch intelligence, forecasting, scenario modelling and board reporting.",
  keywords: [
    "AI CFO",
    "management reporting",
    "financial forecasting",
    "conversational BI",
    "branch reporting",
    "scenario modelling",
    "board reporting",
  ],
  robots: {
    index: true,
    follow: true,
  },
  openGraph: {
    title: "FinCruiz | AI CFO & Management Intelligence",
    description:
      "Ask the business, see the evidence and model the next management decision.",
    type: "website",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html
      lang="en"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
    >
      <body className="flex min-h-full flex-col">{children}</body>
    </html>
  );
}
