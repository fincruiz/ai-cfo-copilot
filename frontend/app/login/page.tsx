"use client";

import { FormEvent, Suspense, useState } from "react";
import Link from "next/link";
import { useRouter, useSearchParams } from "next/navigation";
import {
  ArrowLeft,
  BarChart3,
  BrainCircuit,
  Loader2,
  LockKeyhole,
  PlayCircle,
  ShieldCheck,
  Sparkles,
} from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { getApiErrorMessage } from "@/lib/api";
import { authService } from "@/services/auth-service";

function LoginLoading() {
  return (
    <main className="flex min-h-screen items-center justify-center bg-[#f7f8fb]">
      <div className="flex items-center gap-3 text-sm text-muted-foreground">
        <Loader2 className="size-5 animate-spin text-primary" />
        Preparing secure sign in…
      </div>
    </main>
  );
}

function LoginContent() {
  const router = useRouter();
  const searchParams = useSearchParams();

  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [errorMessage, setErrorMessage] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [recoverySent, setRecoverySent] = useState(false);

  async function forgotPassword() {
    setErrorMessage("");
    setRecoverySent(false);

    if (!email.trim()) {
      setErrorMessage("Enter your email address first.");
      return;
    }

    setIsSubmitting(true);

    try {
      await authService.forgotPassword(email.trim());
      setRecoverySent(true);
    } catch (error: unknown) {
      setErrorMessage(getApiErrorMessage(error));
    } finally {
      setIsSubmitting(false);
    }
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setErrorMessage("");
    setRecoverySent(false);
    setIsSubmitting(true);

    try {
      await authService.login({
        email: email.trim(),
        password,
      });

      try {
        await authService.getCurrentCompany();
        router.replace("/dashboard");
      } catch (error: unknown) {
        const apiError = error as {
          response?: {
            status?: number;
            data?: {
              error_code?: string;
            };
          };
        };

        if (
          apiError.response?.status === 404 ||
          apiError.response?.data?.error_code ===
            "COMPANY_MEMBERSHIP_NOT_FOUND"
        ) {
          router.replace("/onboarding");
          return;
        }

        throw error;
      }
    } catch (error: unknown) {
      setErrorMessage(getApiErrorMessage(error));
    } finally {
      setIsSubmitting(false);
    }
  }

  const reason = searchParams.get("reason");

  return (
    <main className="relative grid min-h-screen overflow-hidden bg-[#f7f8fb] lg:grid-cols-[1.05fr_.95fr]">
      <div className="pointer-events-none fixed inset-0 landing-aurora" />

      <section className="relative hidden overflow-hidden bg-slate-950 p-12 text-white lg:flex lg:flex-col lg:justify-between">
        <div className="absolute inset-0 demo-aurora opacity-80" />

        <Link
          href="/"
          className="relative z-10 flex items-center gap-3"
        >
          <div className="flex size-10 items-center justify-center rounded-xl bg-white text-slate-950">
            <BarChart3 className="size-5" />
          </div>
          <span className="text-xl font-semibold">FinCruiz</span>
        </Link>

        <div className="relative z-10 max-w-xl">
          <p className="text-sm font-bold uppercase tracking-[.22em] text-indigo-300">
            Welcome back
          </p>

          <h1 className="mt-4 text-5xl font-black leading-tight tracking-[-.04em]">
            Your management briefing is waiting.
          </h1>

          <p className="mt-6 text-lg leading-8 text-slate-300">
            Return to financial intelligence, forecasts, working-capital
            signals, board reporting and the Organizational Brain.
          </p>

          <div className="mt-8 grid gap-3">
            {[
              [BrainCircuit, "Ask FinCruiz in plain English"],
              [
                ShieldCheck,
                "Finance evidence remains separate from AI narrative",
              ],
              [
                Sparkles,
                "Advanced features stay discoverable through Explore",
              ],
            ].map(([Icon, text]) => {
              const C = Icon as typeof BrainCircuit;

              return (
                <div
                  key={String(text)}
                  className="flex items-center gap-3 rounded-2xl border border-white/10 bg-white/[.05] p-3"
                >
                  <C className="size-4 text-indigo-300" />
                  <span className="text-sm text-slate-200">
                    {String(text)}
                  </span>
                </div>
              );
            })}
          </div>
        </div>

        <Link
          href="/demo"
          className="relative z-10 inline-flex w-fit items-center gap-2 text-sm font-semibold text-indigo-200 hover:text-white"
        >
          <PlayCircle className="size-4" />
          Not ready to connect data? Try the demo first.
        </Link>
      </section>

      <section className="relative z-10 flex items-center justify-center px-6 py-12">
        <div className="w-full max-w-md">
          <div className="mb-5 flex items-center justify-between lg:hidden">
            <Link
              href="/"
              className="flex items-center gap-2 text-sm font-semibold"
            >
              <ArrowLeft className="size-4" />
              FinCruiz
            </Link>

            <Link
              href="/demo"
              className="flex items-center gap-2 rounded-xl border bg-white px-3 py-2 text-sm font-semibold"
            >
              <PlayCircle className="size-4" />
              Demo
            </Link>
          </div>

          <Card className="border-white/70 bg-white/90 shadow-[0_30px_80px_rgba(15,23,42,.10)] backdrop-blur">
            <CardHeader className="space-y-3">
              <div className="flex size-11 items-center justify-center rounded-xl bg-primary text-primary-foreground">
                <LockKeyhole className="size-5" />
              </div>

              <CardTitle className="text-3xl tracking-tight">
                Welcome back
              </CardTitle>

              <CardDescription>
                Sign in to continue to your FinCruiz workspace.
              </CardDescription>
            </CardHeader>

            <CardContent>
              <form className="space-y-5" onSubmit={handleSubmit}>
                {reason === "email-confirmed" ? (
                  <Alert>
                    <AlertDescription>
                      Email confirmed. Sign in to continue to your FinCruiz
                      workspace.
                    </AlertDescription>
                  </Alert>
                ) : null}

                {reason === "session-expired" ? (
                  <Alert>
                    <AlertDescription>
                      Your session expired securely. Sign in again to continue.
                    </AlertDescription>
                  </Alert>
                ) : null}

                {reason === "inactivity" ? (
                  <Alert>
                    <AlertDescription>
                      You were signed out after a period of inactivity to protect your financial data.
                    </AlertDescription>
                  </Alert>
                ) : null}

                {reason === "session-limit" ? (
                  <Alert>
                    <AlertDescription>
                      Your secure session reached its maximum duration. Sign in again to continue.
                    </AlertDescription>
                  </Alert>
                ) : null}

                {reason === "signed-out" ? (
                  <Alert>
                    <AlertDescription>
                      You have been signed out securely across this browser session.
                    </AlertDescription>
                  </Alert>
                ) : null}

                {reason === "password-reset" ? (
                  <Alert>
                    <AlertDescription>
                      Password updated. Sign in with your new password.
                    </AlertDescription>
                  </Alert>
                ) : null}

                {recoverySent ? (
                  <Alert>
                    <AlertDescription>
                      If an account exists for that email, a password reset link
                      has been sent.
                    </AlertDescription>
                  </Alert>
                ) : null}

                {errorMessage ? (
                  <Alert variant="destructive">
                    <AlertDescription>{errorMessage}</AlertDescription>
                  </Alert>
                ) : null}

                <div className="space-y-2">
                  <Label htmlFor="email">Email address</Label>
                  <Input
                    id="email"
                    type="email"
                    placeholder="you@company.com"
                    value={email}
                    onChange={(event) => setEmail(event.target.value)}
                    autoComplete="email"
                    required
                    disabled={isSubmitting}
                  />
                </div>

                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <Label htmlFor="password">Password</Label>

                    <button
                      type="button"
                      onClick={forgotPassword}
                      disabled={isSubmitting}
                      className="text-sm font-medium text-primary hover:underline disabled:opacity-50"
                    >
                      Forgot password?
                    </button>
                  </div>

                  <Input
                    id="password"
                    type="password"
                    placeholder="Enter your password"
                    value={password}
                    onChange={(event) => setPassword(event.target.value)}
                    autoComplete="current-password"
                    required
                    disabled={isSubmitting}
                  />
                </div>

                <Button
                  type="submit"
                  className="w-full"
                  size="lg"
                  disabled={isSubmitting}
                >
                  {isSubmitting ? (
                    <>
                      <Loader2 className="animate-spin" />
                      Signing in...
                    </>
                  ) : (
                    <>
                      <LockKeyhole />
                      Sign in
                    </>
                  )}
                </Button>
              </form>

              <div className="mt-6 grid gap-3">
                <p className="text-center text-sm">
                  New to FinCruiz?{" "}
                  <Link
                    href="/signup"
                    className="font-semibold text-primary hover:underline"
                  >
                    Create your workspace
                  </Link>
                </p>

                <Link
                  href="/demo"
                  className="flex items-center justify-center gap-2 rounded-xl border bg-muted/30 px-4 py-3 text-sm font-semibold hover:bg-muted"
                >
                  <PlayCircle className="size-4" />
                  Explore with synthetic data first
                </Link>
              </div>

              <p className="mt-4 text-center text-xs leading-5 text-muted-foreground">
                By continuing, you agree to the FinCruiz terms of service and
                privacy policy.
              </p>
            </CardContent>
          </Card>
        </div>
      </section>
    </main>
  );
}

export default function LoginPage() {
  return (
    <Suspense fallback={<LoginLoading />}>
      <LoginContent />
    </Suspense>
  );
}