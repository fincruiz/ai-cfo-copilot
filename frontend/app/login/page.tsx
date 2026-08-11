"use client";

import { FormEvent, useState } from "react";
import { useRouter } from "next/navigation";
import {
  BarChart3,
  Loader2,
  LockKeyhole,
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
import {
  getApiErrorMessage,
} from "@/lib/api";
import {
  authService,
} from "@/services/auth-service";

export default function LoginPage() {
  const router = useRouter();

  const [email, setEmail] = useState("");
  const [password, setPassword] =
    useState("");
  const [errorMessage, setErrorMessage] =
    useState("");
  const [isSubmitting, setIsSubmitting] =
    useState(false);

  async function handleSubmit(
    event: FormEvent<HTMLFormElement>,
  ) {
    event.preventDefault();

    setErrorMessage("");
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
        const apiError =
          error as {
            response?: {
              status?: number;
              data?: {
                error_code?: string;
              };
            };
          };

        const status =
          apiError.response?.status;

        const errorCode =
          apiError.response?.data?.error_code;

        if (
          status === 404 ||
          errorCode ===
            "COMPANY_MEMBERSHIP_NOT_FOUND"
        ) {
          router.replace("/onboarding");
          return;
        }

        throw error;
      }
    } catch (error: unknown) {
      setErrorMessage(
        getApiErrorMessage(error),
      );
    } finally {
      setIsSubmitting(false);
    }
  }

  return (
    <main className="grid min-h-screen lg:grid-cols-2">
      <section className="hidden bg-slate-950 p-12 text-white lg:flex lg:flex-col lg:justify-between">
        <div className="flex items-center gap-3">
          <div className="flex size-10 items-center justify-center rounded-xl bg-white text-slate-950">
            <BarChart3 className="size-5" />
          </div>

          <span className="text-xl font-semibold tracking-tight">
            FinCruiz
          </span>
        </div>

        <div className="max-w-xl">
          <p className="mb-4 text-sm font-medium uppercase tracking-[0.25em] text-slate-400">
            AI-powered finance intelligence
          </p>

          <h1 className="text-5xl font-semibold leading-tight tracking-tight">
            Turn financial data into clear
            business decisions.
          </h1>

          <p className="mt-6 max-w-lg text-lg leading-8 text-slate-300">
            Reporting, forecasting, KPIs and
            AI CFO insights in one secure
            workspace.
          </p>
        </div>

        <p className="text-sm text-slate-500">
          Financial clarity for growing
          businesses.
        </p>
      </section>

      <section className="flex items-center justify-center bg-background px-6 py-12">
        <Card className="w-full max-w-md border-border/70 shadow-xl shadow-black/5">
          <CardHeader className="space-y-3">
            <div className="flex size-11 items-center justify-center rounded-xl bg-primary text-primary-foreground lg:hidden">
              <BarChart3 className="size-5" />
            </div>

            <CardTitle className="text-3xl tracking-tight">
              Welcome back
            </CardTitle>

            <CardDescription>
              Sign in to access your FinCruiz
              workspace.
            </CardDescription>
          </CardHeader>

          <CardContent>
            <form
              className="space-y-5"
              onSubmit={handleSubmit}
            >
              {errorMessage ? (
                <Alert variant="destructive">
                  <AlertDescription>
                    {errorMessage}
                  </AlertDescription>
                </Alert>
              ) : null}

              <div className="space-y-2">
                <Label htmlFor="email">
                  Email address
                </Label>

                <Input
                  id="email"
                  type="email"
                  placeholder="you@company.com"
                  value={email}
                  onChange={(event) =>
                    setEmail(
                      event.target.value,
                    )
                  }
                  autoComplete="email"
                  required
                  disabled={isSubmitting}
                />
              </div>

              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <Label htmlFor="password">
                    Password
                  </Label>

                  <button
                    type="button"
                    className="text-sm font-medium text-primary hover:underline"
                  >
                    Forgot password?
                  </button>
                </div>

                <Input
                  id="password"
                  type="password"
                  placeholder="Enter your password"
                  value={password}
                  onChange={(event) =>
                    setPassword(
                      event.target.value,
                    )
                  }
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

            <div className="mt-6 text-center text-sm">
              New to FinCruiz?{" "}
              <a href="/signup" className="font-medium text-primary hover:underline">
                Create your FinCruiz workspace
              </a>
            </div>

            <p className="mt-4 text-center text-xs leading-5 text-muted-foreground">
              By continuing, you agree to the
              FinCruiz terms of service and
              privacy policy.
            </p>
          </CardContent>
        </Card>
      </section>
    </main>
  );
}