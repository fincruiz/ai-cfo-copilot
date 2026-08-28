"use client";

import { Suspense, useEffect, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import { CheckCircle2, Loader2, TriangleAlert } from "lucide-react";

import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { authService } from "@/services/auth-service";

function safeNext(value: string | null): string {
  if (!value || !value.startsWith("/") || value.startsWith("//")) {
    return "/onboarding";
  }
  return value;
}

function CallbackLoadingCard() {
  return (
    <main className="flex min-h-screen items-center justify-center bg-muted/30 px-6">
      <Card className="w-full max-w-lg">
        <CardHeader>
          <Loader2 className="mb-3 size-10 animate-spin text-primary" />
          <CardTitle>Completing authentication</CardTitle>
          <CardDescription>
            FinCruiz is validating your secure link and preparing your workspace.
          </CardDescription>
        </CardHeader>
      </Card>
    </main>
  );
}

function AuthCallbackContent() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const [error, setError] = useState("");
  const [success, setSuccess] = useState(false);
  const [destination, setDestination] = useState("/login?reason=email-confirmed");

  useEffect(() => {
    const queryError =
      searchParams.get("error_description") || searchParams.get("error");

    if (queryError) {
      setError(queryError.replace(/\+/g, " "));
      return;
    }

    const hashParams = new URLSearchParams(
      window.location.hash.replace(/^#/, ""),
    );

    const hashError =
      hashParams.get("error_description") || hashParams.get("error");

    if (hashError) {
      setError(hashError.replace(/\+/g, " "));
      return;
    }

    const accessToken = hashParams.get("access_token");
    const refreshToken = hashParams.get("refresh_token");
    const expiresIn = Number(hashParams.get("expires_in") || 0);
    const isConfirmation = searchParams.get("confirmation") === "1";

    let next = "/login?reason=email-confirmed";

    if (accessToken) {
      authService.persistSession({
        access_token: accessToken,
        refresh_token: refreshToken,
        expires_in: expiresIn || undefined,
      });
      next = safeNext(searchParams.get("next"));

      window.history.replaceState(
        {},
        document.title,
        window.location.pathname + window.location.search,
      );
    } else if (!isConfirmation) {
      router.replace("/login");
      return;
    }

    setDestination(next);
    setSuccess(true);

    const redirectTimer = window.setTimeout(() => {
      router.replace(next);
    }, 2200);

    return () => window.clearTimeout(redirectTimer);
  }, [router, searchParams]);

  if (error) {
    return (
      <main className="flex min-h-screen items-center justify-center bg-muted/30 px-6">
        <Card className="w-full max-w-lg">
          <CardHeader>
            <TriangleAlert className="mb-3 size-10 text-amber-600" />
            <CardTitle>Authentication link could not be completed</CardTitle>
            <CardDescription>{error}</CardDescription>
          </CardHeader>
          <CardContent className="grid gap-3">
            <Button onClick={() => router.push("/login")}>Return to sign in</Button>
            <Button variant="outline" onClick={() => router.push("/signup")}>
              Create a new account
            </Button>
          </CardContent>
        </Card>
      </main>
    );
  }

  if (success) {
    return (
      <main className="flex min-h-screen items-center justify-center bg-[#f7f8fb] px-6">
        <Card className="w-full max-w-lg border-emerald-200/70 shadow-xl">
          <CardHeader>
            <div className="mb-3 flex size-12 items-center justify-center rounded-2xl bg-emerald-500/10 text-emerald-700">
              <CheckCircle2 className="size-6" />
            </div>
            <CardTitle>Email verified successfully</CardTitle>
            <CardDescription>
              Your email address has been confirmed. FinCruiz is taking you to the next secure step.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <Button className="w-full" onClick={() => router.replace(destination)}>
              Continue to FinCruiz
            </Button>
          </CardContent>
        </Card>
      </main>
    );
  }

  return <CallbackLoadingCard />;
}

export default function AuthCallbackPage() {
  return (
    <Suspense fallback={<CallbackLoadingCard />}>
      <AuthCallbackContent />
    </Suspense>
  );
}
