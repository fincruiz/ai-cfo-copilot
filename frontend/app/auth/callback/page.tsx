"use client";

import { Suspense, useEffect, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import { Loader2, TriangleAlert } from "lucide-react";

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

    if (accessToken) {
      authService.persistSession({
        access_token: accessToken,
        refresh_token: refreshToken,
        expires_in: expiresIn || undefined,
      });

      window.history.replaceState(
        {},
        document.title,
        window.location.pathname + window.location.search,
      );

      router.replace(safeNext(searchParams.get("next")));
      return;
    }

    /*
     * Supabase can confirm the email server-side and return without
     * browser session tokens. In that case, route the user to the
     * FinCruiz login page rather than relying on localhost/null or
     * another implicit provider redirect.
     */
    router.replace("/login?reason=email-confirmed");
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
            <Button onClick={() => router.push("/login")}>
              Return to sign in
            </Button>

            <Button
              variant="outline"
              onClick={() => router.push("/signup")}
            >
              Create a new account
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