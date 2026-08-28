import {
  ACCESS_TOKEN_KEY,
  api,
  REFRESH_TOKEN_KEY,
} from "@/lib/api";
import {
  beginSession,
  broadcastSessionLogout,
  clearSessionMetadata,
  type SessionLogoutReason,
} from "@/lib/session-security";

import type {
  ApiResponse,
  AuthTokens,
  Company,
  CurrentUser,
  LoginRequest,
  SignupResponse,
} from "@/types/auth";

function extractTokens(
  response: ApiResponse<AuthTokens> | AuthTokens,
): AuthTokens {
  if ("data" in response) {
    return response.data;
  }

  return response;
}

function persistTokens(tokens: AuthTokens): void {
  window.localStorage.setItem(
    ACCESS_TOKEN_KEY,
    tokens.access_token,
  );

  if (tokens.refresh_token) {
    window.localStorage.setItem(
      REFRESH_TOKEN_KEY,
      tokens.refresh_token,
    );
  }

  beginSession();
}

function clearLocalSession(reason: SessionLogoutReason): string | null {
  if (typeof window === "undefined") return null;
  const accessToken = window.localStorage.getItem(ACCESS_TOKEN_KEY);
  window.localStorage.removeItem(ACCESS_TOKEN_KEY);
  window.localStorage.removeItem(REFRESH_TOKEN_KEY);
  clearSessionMetadata();
  broadcastSessionLogout(reason);
  return accessToken;
}

async function revokeServerSession(
  scope: "global" | "local",
  reason: SessionLogoutReason,
): Promise<void> {
  const accessToken = clearLocalSession(reason);
  if (!accessToken) return;

  try {
    await api.post(
      "/auth/logout",
      { scope },
      {
        headers: { Authorization: `Bearer ${accessToken}` },
        timeout: 5000,
      },
    );
  } catch {
    // Local logout must succeed even when the network/auth service is unavailable.
  }
}

export const authService = {
  async signup(
    payload: Record<string, unknown>,
  ): Promise<SignupResponse> {
    const response = await api.post<
      ApiResponse<SignupResponse>
    >("/auth/signup", payload);

    const data = response.data.data;

    if (data.access_token) {
      persistTokens({
        access_token: data.access_token,
        refresh_token: data.refresh_token,
        expires_in: data.expires_in ?? undefined,
      });
    }

    return data;
  },

  async login(
    credentials: LoginRequest,
  ): Promise<AuthTokens> {
    const response = await api.post<
      ApiResponse<AuthTokens> | AuthTokens
    >("/auth/login", credentials);

    const tokens = extractTokens(response.data);

    if (!tokens.access_token) {
      throw new Error(
        "The login response did not contain an access token.",
      );
    }

    persistTokens(tokens);
    return tokens;
  },

  async getCurrentUser(): Promise<CurrentUser> {
    const response = await api.get<
      ApiResponse<CurrentUser>
    >("/auth/me");

    return response.data.data;
  },

  async getCurrentCompany(): Promise<Company> {
    const response = await api.get<
      ApiResponse<Company>
    >("/companies/me");

    return response.data.data;
  },

  persistSession(tokens: AuthTokens): void { persistTokens(tokens); },
  async resendConfirmation(email:string):Promise<void>{await api.post("/auth/resend-confirmation",{email});},
  async forgotPassword(email:string):Promise<void>{await api.post("/auth/forgot-password",{email});},
  async resetPassword(accessToken:string,password:string):Promise<void>{await api.post("/auth/reset-password",{access_token:accessToken,password});},

  logout(reason: SessionLogoutReason = "signed-out"): void {
    clearLocalSession(reason);
  },

  async logoutEverywhere(reason: SessionLogoutReason = "signed-out"): Promise<void> {
    await revokeServerSession("global", reason);
  },

  async logoutCurrentSession(reason: SessionLogoutReason = "signed-out"): Promise<void> {
    await revokeServerSession("local", reason);
  },

  hasAccessToken(): boolean {
    if (typeof window === "undefined") {
      return false;
    }

    return Boolean(
      window.localStorage.getItem(
        ACCESS_TOKEN_KEY,
      ),
    );
  },
};
