import {
  ACCESS_TOKEN_KEY,
  api,
  REFRESH_TOKEN_KEY,
} from "@/lib/api";

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

export const authService = {

  async signup(payload: Record<string, unknown>): Promise<SignupResponse> {
    const response = await api.post<ApiResponse<SignupResponse>>("/auth/signup", payload);
    const data = response.data.data;
    if (data.access_token) {
      window.localStorage.setItem(ACCESS_TOKEN_KEY, data.access_token);
      if (data.refresh_token) window.localStorage.setItem(REFRESH_TOKEN_KEY, data.refresh_token);
    }
    return data;
  },

  async login(
    credentials: LoginRequest,
  ): Promise<AuthTokens> {
    const response = await api.post<
      ApiResponse<AuthTokens> | AuthTokens
    >(
      "/auth/login",
      credentials,
    );

    const tokens = extractTokens(response.data);

    if (!tokens.access_token) {
      throw new Error(
        "The login response did not contain an access token.",
      );
    }

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

  logout(): void {
    window.localStorage.removeItem(
      ACCESS_TOKEN_KEY,
    );

    window.localStorage.removeItem(
      REFRESH_TOKEN_KEY,
    );
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