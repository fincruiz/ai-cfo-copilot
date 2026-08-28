import axios, {
  AxiosError,
  InternalAxiosRequestConfig,
} from "axios";

import { broadcastSessionLogout, clearSessionMetadata } from "@/lib/session-security";

const API_URL =
  process.env.NEXT_PUBLIC_API_URL ??
  "http://127.0.0.1:8000/api/v1";

export const ACCESS_TOKEN_KEY = "fincruiz_access_token";
export const REFRESH_TOKEN_KEY = "fincruiz_refresh_token";

type RetryableRequestConfig = InternalAxiosRequestConfig & {
  _retry?: boolean;
};

type TokenPayload = {
  access_token: string;
  refresh_token?: string | null;
};

type TokenApiResponse = {
  success: boolean;
  message: string;
  data: TokenPayload;
};

export const api = axios.create({
  baseURL: API_URL,
  headers: {
    "Content-Type": "application/json",
    Accept: "application/json",
  },
  timeout: 30000,
});

function clearStoredSession(reason: "session-expired" | "signed-out" = "session-expired"): void {
  if (typeof window === "undefined") return;
  window.localStorage.removeItem(ACCESS_TOKEN_KEY);
  window.localStorage.removeItem(REFRESH_TOKEN_KEY);
  clearSessionMetadata();
  broadcastSessionLogout(reason);
}

function storeTokens(tokens: TokenPayload): void {
  if (typeof window === "undefined") return;

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
}

let refreshPromise: Promise<string> | null = null;

async function refreshAccessToken(): Promise<string> {
  if (typeof window === "undefined") {
    throw new Error("Token refresh is only available in the browser.");
  }

  const refreshToken = window.localStorage.getItem(
    REFRESH_TOKEN_KEY,
  );

  if (!refreshToken) {
    throw new Error("No refresh token is available.");
  }

  if (!refreshPromise) {
    refreshPromise = axios
      .post<TokenApiResponse>(
        `${API_URL}/auth/refresh`,
        { refresh_token: refreshToken },
        {
          headers: {
            "Content-Type": "application/json",
            Accept: "application/json",
          },
          timeout: 15000,
        },
      )
      .then((response) => {
        const tokens = response.data.data;

        if (!tokens?.access_token) {
          throw new Error(
            "The refresh response did not contain an access token.",
          );
        }

        storeTokens(tokens);
        return tokens.access_token;
      })
      .finally(() => {
        refreshPromise = null;
      });
  }

  const pendingRefresh = refreshPromise;
  if (!pendingRefresh) {
    throw new Error("Unable to start token refresh.");
  }

  return pendingRefresh;
}

api.interceptors.request.use(
  (
    config: InternalAxiosRequestConfig,
  ): InternalAxiosRequestConfig => {
    if (typeof window === "undefined") {
      return config;
    }

    const accessToken = window.localStorage.getItem(
      ACCESS_TOKEN_KEY,
    );

    if (accessToken) {
      config.headers.Authorization = `Bearer ${accessToken}`;
    }

    return config;
  },
  (error: AxiosError) => Promise.reject(error),
);

api.interceptors.response.use(
  (response) => response,
  async (error: AxiosError) => {
    const originalRequest = error.config as
      | RetryableRequestConfig
      | undefined;

    const isUnauthorized = error.response?.status === 401;
    const isRefreshRequest = originalRequest?.url?.includes(
      "/auth/refresh",
    );
    const isLoginRequest = originalRequest?.url?.includes(
      "/auth/login",
    );

    if (
      typeof window !== "undefined" &&
      isUnauthorized &&
      originalRequest &&
      !originalRequest._retry &&
      !isRefreshRequest &&
      !isLoginRequest
    ) {
      originalRequest._retry = true;

      try {
        const newAccessToken = await refreshAccessToken();
        originalRequest.headers.Authorization =
          `Bearer ${newAccessToken}`;
        return api(originalRequest);
      } catch {
        clearStoredSession();

        if (window.location.pathname.startsWith("/dashboard")) {
          window.location.replace("/login?reason=session-expired");
        }
      }
    }

    if (
      typeof window !== "undefined" &&
      isUnauthorized &&
      (isRefreshRequest || originalRequest?._retry)
    ) {
      clearStoredSession();
    }

    return Promise.reject(error);
  },
);

export function getApiErrorMessage(
  error: unknown,
): string {
  if (!axios.isAxiosError(error)) {
    return "Something went wrong. Please try again.";
  }

  const responseData = error.response?.data as
    | {
        message?: string;
        detail?: string;
      }
    | undefined;

  if (error.code === "ECONNABORTED") {
    return "FinCruiz is taking longer than expected to respond. Please wait a moment and try again. If this continues, contact support.";
  }

  return (
    responseData?.message ??
    responseData?.detail ??
    error.message ??
    "Unable to complete the request."
  );
}


export function getApiSupportId(error: unknown): string | null {
  if (!axios.isAxiosError(error)) return null;
  const body = error.response?.data as { support_id?: string | null } | undefined;
  const header = error.response?.headers?.["x-request-id"];
  return body?.support_id ?? (typeof header === "string" ? header : null);
}
