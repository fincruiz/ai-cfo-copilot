import axios, {
  AxiosError,
  InternalAxiosRequestConfig,
} from "axios";

const API_URL =
  process.env.NEXT_PUBLIC_API_URL ??
  "http://127.0.0.1:8000/api/v1";

export const ACCESS_TOKEN_KEY =
  "fincruiz_access_token";

export const REFRESH_TOKEN_KEY =
  "fincruiz_refresh_token";

export const api = axios.create({
  baseURL: API_URL,
  headers: {
    "Content-Type": "application/json",
    Accept: "application/json",
  },
  timeout: 30000,
});

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
      config.headers.Authorization =
        `Bearer ${accessToken}`;
    }

    return config;
  },
  (error: AxiosError) => {
    return Promise.reject(error);
  },
);

api.interceptors.response.use(
  (response) => response,
  (error: AxiosError) => {
    if (
      typeof window !== "undefined" &&
      error.response?.status === 401
    ) {
      window.localStorage.removeItem(
        ACCESS_TOKEN_KEY,
      );

      window.localStorage.removeItem(
        REFRESH_TOKEN_KEY,
      );
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

  return (
    responseData?.message ??
    responseData?.detail ??
    error.message ??
    "Unable to complete the request."
  );
}