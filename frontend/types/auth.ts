export interface LoginRequest {
  email: string;
  password: string;
}

export interface AuthTokens {
  access_token: string;
  refresh_token?: string | null;
  token_type?: string;
  expires_in?: number;
}

export interface ApiResponse<T> {
  success: boolean;
  message: string;
  data: T;
}

export interface CurrentUser {
  id: string;
  email?: string | null;
  phone?: string | null;
  role?: string | null;
  aud?: string | null;
  user_metadata?: Record<string, unknown>;
}

export interface Company {
  id: string;
  legal_name: string;
  trading_name?: string | null;
  abn?: string | null;
  country_code: string;
  currency_code: string;
  financial_year_end_month: number;
  industry?: string | null;
  business_model?: string | null;
  employee_count?: number | null;
  annual_revenue?: string | number | null;
  logo_path?: string | null;
  website_url?: string | null;
  is_active: boolean;
  created_by?: string | null;
  created_at: string;
  updated_at: string;
}

export interface CreateCompanyRequest {
  legal_name: string;
  trading_name?: string | null;
  abn?: string | null;
  country_code: string;
  currency_code: string;
  financial_year_end_month: number;
  industry?: string | null;
  business_model?: string | null;
  employee_count?: number | null;
  annual_revenue?: number | null;
  logo_path?: string | null;
  website_url?: string | null;
}
export interface SignupResponse { confirmation_required: boolean; email: string; access_token?: string | null; refresh_token?: string | null; expires_in?: number | null; }
