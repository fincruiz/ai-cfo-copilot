from functools import lru_cache

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    app_name: str = "AI CFO Copilot API"
    app_version: str = "0.1.0"
    environment: str = "development"
    debug: bool = False

    api_v1_prefix: str = "/api/v1"

    database_url: str | None = None

    # Large-file ingestion. Point this at a persistent disk in production.
    import_staging_dir: str = "uploads/staging"
    import_max_upload_bytes: int = 1024 * 1024 * 1024  # 1 GiB
    import_chunk_rows: int = 2000
    import_worker_poll_seconds: float = 2.0
    # Operational hardening thresholds.
    import_staging_persistent: bool = False
    ingestion_stale_minutes: int = 30
    operational_failure_window_hours: int = 24
    slow_request_ms: int = 1500
    database_degraded_ms: int = 500
    database_unhealthy_ms: int = 1500

    supabase_url: str | None = None
    supabase_publishable_key: str | None = None
    supabase_service_role_key: str | None = None
    auth_frontend_url: str = "http://localhost:3000"

    openai_api_key: str | None = None
    openai_model: str = "gpt-5.6"



    # Commercial billing. Keep provider secrets server-side only.
    billing_frontend_url: str = "http://localhost:3000"
    billing_allow_live_payments: bool = False
    stripe_secret_key: str | None = None
    stripe_webhook_secret: str | None = None
    stripe_starter_monthly_price_id: str | None = None
    stripe_starter_annual_price_id: str | None = None
    stripe_growth_monthly_price_id: str | None = None
    stripe_growth_annual_price_id: str | None = None

    razorpay_key_id: str | None = None
    razorpay_key_secret: str | None = None
    razorpay_webhook_secret: str | None = None
    razorpay_starter_monthly_plan_id: str | None = None
    razorpay_starter_annual_plan_id: str | None = None
    razorpay_growth_monthly_plan_id: str | None = None
    razorpay_growth_annual_plan_id: str | None = None

    # Integration platform
    integration_encryption_key: str | None = None
    integration_frontend_url: str = "http://localhost:3000"

    xero_client_id: str | None = None
    xero_client_secret: str | None = None
    xero_redirect_uri: str | None = None

    zoho_client_id: str | None = None
    zoho_client_secret: str | None = None
    zoho_redirect_uri: str | None = None
    zoho_accounts_base_url: str = "https://accounts.zoho.com"
    zoho_api_base_url: str = "https://www.zohoapis.com/books/v3"

    # Comma-separated list, e.g.:
    # CORS_ORIGINS=http://localhost:3000,https://app.example.com
    cors_origins: str = (
        "http://localhost:3000,http://127.0.0.1:3000"
    )

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    @property
    def allowed_cors_origins(self) -> list[str]:
        return [
            origin.strip()
            for origin in self.cors_origins.split(",")
            if origin.strip()
        ]

    @property
    def is_production(self) -> bool:
        return self.environment.lower() == "production"


@lru_cache
def get_settings() -> Settings:
    return Settings()


settings = get_settings()
