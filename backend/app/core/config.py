from functools import lru_cache

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    app_name: str = "AI CFO Copilot API"
    app_version: str = "0.1.0"
    environment: str = "development"
    debug: bool = False

    api_v1_prefix: str = "/api/v1"

    database_url: str | None = None

    supabase_url: str | None = None
    supabase_publishable_key: str | None = None
    supabase_service_role_key: str | None = None

    openai_api_key: str | None = None
    openai_model: str = "gpt-5.6"

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
