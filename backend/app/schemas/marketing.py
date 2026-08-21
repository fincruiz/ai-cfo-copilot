from pydantic import BaseModel, Field, field_validator


class DemoLeadRequest(BaseModel):
    name: str = Field(min_length=2, max_length=120)
    work_email: str = Field(min_length=5, max_length=254)
    company_name: str = Field(min_length=2, max_length=180)
    role: str | None = Field(default=None, max_length=120)
    persona: str | None = Field(default=None, max_length=40)
    country: str | None = Field(default=None, max_length=100)
    team_size: str | None = Field(default=None, max_length=80)
    message: str | None = Field(default=None, max_length=1200)
    source_path: str | None = Field(default=None, max_length=250)
    website: str | None = Field(default=None, max_length=200)  # honeypot

    @field_validator("work_email")
    @classmethod
    def validate_email_shape(cls, value: str) -> str:
        cleaned = value.strip().lower()
        if "@" not in cleaned or cleaned.startswith("@") or cleaned.endswith("@") or "." not in cleaned.rsplit("@", 1)[-1]:
            raise ValueError("Enter a valid work email address.")
        return cleaned

    @field_validator("name", "company_name", "role", "persona", "country", "team_size", "message", "source_path", mode="before")
    @classmethod
    def strip_text(cls, value):
        return value.strip() if isinstance(value, str) else value


class DemoLeadResponse(BaseModel):
    accepted: bool
    reference: str | None = None
