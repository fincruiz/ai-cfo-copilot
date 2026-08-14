from __future__ import annotations
import base64, hashlib
from cryptography.fernet import Fernet
from app.core.config import settings


def _fernet() -> Fernet:
    raw = settings.integration_encryption_key
    if not raw:
        raise RuntimeError("INTEGRATION_ENCRYPTION_KEY is not configured.")
    try:
        return Fernet(raw.encode())
    except Exception:
        # Permit a long random secret and derive a valid Fernet key from it.
        key = base64.urlsafe_b64encode(hashlib.sha256(raw.encode()).digest())
        return Fernet(key)


def encrypt_secret(value: str | None) -> str | None:
    return _fernet().encrypt(value.encode()).decode() if value else None


def decrypt_secret(value: str | None) -> str | None:
    return _fernet().decrypt(value.encode()).decode() if value else None
