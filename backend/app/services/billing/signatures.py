from __future__ import annotations
import hashlib
import hmac
import json
import time

from app.core.exceptions import ApplicationError


def verify_stripe_webhook(payload: bytes, signature_header: str, secret: str, tolerance_seconds: int = 300) -> dict:
    parts: dict[str, list[str]] = {}
    for chunk in signature_header.split(","):
        if "=" not in chunk:
            continue
        key, value = chunk.split("=", 1)
        parts.setdefault(key.strip(), []).append(value.strip())
    try:
        timestamp = int(parts["t"][0])
    except (KeyError, ValueError, IndexError):
        raise ApplicationError(message="Invalid Stripe webhook signature.", error_code="BILLING_WEBHOOK_INVALID", status_code=400)
    if abs(int(time.time()) - timestamp) > tolerance_seconds:
        raise ApplicationError(message="Expired Stripe webhook signature.", error_code="BILLING_WEBHOOK_EXPIRED", status_code=400)
    signed = f"{timestamp}.".encode() + payload
    expected = hmac.new(secret.encode(), signed, hashlib.sha256).hexdigest()
    if not any(hmac.compare_digest(expected, value) for value in parts.get("v1", [])):
        raise ApplicationError(message="Invalid Stripe webhook signature.", error_code="BILLING_WEBHOOK_INVALID", status_code=400)
    try:
        return json.loads(payload)
    except ValueError:
        raise ApplicationError(message="Invalid Stripe webhook payload.", error_code="BILLING_WEBHOOK_INVALID", status_code=400)


def verify_razorpay_webhook(payload: bytes, signature_header: str, secret: str) -> dict:
    expected = hmac.new(secret.encode(), payload, hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected, signature_header):
        raise ApplicationError(message="Invalid Razorpay webhook signature.", error_code="BILLING_WEBHOOK_INVALID", status_code=400)
    try:
        return json.loads(payload)
    except ValueError:
        raise ApplicationError(message="Invalid Razorpay webhook payload.", error_code="BILLING_WEBHOOK_INVALID", status_code=400)


def verify_razorpay_subscription_payment(*, payment_id: str, subscription_id: str, signature: str, secret: str) -> bool:
    message = f"{payment_id}|{subscription_id}".encode()
    expected = hmac.new(secret.encode(), message, hashlib.sha256).hexdigest()
    return hmac.compare_digest(expected, signature)
