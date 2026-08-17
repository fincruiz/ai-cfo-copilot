import hashlib
import hmac
import json
import time

from app.services.billing.signatures import (
    verify_razorpay_subscription_payment,
    verify_razorpay_webhook,
    verify_stripe_webhook,
)


def test_razorpay_subscription_signature_matches_documented_order():
    secret = "test_secret"
    payment_id = "pay_123"
    subscription_id = "sub_456"
    signature = hmac.new(
        secret.encode(),
        f"{payment_id}|{subscription_id}".encode(),
        hashlib.sha256,
    ).hexdigest()
    assert verify_razorpay_subscription_payment(
        payment_id=payment_id,
        subscription_id=subscription_id,
        signature=signature,
        secret=secret,
    )


def test_razorpay_webhook_verifies_raw_body():
    secret = "webhook_secret"
    payload = json.dumps({"event": "subscription.activated"}).encode()
    signature = hmac.new(secret.encode(), payload, hashlib.sha256).hexdigest()
    assert verify_razorpay_webhook(payload, signature, secret)["event"] == "subscription.activated"


def test_stripe_webhook_verifies_timestamped_raw_body():
    secret = "whsec_test"
    payload = json.dumps({"id": "evt_1", "type": "invoice.paid"}).encode()
    timestamp = int(time.time())
    expected = hmac.new(secret.encode(), f"{timestamp}.".encode() + payload, hashlib.sha256).hexdigest()
    header = f"t={timestamp},v1={expected}"
    event = verify_stripe_webhook(payload, header, secret)
    assert event["id"] == "evt_1"
