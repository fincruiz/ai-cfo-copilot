from typing import Literal
from pydantic import BaseModel

BillingProviderName = Literal["stripe", "razorpay"]
PaidPlanName = Literal["founding", "growth"]
BillingInterval = Literal["monthly", "annual"]


class CheckoutRequest(BaseModel):
    plan: PaidPlanName
    billing_interval: BillingInterval = "monthly"


class CheckoutSessionOut(BaseModel):
    provider: BillingProviderName
    checkout_url: str | None = None
    provider_session_id: str
    public_key: str | None = None
    subscription_id: str | None = None
    plan: PaidPlanName
    billing_interval: BillingInterval


class RazorpayVerifyRequest(BaseModel):
    razorpay_payment_id: str
    razorpay_subscription_id: str
    razorpay_signature: str


class BillingPortalOut(BaseModel):
    provider: BillingProviderName
    url: str


class BillingWebhookOut(BaseModel):
    received: bool = True
    duplicate: bool = False
