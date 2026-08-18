from typing import Annotated
from fastapi import APIRouter, Depends, Header, Request
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.config import settings
from app.core.exceptions import ApplicationError
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company, require_company_admin
from app.schemas.billing import BillingPortalOut, BillingWebhookOut, CheckoutRequest, CheckoutSessionOut, RazorpayVerifyRequest, BillingReadinessOut, BillingEventSummaryOut
from app.schemas.responses import APIResponse
from app.services.billing.service import BillingService
from app.services.billing.signatures import verify_razorpay_subscription_payment, verify_razorpay_webhook, verify_stripe_webhook

router = APIRouter(prefix="/billing", tags=["Billing"])




@router.get("/readiness", response_model=APIResponse[BillingReadinessOut])
async def billing_readiness(
    company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
    _admin=Depends(require_company_admin),
):
    data = await BillingService(session).readiness(company_id=company.id)
    return APIResponse(message="Billing certification readiness retrieved.", data=BillingReadinessOut(**data))


@router.get("/events", response_model=APIResponse[list[BillingEventSummaryOut]])
async def billing_events(
    company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
    _admin=Depends(require_company_admin),
):
    data = await BillingService(session).recent_events(company_id=company.id)
    return APIResponse(message="Recent verified billing events retrieved.", data=[BillingEventSummaryOut(**row) for row in data])

@router.post("/checkout", response_model=APIResponse[CheckoutSessionOut])
async def checkout(
    payload: CheckoutRequest,
    company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
    _admin=Depends(require_company_admin),
):
    email = None
    result = await BillingService(session).create_checkout(company_id=company.id, email=email, plan=payload.plan, interval=payload.billing_interval)
    return APIResponse(message="Billing checkout created.", data=CheckoutSessionOut(**result))


@router.post("/portal", response_model=APIResponse[BillingPortalOut])
async def portal(
    company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
    _admin=Depends(require_company_admin),
):
    return APIResponse(message="Billing portal created.", data=BillingPortalOut(**await BillingService(session).create_portal(company_id=company.id)))


@router.post("/razorpay/verify", response_model=APIResponse[dict])
async def razorpay_verify(
    payload: RazorpayVerifyRequest,
    company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
    _admin=Depends(require_company_admin),
):
    if not settings.razorpay_key_secret:
        raise ApplicationError(message="Razorpay billing is not configured.", error_code="BILLING_PROVIDER_NOT_CONFIGURED", status_code=503)
    row = (await session.execute(
        __import__("sqlalchemy").text("SELECT last_checkout_id FROM public.company_subscriptions WHERE company_id=:company_id"),
        {"company_id": company.id},
    )).scalar_one_or_none()
    if str(row or "") != payload.razorpay_subscription_id:
        raise ApplicationError(message="Subscription does not belong to this workspace.", error_code="BILLING_SUBSCRIPTION_MISMATCH", status_code=403)
    valid = verify_razorpay_subscription_payment(
        payment_id=payload.razorpay_payment_id,
        subscription_id=payload.razorpay_subscription_id,
        signature=payload.razorpay_signature,
        secret=settings.razorpay_key_secret,
    )
    if not valid:
        raise ApplicationError(message="Razorpay payment verification failed.", error_code="BILLING_PAYMENT_INVALID", status_code=400)
    return APIResponse(message="Payment signature verified. Subscription access will be activated by the provider webhook.", data={"verified": True})


@router.post("/webhooks/stripe", response_model=BillingWebhookOut)
async def stripe_webhook(request: Request, stripe_signature: Annotated[str | None, Header(alias="Stripe-Signature")] = None, session: AsyncSession = Depends(get_db_session)):
    if not settings.stripe_webhook_secret or not stripe_signature:
        raise ApplicationError(message="Stripe webhook verification is not configured.", error_code="BILLING_WEBHOOK_NOT_CONFIGURED", status_code=503)
    payload = await request.body()
    event = verify_stripe_webhook(payload, stripe_signature, settings.stripe_webhook_secret)
    duplicate = await BillingService(session).handle_stripe_event(event)
    return BillingWebhookOut(received=True, duplicate=duplicate)


@router.post("/webhooks/razorpay", response_model=BillingWebhookOut)
async def razorpay_webhook(request: Request, x_razorpay_signature: Annotated[str | None, Header(alias="X-Razorpay-Signature")] = None, session: AsyncSession = Depends(get_db_session)):
    if not settings.razorpay_webhook_secret or not x_razorpay_signature:
        raise ApplicationError(message="Razorpay webhook verification is not configured.", error_code="BILLING_WEBHOOK_NOT_CONFIGURED", status_code=503)
    payload = await request.body()
    event = verify_razorpay_webhook(payload, x_razorpay_signature, settings.razorpay_webhook_secret)
    duplicate = await BillingService(session).handle_razorpay_event(event)
    return BillingWebhookOut(received=True, duplicate=duplicate)
