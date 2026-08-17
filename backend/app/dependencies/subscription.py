from fastapi import Depends
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.exceptions import ApplicationError
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.services.subscription_service import SubscriptionService

def require_entitlement(entitlement:str):
    async def dependency(company:Company=Depends(get_current_company),session:AsyncSession=Depends(get_db_session)):
        status=await SubscriptionService(session).status(company_id=company.id)
        if not status['is_access_active']:
            raise ApplicationError(message='Your FinCruiz trial or subscription is not active.',error_code='SUBSCRIPTION_ACCESS_INACTIVE',status_code=402)
        if status['entitlements'].get(entitlement) is not True:
            raise ApplicationError(message='Your current plan does not include this capability.',error_code='ENTITLEMENT_REQUIRED',status_code=403)
        return status
    return dependency
