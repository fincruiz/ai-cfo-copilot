from typing import Annotated

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company, require_company_admin
from app.schemas.operations import OperationalReadiness, PaidLaunchCertificationOut
from app.schemas.responses import APIResponse
from app.services.operations_service import OperationsService
from app.services.paid_launch_certification_service import paid_launch_summary

router = APIRouter(prefix="/operations", tags=["Operations"])


@router.get("/readiness", response_model=APIResponse[OperationalReadiness])
async def operational_readiness(
    company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    data = await OperationsService(session).readiness(company.id)
    return APIResponse(message="Operational readiness retrieved.", data=OperationalReadiness(**data))


@router.get("/paid-launch-certification", response_model=APIResponse[PaidLaunchCertificationOut])
async def paid_launch_certification(
    company: Annotated[Company, Depends(get_current_company)],
    _admin=Depends(require_company_admin),
):
    # Company resolution + admin dependency intentionally gate this operator-facing
    # endpoint even though the checks themselves are deployment configuration checks.
    _ = company.id
    data = paid_launch_summary()
    return APIResponse(
        message="Paid launch certification retrieved. Operator-evidence checks are only green when explicitly recorded.",
        data=PaidLaunchCertificationOut(**data),
    )
