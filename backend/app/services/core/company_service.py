from uuid import UUID

from app.database.models.core.company import Company
from app.repositories.core.company_repository import CompanyRepository
from app.schemas.core.company import CreateCompanyRequest
from app.services.base import BaseService


class CompanyService(BaseService[Company]):
    def __init__(
        self,
        repository: CompanyRepository,
    ) -> None:
        super().__init__(
            repository=repository,
            resource_name="Company",
        )

        self.company_repository = repository

    async def list_companies(
        self,
        *,
        active_only: bool,
        limit: int,
        offset: int,
    ) -> tuple[list[Company], int]:
        return await self.company_repository.list_companies(
            active_only=active_only,
            limit=limit,
            offset=offset,
        )

    async def get_company(
        self,
        company_id: UUID,
    ) -> Company:
        return await self.get_or_404(company_id)

    async def create_company(
        self,
        *,
        request: CreateCompanyRequest,
        created_by: UUID,
    ) -> Company:
        company = await self.company_repository.create(
            {
                "legal_name": request.legal_name.strip(),
                "trading_name": (
                    request.trading_name.strip()
                    if request.trading_name
                    else None
                ),
                "abn": (
                    request.abn.strip()
                    if request.abn
                    else None
                ),
                "country_code": request.country_code.upper(),
                "currency_code": request.currency_code.upper(),
                "financial_year_end_month": (
                    request.financial_year_end_month
                ),
                "industry": (
                    request.industry.strip()
                    if request.industry
                    else None
                ),
                "business_model": (
                    request.business_model.strip()
                    if request.business_model
                    else None
                ),
                "employee_count": request.employee_count,
                "annual_revenue": request.annual_revenue,
                "logo_path": request.logo_path,
                "website_url": (
                    str(request.website_url)
                    if request.website_url
                    else None
                ),
                "created_by": created_by,
            }
        )

        return company

async def update_logo(
    self,
    company: Company,
    logo_path: str,
) -> Company:
    updated = await self.company_repository.update(
        company,
        {"logo_path": logo_path},
    )
    await self.company_repository.session.commit()
    await self.company_repository.session.refresh(updated)
    return updated


async def update_company(self, company: Company, values: dict) -> Company:
    cleaned = {key: value for key, value in values.items() if value is not None}
    updated = await self.company_repository.update(company, cleaned)
    await self.company_repository.session.commit()
    await self.company_repository.session.refresh(updated)
    return updated
