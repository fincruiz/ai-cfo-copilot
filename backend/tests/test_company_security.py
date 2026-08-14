import asyncio
from types import SimpleNamespace
from uuid import uuid4

import pytest

from app.api.v1.core.companies.router import get_company, list_companies
from app.core.exceptions import ApplicationError
from app.schemas.auth import CurrentUser


def test_list_companies_is_scoped_to_current_user_memberships():
    async def run():
        user_id = uuid4()
        company_a = uuid4()
        company_b = uuid4()
        current_user = CurrentUser(id=user_id, email="user@example.com")

        class MemberService:
            async def list_active_memberships_by_user(self, *, user_id):
                assert user_id == current_user.id
                return [
                    SimpleNamespace(company_id=company_a),
                    SimpleNamespace(company_id=company_b),
                ]

        class CompanyService:
            async def list_companies_by_ids(
                self, *, company_ids, active_only, limit, offset
            ):
                assert company_ids == [company_a, company_b]
                assert active_only is True
                assert limit == 100
                assert offset == 0
                return [], 0

        response = await list_companies(
            current_user=current_user,
            service=CompanyService(),
            member_service=MemberService(),
            active_only=True,
            limit=100,
            offset=0,
        )
        assert response.count == 0
        assert response.data == []

    asyncio.run(run())


def test_get_company_hides_other_tenant_company():
    async def run():
        current_user = CurrentUser(id=uuid4(), email="user@example.com")
        other_company_id = uuid4()

        class MemberService:
            async def get_active_membership(self, *, user_id, company_id):
                assert user_id == current_user.id
                assert company_id == other_company_id
                return None

        class CompanyService:
            async def get_company(self, company_id):
                raise AssertionError("Company lookup must not run without membership")

        with pytest.raises(ApplicationError) as exc_info:
            await get_company(
                company_id=other_company_id,
                current_user=current_user,
                service=CompanyService(),
                member_service=MemberService(),
            )

        assert exc_info.value.status_code == 404
        assert exc_info.value.error_code == "COMPANY_NOT_FOUND"

    asyncio.run(run())
