import asyncio
from types import SimpleNamespace

from app.services.core.company_service import CompanyService


def test_company_profile_update_methods_are_real_service_methods():
    async def run():
        class FakeSession:
            async def commit(self):
                return None

            async def refresh(self, record):
                return None

        class FakeRepository:
            session = FakeSession()

            async def update(self, record, values):
                for key, value in values.items():
                    setattr(record, key, value)
                return record

        service = CompanyService(FakeRepository())
        company = SimpleNamespace(logo_path=None, trading_name="Old")

        updated = await service.update_logo(company, "/uploads/logos/logo.png")
        assert updated.logo_path == "/uploads/logos/logo.png"

        updated = await service.update_company(
            company,
            {"trading_name": "New", "industry": None},
        )
        assert updated.trading_name == "New"

    asyncio.run(run())
