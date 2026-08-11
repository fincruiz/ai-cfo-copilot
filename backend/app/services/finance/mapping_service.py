from __future__ import annotations

from uuid import UUID

from app.domain.finance.mapping.classifier import (
    suggest_mapping,
)
from app.repositories.finance.account_mapping_repository import (
    AccountMappingRepository,
)


class MappingService:
    def __init__(
        self,
        repository: AccountMappingRepository,
    ) -> None:
        self.repository = repository
        self.session = repository.session

    async def list_mappings(
        self,
        company_id: UUID,
    ):
        return await self.repository.list_mappings(
            company_id
        )

    async def suggest_unmapped(
        self,
        company_id: UUID,
    ) -> list[dict]:
        suggestions: list[dict] = []

        unmapped_accounts = (
            await self.repository.unmapped_accounts(
                company_id
            )
        )

        for account_code, account_name in unmapped_accounts:
            suggestion = suggest_mapping(
                account_code,
                account_name,
            )

            suggestions.append(
                {
                    "source_account_code":
                        account_code,
                    "source_account_name":
                        account_name,
                    "statement":
                        suggestion.statement,
                    "reporting_group":
                        suggestion.reporting_group,
                    "reporting_subgroup":
                        suggestion.reporting_subgroup,
                    "sign_convention":
                        suggestion.sign_convention,
                    "confidence":
                        suggestion.confidence,
                    "reason":
                        suggestion.reason,
                }
            )

        return suggestions

    async def upsert(
        self,
        company_id: UUID,
        items,
    ) -> int:
        rows: list[dict] = []

        for item in items:
            row = item.model_dump()
            row["company_id"] = company_id
            rows.append(row)

        saved_count = (
            await self.repository.upsert_many(
                rows
            )
        )

        await self.session.commit()

        return saved_count