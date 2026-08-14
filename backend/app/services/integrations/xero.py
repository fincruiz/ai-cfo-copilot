from __future__ import annotations

from urllib.parse import urlencode
from uuid import UUID

import httpx

from app.core.config import settings
from app.services.integrations.base import IntegrationStore


AUTH_URL = "https://login.xero.com/identity/connect/authorize"
TOKEN_URL = "https://identity.xero.com/connect/token"
API_URL = "https://api.xero.com/api.xro/2.0"
CONNECTIONS_URL = "https://api.xero.com/connections"


# Xero introduced granular Accounting API scopes for apps created
# on/after 2 March 2026.
#
# FinCruiz is currently read/sync-only, so we deliberately request
# read-only permissions.
XERO_SCOPES = " ".join(
    [
        "openid",
        "profile",
        "email",
        "offline_access",
        "accounting.settings.read",
        "accounting.contacts.read",
        "accounting.invoices.read",
        "accounting.banktransactions.read",
    ]
)


class XeroConnector:
    def __init__(self, store: IntegrationStore):
        self.store = store

    def configured(self) -> bool:
        return bool(
            settings.xero_client_id
            and settings.xero_client_secret
            and settings.xero_redirect_uri
        )

    async def authorization_url(
        self,
        company_id: UUID,
        user_id: UUID,
    ) -> str:
        state = await self.store.save_oauth_state(
            company_id,
            user_id,
            "xero",
        )

        params = {
            "response_type": "code",
            "client_id": settings.xero_client_id,
            "redirect_uri": settings.xero_redirect_uri,
            "scope": XERO_SCOPES,
            "state": state,
        }

        return f"{AUTH_URL}?{urlencode(params)}"

    async def callback(
        self,
        code: str,
        state: str,
    ):
        ctx = await self.store.consume_oauth_state(
            state,
            "xero",
        )

        if not ctx:
            raise ValueError(
                "Invalid or expired Xero OAuth state."
            )

        async with httpx.AsyncClient(
            timeout=30
        ) as client:
            token_response = await client.post(
                TOKEN_URL,
                data={
                    "grant_type": "authorization_code",
                    "code": code,
                    "redirect_uri": settings.xero_redirect_uri,
                },
                auth=(
                    settings.xero_client_id or "",
                    settings.xero_client_secret or "",
                ),
            )

            token_response.raise_for_status()
            token_data = token_response.json()

            connection_response = await client.get(
                CONNECTIONS_URL,
                headers={
                    "Authorization": (
                        f"Bearer {token_data['access_token']}"
                    )
                },
            )

            connection_response.raise_for_status()
            tenants = connection_response.json()

        chosen = tenants[0] if len(tenants) == 1 else None

        await self.store.upsert_connection(
            company_id=ctx["company_id"],
            provider="xero",
            user_id=ctx["user_id"],
            status=(
                "connected"
                if chosen
                else "selection_required"
            ),
            external_tenant_id=(
                chosen.get("tenantId")
                if chosen
                else None
            ),
            external_tenant_name=(
                chosen.get("tenantName")
                if chosen
                else None
            ),
            access_token=token_data["access_token"],
            refresh_token=token_data.get(
                "refresh_token"
            ),
            expires_in=token_data.get(
                "expires_in"
            ),
            metadata={
                "tenants": tenants,
                "granted_scope": token_data.get(
                    "scope"
                ),
            },
        )

        return ctx["company_id"]

    async def select_tenant(
        self,
        company_id: UUID,
        tenant_id: str,
    ):
        connection = await self.store.get(
            company_id,
            "xero",
        )

        tenants = (
            connection or {}
        ).get(
            "metadata",
            {},
        ).get(
            "tenants",
            [],
        )

        chosen = next(
            (
                tenant
                for tenant in tenants
                if tenant.get("tenantId")
                == tenant_id
            ),
            None,
        )

        if not chosen:
            raise ValueError(
                "That Xero organisation is not "
                "available for this connection."
            )

        await self.store.upsert_connection(
            company_id=company_id,
            provider="xero",
            status="connected",
            external_tenant_id=tenant_id,
            external_tenant_name=chosen.get(
                "tenantName"
            ),
        )

    async def _refresh(
        self,
        company_id: UUID,
        connection: dict,
    ):
        refresh_token = connection.get(
            "refresh_token"
        )

        if not refresh_token:
            return connection

        async with httpx.AsyncClient(
            timeout=30
        ) as client:
            response = await client.post(
                TOKEN_URL,
                data={
                    "grant_type": "refresh_token",
                    "refresh_token": refresh_token,
                },
                auth=(
                    settings.xero_client_id or "",
                    settings.xero_client_secret or "",
                ),
            )

            response.raise_for_status()
            token_data = response.json()

        await self.store.upsert_connection(
            company_id=company_id,
            provider="xero",
            access_token=token_data[
                "access_token"
            ],
            refresh_token=token_data.get(
                "refresh_token"
            ),
            expires_in=token_data.get(
                "expires_in"
            ),
        )

        return await self.store.credentials(
            company_id,
            "xero",
        )

    async def sync(
        self,
        company_id: UUID,
    ):
        connection = await self.store.credentials(
            company_id,
            "xero",
        )

        if (
            not connection
            or not connection.get(
                "external_tenant_id"
            )
        ):
            raise ValueError(
                "Connect and select a Xero "
                "organisation first."
            )

        connection = await self._refresh(
            company_id,
            connection,
        )

        headers = {
            "Authorization": (
                f"Bearer "
                f"{connection['access_token']}"
            ),
            "Xero-tenant-id": connection[
                "external_tenant_id"
            ],
            "Accept": "application/json",
        }

        endpoints = {
            "account": "Accounts",
            "contact": "Contacts",
            "invoice": "Invoices",
            "bank_transaction": (
                "BankTransactions"
            ),
        }

        counts = {}

        try:
            async with httpx.AsyncClient(
                timeout=45
            ) as client:
                for entity, endpoint in endpoints.items():
                    response = await client.get(
                        f"{API_URL}/{endpoint}",
                        headers=headers,
                    )

                    response.raise_for_status()

                    body = response.json()
                    items = body.get(
                        endpoint,
                        [],
                    )

                    normalized = []

                    for item in items:
                        external_id = (
                            item.get("AccountID")
                            or item.get("ContactID")
                            or item.get("InvoiceID")
                            or item.get(
                                "BankTransactionID"
                            )
                        )

                        if not external_id:
                            continue

                        normalized.append(
                            {
                                "external_id": (
                                    external_id
                                ),
                                "name": (
                                    item.get("Name")
                                    or item.get(
                                        "Contact",
                                        {},
                                    ).get("Name")
                                    or item.get(
                                        "InvoiceNumber"
                                    )
                                    or item.get(
                                        "Reference"
                                    )
                                ),
                                "amount": (
                                    item.get("Total")
                                    or item.get(
                                        "Balance"
                                    )
                                ),
                                "currency_code": (
                                    item.get(
                                        "CurrencyCode"
                                    )
                                ),
                                "payload": item,
                            }
                        )

                    await self.store.replace_records(
                        company_id,
                        "xero",
                        entity,
                        normalized,
                    )

                    counts[entity] = len(
                        normalized
                    )

            await self.store.mark_sync(
                company_id,
                "xero",
                "success",
                (
                    f"Synced "
                    f"{sum(counts.values())} "
                    f"Xero records."
                ),
            )

            return counts

        except Exception as exc:
            await self.store.mark_sync(
                company_id,
                "xero",
                "failed",
                str(exc),
            )
            raise