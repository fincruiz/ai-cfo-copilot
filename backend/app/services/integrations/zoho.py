from __future__ import annotations

from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from typing import Any
from urllib.parse import urlencode
from uuid import UUID

import httpx

from app.core.config import settings
from app.services.integrations.base import IntegrationStore

SCOPES = "ZohoBooks.fullaccess.all"


def _decimal(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal("0")
    try:
        return Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError):
        return Decimal("0")


def _zoho_datetime(value: Any) -> datetime | None:
    if not value:
        return None
    if isinstance(value, datetime):
        return value if value.tzinfo else value.replace(tzinfo=timezone.utc)
    text = str(value).strip().replace("Z", "+00:00")
    try:
        result = datetime.fromisoformat(text)
        return result if result.tzinfo else result.replace(tzinfo=timezone.utc)
    except ValueError:
        try:
            return datetime.fromisoformat(text[:10]).replace(tzinfo=timezone.utc)
        except ValueError:
            return None


class ZohoConnector:
    def __init__(self, store: IntegrationStore):
        self.store = store

    def configured(self):
        return bool(
            settings.zoho_client_id
            and settings.zoho_client_secret
            and settings.zoho_redirect_uri
        )

    @staticmethod
    def _books_base(metadata: dict[str, Any] | None = None) -> str:
        api_domain = str((metadata or {}).get("api_domain") or "").rstrip("/")
        if api_domain:
            return f"{api_domain}/books/v3"
        return settings.zoho_api_base_url.rstrip("/")

    async def authorization_url(self, company_id: UUID, user_id: UUID):
        state = await self.store.save_oauth_state(company_id, user_id, "zoho")
        return settings.zoho_accounts_base_url + "/oauth/v2/auth?" + urlencode(
            {
                "scope": SCOPES,
                "client_id": settings.zoho_client_id,
                "response_type": "code",
                "access_type": "offline",
                "prompt": "consent",
                "redirect_uri": settings.zoho_redirect_uri,
                "state": state,
            }
        )

    async def callback(self, code: str, state: str):
        ctx = await self.store.consume_oauth_state(state, "zoho")
        if not ctx:
            raise ValueError("Invalid or expired Zoho OAuth state.")
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(
                settings.zoho_accounts_base_url + "/oauth/v2/token",
                params={
                    "grant_type": "authorization_code",
                    "client_id": settings.zoho_client_id,
                    "client_secret": settings.zoho_client_secret,
                    "redirect_uri": settings.zoho_redirect_uri,
                    "code": code,
                },
            )
            response.raise_for_status()
            token = response.json()
            api_domain = token.get("api_domain")
            books_base = self._books_base({"api_domain": api_domain})
            org_response = await client.get(
                f"{books_base}/organizations",
                headers={"Authorization": f"Zoho-oauthtoken {token['access_token']}"},
            )
            org_response.raise_for_status()
            organisations = org_response.json().get("organizations", [])

        chosen = organisations[0] if len(organisations) == 1 else None
        await self.store.upsert_connection(
            company_id=ctx["company_id"],
            provider="zoho",
            user_id=ctx["user_id"],
            status="connected" if chosen else "selection_required",
            external_tenant_id=str(chosen.get("organization_id")) if chosen else None,
            external_tenant_name=chosen.get("name") if chosen else None,
            access_token=token["access_token"],
            refresh_token=token.get("refresh_token"),
            expires_in=token.get("expires_in_sec") or token.get("expires_in"),
            metadata={"organizations": organisations, "api_domain": api_domain},
        )
        return ctx["company_id"]

    async def select_tenant(self, company_id: UUID, tenant_id: str):
        connection = await self.store.get(company_id, "zoho")
        organisations = (connection or {}).get("metadata", {}).get("organizations", [])
        chosen = next(
            (
                item
                for item in organisations
                if str(item.get("organization_id")) == tenant_id
            ),
            None,
        )
        if not chosen:
            raise ValueError("That Zoho Books organisation is not available.")
        await self.store.upsert_connection(
            company_id=company_id,
            provider="zoho",
            status="connected",
            external_tenant_id=tenant_id,
            external_tenant_name=chosen.get("name"),
        )

    async def _refresh(self, company_id: UUID, connection: dict):
        if not connection.get("refresh_token"):
            return connection
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(
                settings.zoho_accounts_base_url + "/oauth/v2/token",
                params={
                    "grant_type": "refresh_token",
                    "client_id": settings.zoho_client_id,
                    "client_secret": settings.zoho_client_secret,
                    "refresh_token": connection["refresh_token"],
                },
            )
            response.raise_for_status()
            token = response.json()
        await self.store.upsert_connection(
            company_id=company_id,
            provider="zoho",
            access_token=token["access_token"],
            expires_in=token.get("expires_in_sec") or token.get("expires_in"),
        )
        if token.get("api_domain"):
            await self.store.merge_metadata(
                company_id,
                "zoho",
                {"api_domain": token["api_domain"]},
            )
        return await self.store.credentials(company_id, "zoho")

    async def _get_all_pages(
        self,
        client: httpx.AsyncClient,
        *,
        base_url: str,
        endpoint: str,
        key: str,
        headers: dict[str, str],
        params: dict[str, Any],
        max_pages: int = 5000,
    ) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for page in range(1, max_pages + 1):
            query = {**params, "page": page, "per_page": 200}
            response = await client.get(
                f"{base_url}/{endpoint}", headers=headers, params=query
            )
            response.raise_for_status()
            body = response.json()
            batch = body.get(key, [])
            rows.extend(batch)
            page_context = body.get("page_context") or {}
            if page_context:
                if not page_context.get("has_more_page"):
                    break
            elif len(batch) < 200:
                break
        return rows

    async def _account_transactions(
        self,
        client: httpx.AsyncClient,
        *,
        base_url: str,
        headers: dict[str, str],
        organisation_id: str,
        account: dict[str, Any],
    ) -> list[dict[str, Any]]:
        account_id = str(account.get("account_id") or "")
        if not account_id or account.get("is_involved_in_transaction") is False:
            return []

        rows: list[dict[str, Any]] = []
        params = {
            "organization_id": organisation_id,
            "account_id": account_id,
            "filter_by": "TransactionType.All",
        }
        for page in range(1, 5001):
            response = await client.get(
                f"{base_url}/chartofaccounts/accounttransactions",
                headers=headers,
                params={**params, "page": page, "per_page": 200},
            )
            response.raise_for_status()
            body = response.json()
            batch = body.get("transactions", [])
            rows.extend(batch)
            page_context = body.get("page_context") or {}
            if page_context:
                if not page_context.get("has_more_page"):
                    break
            elif len(batch) < 200:
                break
        return rows

    @staticmethod
    def _simple_records(items: list[dict[str, Any]], entity: str) -> list[dict[str, Any]]:
        id_keys = {
            "account": ("account_id",),
            "contact": ("contact_id",),
            "invoice": ("invoice_id",),
            "bill": ("bill_id",),
            "journal": ("journal_id",),
            "bank_transaction": ("transaction_id", "bank_transaction_id"),
        }
        name_keys = (
            "account_name",
            "contact_name",
            "invoice_number",
            "bill_number",
            "vendor_name",
            "entry_number",
            "reference_number",
        )
        records = []
        for item in items:
            external_id = next(
                (item.get(key) for key in id_keys.get(entity, ()) if item.get(key)),
                None,
            )
            if not external_id:
                continue
            name = next((item.get(key) for key in name_keys if item.get(key)), None)
            records.append(
                {
                    "external_id": str(external_id),
                    "name": name,
                    "amount": item.get("total") or item.get("balance"),
                    "currency_code": item.get("currency_code"),
                    "occurred_at": _zoho_datetime(
                        item.get("date")
                        or item.get("journal_date")
                        or item.get("transaction_date")
                    ),
                    "source_updated_at": _zoho_datetime(
                        item.get("last_modified_time") or item.get("updated_time")
                    ),
                    "payload": item,
                }
            )
        return records

    @staticmethod
    def _gl_records(
        accounts: list[dict[str, Any]],
        account_transactions: dict[str, list[dict[str, Any]]],
        *,
        functional_currency_code: str | None = None,
    ) -> list[dict[str, Any]]:
        account_lookup = {
            str(account.get("account_id")): account
            for account in accounts
            if account.get("account_id")
        }
        lines: list[dict[str, Any]] = []
        for account_id, transactions in account_transactions.items():
            account = account_lookup.get(account_id, {})
            account_code = account.get("account_code") or f"ZOHO:{account_id}"
            account_name = account.get("account_name")
            for index, tx in enumerate(transactions, start=1):
                debit = _decimal(tx.get("debit_amount"))
                credit = _decimal(tx.get("credit_amount"))
                if debit == 0 and credit == 0:
                    direction = str(tx.get("debit_or_credit") or "").lower()
                    amount = _decimal(tx.get("amount"))
                    if direction == "debit":
                        debit = abs(amount)
                    elif direction == "credit":
                        credit = abs(amount)
                if debit == 0 and credit == 0:
                    continue

                transaction_id = (
                    tx.get("transaction_id")
                    or tx.get("categorized_transaction_id")
                    or tx.get("entry_number")
                    or f"{tx.get('transaction_date')}:{index}"
                )
                line_identity = (
                    tx.get("categorized_transaction_id")
                    or tx.get("line_id")
                    or (
                        f"{transaction_id}:{tx.get('transaction_date')}:"
                        f"{debit}:{credit}:{index}"
                    )
                )
                external_id = f"register:{account_id}:{line_identity}"
                occurred_at = _zoho_datetime(tx.get("transaction_date"))
                branch_reference = tx.get("location_name") or tx.get("branch_name")
                lines.append(
                    {
                        "external_id": external_id,
                        "name": account_name or tx.get("description"),
                        "amount": debit or credit,
                        "occurred_at": occurred_at,
                        "payload": {
                            "source": "zoho_account_register",
                            "account_id": account_id,
                            "raw_transaction": tx,
                            "fincruiz": {
                                "record_kind": "canonical_gl_line",
                                "transaction_date": (
                                    occurred_at.date().isoformat() if occurred_at else None
                                ),
                                "account_code": str(account_code),
                                "account_name": account_name,
                                "debit": str(debit),
                                "credit": str(credit),
                                "description": tx.get("description")
                                or tx.get("transaction_type_formatted"),
                                "reference": tx.get("reference_number"),
                                "document_number": tx.get("entry_number"),
                                "journal_number": tx.get("entry_number"),
                                "customer_code": tx.get("customer_id"),
                                "supplier_code": tx.get("vendor_id"),
                                "source_transaction_id": str(transaction_id),
                                "source_line_id": str(line_identity),
                                "source_type": tx.get("transaction_type"),
                                "branch_reference": branch_reference,
                                # Register values are treated as base-ledger amounts. The
                                # canonical service will use the company functional currency.
                                "functional_currency_code": functional_currency_code,
                                "transaction_currency_code": tx.get("currency_code"),
                                "exchange_rate": "1",
                            },
                        },
                    }
                )
        return lines

    async def sync(self, company_id: UUID):
        connection = await self.store.credentials(company_id, "zoho")
        if not connection or not connection.get("external_tenant_id"):
            raise ValueError("Connect and select a Zoho Books organisation first.")
        connection = await self._refresh(company_id, connection)
        metadata = connection.get("metadata") or {}
        base_url = self._books_base(metadata)
        headers = {"Authorization": f"Zoho-oauthtoken {connection['access_token']}"}
        organisation_id = str(connection["external_tenant_id"])
        organisations = metadata.get("organizations") or []
        selected_organisation = next(
            (item for item in organisations if str(item.get("organization_id")) == organisation_id),
            {},
        )
        functional_currency = selected_organisation.get("currency_code")
        common = {"organization_id": organisation_id}

        try:
            async with httpx.AsyncClient(timeout=60) as client:
                accounts = await self._get_all_pages(
                    client,
                    base_url=base_url,
                    endpoint="chartofaccounts",
                    key="chartofaccounts",
                    headers=headers,
                    params=common,
                )
                contacts = await self._get_all_pages(
                    client,
                    base_url=base_url,
                    endpoint="contacts",
                    key="contacts",
                    headers=headers,
                    params=common,
                )
                invoices = await self._get_all_pages(
                    client,
                    base_url=base_url,
                    endpoint="invoices",
                    key="invoices",
                    headers=headers,
                    params=common,
                )
                bills = await self._get_all_pages(
                    client,
                    base_url=base_url,
                    endpoint="bills",
                    key="bills",
                    headers=headers,
                    params=common,
                )
                journals = await self._get_all_pages(
                    client,
                    base_url=base_url,
                    endpoint="journals",
                    key="journals",
                    headers=headers,
                    params=common,
                )

                account_transactions: dict[str, list[dict[str, Any]]] = {}
                for account in accounts:
                    account_id = str(account.get("account_id") or "")
                    if not account_id:
                        continue
                    account_transactions[account_id] = await self._account_transactions(
                        client,
                        base_url=base_url,
                        headers=headers,
                        organisation_id=organisation_id,
                        account=account,
                    )

                account_records = self._simple_records(accounts, "account")
                contact_records = self._simple_records(contacts, "contact")
                invoice_records = self._simple_records(invoices, "invoice")
                bill_records = self._simple_records(bills, "bill")
                journal_records = self._simple_records(journals, "journal")
                gl_lines = self._gl_records(
                    accounts,
                    account_transactions,
                    functional_currency_code=functional_currency,
                )

                snapshots = {
                    "account": account_records,
                    "contact": contact_records,
                    "invoice": invoice_records,
                    "bill": bill_records,
                    "journal": journal_records,
                    "gl_line": gl_lines,
                }
                for entity_type, records in snapshots.items():
                    await self.store.replace_records_snapshot(
                        company_id, "zoho", entity_type, records, commit=False
                    )

                counts = {
                    "accounts": len(account_records),
                    "contacts": len(contact_records),
                    "invoices": len(invoice_records),
                    "bills": len(bill_records),
                    "journals": len(journal_records),
                    "gl_lines": len(gl_lines),
                }

            await self.store.merge_metadata(
                company_id,
                "zoho",
                {
                    "source_counts": counts,
                    "source_functional_currency": functional_currency,
                    "register_sync": {
                        "status": "available",
                        "message": (
                            f"Retrieved {counts['gl_lines']:,} account-register GL lines "
                            "for finance-truth activation."
                        ),
                    },
                },
                commit=False,
            )
            await self.store.mark_sync(
                company_id,
                "zoho",
                "success",
                (
                    f"Synced {sum(counts.values()):,} Zoho Books source records including "
                    f"{counts['gl_lines']:,} journal-grade account-register lines."
                ),
                commit=False,
            )
            await self.store.session.commit()
            return counts
        except httpx.HTTPStatusError as exc:
            await self.store.session.rollback()
            message = f"Zoho Books API {exc.response.status_code}: {exc.response.text[:700]}"
            await self.store.mark_sync(company_id, "zoho", "failed", message)
            raise ValueError(message) from exc
        except Exception as exc:
            await self.store.session.rollback()
            await self.store.mark_sync(company_id, "zoho", "failed", str(exc))
            raise
