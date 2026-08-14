from __future__ import annotations

import re
from datetime import datetime, timezone
from decimal import Decimal
from typing import Any
from urllib.parse import urlencode
from uuid import UUID

import httpx

from app.core.config import settings
from app.services.integrations.base import IntegrationStore

AUTH_URL = "https://login.xero.com/identity/connect/authorize"
TOKEN_URL = "https://identity.xero.com/connect/token"
API_URL = "https://api.xero.com/api.xro/2.0"
CONNECTIONS_URL = "https://api.xero.com/connections"

# Xero apps created on/after 2 March 2026 use granular scopes.
# FinCruiz is intentionally read-only at this stage.
XERO_SCOPES = " ".join(
    [
        "openid",
        "profile",
        "email",
        "offline_access",
        "accounting.settings.read",
        "accounting.contacts.read",
        "accounting.invoices.read",
        "accounting.payments.read",
        "accounting.banktransactions.read",
        "accounting.manualjournals.read",
    ]
)

_XERO_MS_DATE = re.compile(r"/Date\((?P<ms>-?\d+)(?:[+-]\d+)?\)/")


def _xero_datetime(value: Any) -> datetime | None:
    """Parse common Xero datetime/date formats into UTC datetimes."""
    if not value:
        return None
    if isinstance(value, datetime):
        return value if value.tzinfo else value.replace(tzinfo=timezone.utc)
    if isinstance(value, str):
        m = _XERO_MS_DATE.fullmatch(value.strip())
        if m:
            return datetime.fromtimestamp(int(m.group("ms")) / 1000, tz=timezone.utc)
        text = value.strip().replace("Z", "+00:00")
        try:
            dt = datetime.fromisoformat(text)
            return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
        except ValueError:
            try:
                return datetime.strptime(text[:10], "%Y-%m-%d").replace(tzinfo=timezone.utc)
            except ValueError:
                return None
    return None


def _decimal(value: Any) -> Decimal | None:
    if value in (None, ""):
        return None
    try:
        return Decimal(str(value))
    except Exception:
        return None


def _safe_name(contact: dict[str, Any] | None) -> str | None:
    if not contact:
        return None
    return contact.get("Name") or contact.get("ContactNumber")


class XeroConnector:
    def __init__(self, store: IntegrationStore):
        self.store = store

    def configured(self) -> bool:
        return bool(
            settings.xero_client_id
            and settings.xero_client_secret
            and settings.xero_redirect_uri
        )

    async def authorization_url(self, company_id: UUID, user_id: UUID) -> str:
        state = await self.store.save_oauth_state(company_id, user_id, "xero")
        params = {
            "response_type": "code",
            "client_id": settings.xero_client_id,
            "redirect_uri": settings.xero_redirect_uri,
            "scope": XERO_SCOPES,
            "state": state,
        }
        return f"{AUTH_URL}?{urlencode(params)}"

    async def callback(self, code: str, state: str):
        ctx = await self.store.consume_oauth_state(state, "xero")
        if not ctx:
            raise ValueError("Invalid or expired Xero OAuth state.")

        async with httpx.AsyncClient(timeout=30) as client:
            token_response = await client.post(
                TOKEN_URL,
                data={
                    "grant_type": "authorization_code",
                    "code": code,
                    "redirect_uri": settings.xero_redirect_uri,
                },
                auth=(settings.xero_client_id or "", settings.xero_client_secret or ""),
            )
            token_response.raise_for_status()
            token_data = token_response.json()

            connection_response = await client.get(
                CONNECTIONS_URL,
                headers={"Authorization": f"Bearer {token_data['access_token']}"},
            )
            connection_response.raise_for_status()
            tenants = connection_response.json()

        chosen = tenants[0] if len(tenants) == 1 else None

        await self.store.upsert_connection(
            company_id=ctx["company_id"],
            provider="xero",
            user_id=ctx["user_id"],
            status="connected" if chosen else "selection_required",
            external_tenant_id=chosen.get("tenantId") if chosen else None,
            external_tenant_name=chosen.get("tenantName") if chosen else None,
            access_token=token_data["access_token"],
            refresh_token=token_data.get("refresh_token"),
            expires_in=token_data.get("expires_in"),
            metadata={
                "tenants": tenants,
                "granted_scope": token_data.get("scope"),
            },
        )
        return ctx["company_id"]

    async def select_tenant(self, company_id: UUID, tenant_id: str):
        connection = await self.store.get(company_id, "xero")
        tenants = (connection or {}).get("metadata", {}).get("tenants", [])
        chosen = next((x for x in tenants if x.get("tenantId") == tenant_id), None)
        if not chosen:
            raise ValueError("That Xero organisation is not available for this connection.")

        await self.store.upsert_connection(
            company_id=company_id,
            provider="xero",
            status="connected",
            external_tenant_id=tenant_id,
            external_tenant_name=chosen.get("tenantName"),
        )

    async def _refresh(self, company_id: UUID, connection: dict):
        if not connection.get("refresh_token"):
            return connection

        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(
                TOKEN_URL,
                data={
                    "grant_type": "refresh_token",
                    "refresh_token": connection["refresh_token"],
                },
                auth=(settings.xero_client_id or "", settings.xero_client_secret or ""),
            )
            response.raise_for_status()
            token_data = response.json()

        await self.store.upsert_connection(
            company_id=company_id,
            provider="xero",
            access_token=token_data["access_token"],
            refresh_token=token_data.get("refresh_token"),
            expires_in=token_data.get("expires_in"),
            metadata={"granted_scope": token_data.get("scope")},
        )
        return await self.store.credentials(company_id, "xero")

    async def _get_all_pages(
        self,
        client: httpx.AsyncClient,
        headers: dict[str, str],
        endpoint: str,
        response_key: str,
        *,
        page_size_supported: bool = True,
        max_pages: int = 500,
    ) -> list[dict[str, Any]]:
        """Retrieve paged Xero endpoints, preserving line-item detail where paging enables it."""
        if not page_size_supported:
            response = await client.get(f"{API_URL}/{endpoint}", headers=headers)
            response.raise_for_status()
            return response.json().get(response_key, [])

        rows: list[dict[str, Any]] = []
        for page in range(1, max_pages + 1):
            response = await client.get(
                f"{API_URL}/{endpoint}",
                headers=headers,
                params={"page": page},
            )
            response.raise_for_status()
            batch = response.json().get(response_key, [])
            if not batch:
                break
            rows.extend(batch)
            # Xero paged Accounting API endpoints return up to 100 records per page.
            if len(batch) < 100:
                break
        return rows

    @staticmethod
    def _account_records(accounts: list[dict[str, Any]]) -> list[dict[str, Any]]:
        records = []
        for account in accounts:
            account_id = account.get("AccountID") or account.get("Code")
            if not account_id:
                continue
            records.append(
                {
                    "external_id": str(account_id),
                    "name": account.get("Name"),
                    "currency_code": account.get("CurrencyCode"),
                    "payload": {
                        **account,
                        "fincruiz": {
                            "record_kind": "chart_of_account",
                            "account_code": account.get("Code"),
                            "account_name": account.get("Name"),
                            "account_type": account.get("Type"),
                            "account_class": account.get("Class"),
                            "status": account.get("Status"),
                            "tax_type": account.get("TaxType"),
                            "system_account": account.get("SystemAccount"),
                            "reporting_code": account.get("ReportingCode"),
                            "reporting_code_name": account.get("ReportingCodeName"),
                        },
                    },
                }
            )
        return records

    @staticmethod
    def _contact_records(contacts: list[dict[str, Any]]) -> list[dict[str, Any]]:
        records = []
        for contact in contacts:
            external_id = contact.get("ContactID")
            if not external_id:
                continue
            records.append(
                {
                    "external_id": external_id,
                    "name": contact.get("Name"),
                    "payload": contact,
                    "source_updated_at": _xero_datetime(contact.get("UpdatedDateUTC")),
                }
            )
        return records

    @staticmethod
    def _invoice_records(invoices: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        documents: list[dict[str, Any]] = []
        ledger_lines: list[dict[str, Any]] = []

        for invoice in invoices:
            invoice_id = invoice.get("InvoiceID")
            if not invoice_id:
                continue

            invoice_type = invoice.get("Type")  # ACCREC / ACCPAY
            contact = invoice.get("Contact") or {}
            occurred_at = _xero_datetime(invoice.get("Date"))
            currency = invoice.get("CurrencyCode")
            documents.append(
                {
                    "external_id": invoice_id,
                    "name": invoice.get("InvoiceNumber") or invoice.get("Reference") or _safe_name(contact),
                    "amount": _decimal(invoice.get("Total")),
                    "currency_code": currency,
                    "occurred_at": occurred_at,
                    "source_updated_at": _xero_datetime(invoice.get("UpdatedDateUTC")),
                    "payload": invoice,
                }
            )

            for idx, line in enumerate(invoice.get("LineItems") or []):
                line_id = line.get("LineItemID") or f"{invoice_id}:{idx + 1}"
                line_amount = _decimal(line.get("LineAmount")) or Decimal("0")
                # Do not label this as a true debit/credit journal line. It is a source transaction line.
                ledger_lines.append(
                    {
                        "external_id": f"invoice:{line_id}",
                        "name": line.get("Description") or invoice.get("InvoiceNumber"),
                        "amount": line_amount,
                        "currency_code": currency,
                        "occurred_at": occurred_at,
                        "source_updated_at": _xero_datetime(invoice.get("UpdatedDateUTC")),
                        "payload": {
                            "source": "invoice",
                            "source_id": invoice_id,
                            "source_number": invoice.get("InvoiceNumber"),
                            "source_type": invoice_type,
                            "status": invoice.get("Status"),
                            "contact_id": contact.get("ContactID"),
                            "contact_name": _safe_name(contact),
                            "reference": invoice.get("Reference"),
                            "date": invoice.get("Date"),
                            "due_date": invoice.get("DueDate"),
                            "currency_code": currency,
                            "currency_rate": invoice.get("CurrencyRate"),
                            "account_code": line.get("AccountCode"),
                            "description": line.get("Description"),
                            "quantity": line.get("Quantity"),
                            "unit_amount": line.get("UnitAmount"),
                            "line_amount": line.get("LineAmount"),
                            "tax_type": line.get("TaxType"),
                            "tax_amount": line.get("TaxAmount"),
                            "discount_rate": line.get("DiscountRate"),
                            "item_code": line.get("ItemCode"),
                            "tracking": line.get("Tracking") or [],
                            "fincruiz": {
                                "record_kind": "transaction_line",
                                "ledger_source": "xero_invoice_line",
                                "is_receivable": invoice_type == "ACCREC",
                                "is_payable": invoice_type == "ACCPAY",
                            },
                        },
                    }
                )

        return documents, ledger_lines

    @staticmethod
    def _bank_transaction_records(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        documents: list[dict[str, Any]] = []
        ledger_lines: list[dict[str, Any]] = []

        for tx in rows:
            tx_id = tx.get("BankTransactionID")
            if not tx_id:
                continue
            contact = tx.get("Contact") or {}
            bank_account = tx.get("BankAccount") or {}
            occurred_at = _xero_datetime(tx.get("Date"))
            currency = tx.get("CurrencyCode")
            documents.append(
                {
                    "external_id": tx_id,
                    "name": tx.get("Reference") or _safe_name(contact) or tx.get("Type"),
                    "amount": _decimal(tx.get("Total")),
                    "currency_code": currency,
                    "occurred_at": occurred_at,
                    "source_updated_at": _xero_datetime(tx.get("UpdatedDateUTC")),
                    "payload": tx,
                }
            )

            for idx, line in enumerate(tx.get("LineItems") or []):
                line_id = line.get("LineItemID") or f"{tx_id}:{idx + 1}"
                ledger_lines.append(
                    {
                        "external_id": f"bank:{line_id}",
                        "name": line.get("Description") or tx.get("Reference"),
                        "amount": _decimal(line.get("LineAmount")),
                        "currency_code": currency,
                        "occurred_at": occurred_at,
                        "source_updated_at": _xero_datetime(tx.get("UpdatedDateUTC")),
                        "payload": {
                            "source": "bank_transaction",
                            "source_id": tx_id,
                            "source_type": tx.get("Type"),
                            "status": tx.get("Status"),
                            "contact_id": contact.get("ContactID"),
                            "contact_name": _safe_name(contact),
                            "reference": tx.get("Reference"),
                            "date": tx.get("Date"),
                            "currency_code": currency,
                            "bank_account_id": bank_account.get("AccountID"),
                            "bank_account_code": bank_account.get("Code"),
                            "bank_account_name": bank_account.get("Name"),
                            "account_code": line.get("AccountCode"),
                            "description": line.get("Description"),
                            "quantity": line.get("Quantity"),
                            "unit_amount": line.get("UnitAmount"),
                            "line_amount": line.get("LineAmount"),
                            "tax_type": line.get("TaxType"),
                            "tax_amount": line.get("TaxAmount"),
                            "tracking": line.get("Tracking") or [],
                            "fincruiz": {
                                "record_kind": "transaction_line",
                                "ledger_source": "xero_bank_transaction_line",
                            },
                        },
                    }
                )

        return documents, ledger_lines

    @staticmethod
    def _manual_journal_records(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        documents: list[dict[str, Any]] = []
        ledger_lines: list[dict[str, Any]] = []

        for journal in rows:
            journal_id = journal.get("ManualJournalID")
            if not journal_id:
                continue
            occurred_at = _xero_datetime(journal.get("Date"))
            documents.append(
                {
                    "external_id": journal_id,
                    "name": journal.get("Narration") or "Manual journal",
                    "occurred_at": occurred_at,
                    "source_updated_at": _xero_datetime(journal.get("UpdatedDateUTC")),
                    "payload": journal,
                }
            )

            for idx, line in enumerate(journal.get("JournalLines") or []):
                line_id = line.get("JournalLineID") or f"{journal_id}:{idx + 1}"
                line_amount = _decimal(line.get("LineAmount")) or Decimal("0")
                ledger_lines.append(
                    {
                        "external_id": f"manualjournal:{line_id}",
                        "name": line.get("Description") or journal.get("Narration"),
                        "amount": line_amount,
                        "occurred_at": occurred_at,
                        "source_updated_at": _xero_datetime(journal.get("UpdatedDateUTC")),
                        "payload": {
                            "source": "manual_journal",
                            "source_id": journal_id,
                            "status": journal.get("Status"),
                            "narration": journal.get("Narration"),
                            "date": journal.get("Date"),
                            "account_code": line.get("AccountCode"),
                            "description": line.get("Description"),
                            "line_amount": line.get("LineAmount"),
                            "tax_type": line.get("TaxType"),
                            "tax_amount": line.get("TaxAmount"),
                            "tracking": line.get("Tracking") or [],
                            "fincruiz": {
                                "record_kind": "transaction_line",
                                "ledger_source": "xero_manual_journal_line",
                                "signed_amount": str(line_amount),
                            },
                        },
                    }
                )

        return documents, ledger_lines

    @staticmethod
    def _payment_records(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        records = []
        for payment in rows:
            payment_id = payment.get("PaymentID")
            if not payment_id:
                continue
            invoice = payment.get("Invoice") or {}
            account = payment.get("Account") or {}
            records.append(
                {
                    "external_id": payment_id,
                    "name": payment.get("Reference") or invoice.get("InvoiceNumber") or "Payment",
                    "amount": _decimal(payment.get("Amount")),
                    "currency_code": payment.get("CurrencyRate") and invoice.get("CurrencyCode"),
                    "occurred_at": _xero_datetime(payment.get("Date")),
                    "source_updated_at": _xero_datetime(payment.get("UpdatedDateUTC")),
                    "payload": {
                        **payment,
                        "fincruiz": {
                            "record_kind": "payment",
                            "invoice_id": invoice.get("InvoiceID"),
                            "invoice_number": invoice.get("InvoiceNumber"),
                            "bank_account_id": account.get("AccountID"),
                            "bank_account_code": account.get("Code"),
                        },
                    },
                }
            )
        return records

    async def sync(self, company_id: UUID):
        connection = await self.store.credentials(company_id, "xero")
        if not connection or not connection.get("external_tenant_id"):
            raise ValueError("Connect and select a Xero organisation first.")

        connection = await self._refresh(company_id, connection)
        headers = {
            "Authorization": f"Bearer {connection['access_token']}",
            "Xero-tenant-id": connection["external_tenant_id"],
            "Accept": "application/json",
        }

        counts: dict[str, int] = {}
        try:
            async with httpx.AsyncClient(timeout=60) as client:
                # Full Chart of Accounts
                accounts_response = await client.get(f"{API_URL}/Accounts", headers=headers)
                accounts_response.raise_for_status()
                accounts = accounts_response.json().get("Accounts", [])
                account_records = self._account_records(accounts)
                await self.store.replace_records(company_id, "xero", "account", account_records)
                counts["accounts"] = len(account_records)

                # Contacts - paged
                contacts = await self._get_all_pages(client, headers, "Contacts", "Contacts")
                contact_records = self._contact_records(contacts)
                await self.store.replace_records(company_id, "xero", "contact", contact_records)
                counts["contacts"] = len(contact_records)

                # Invoices + bills. Paging is important because it returns line detail.
                invoices = await self._get_all_pages(client, headers, "Invoices", "Invoices")
                invoice_records, invoice_lines = self._invoice_records(invoices)
                await self.store.replace_records(company_id, "xero", "invoice", invoice_records)
                counts["invoices"] = len(invoice_records)

                # Bank transactions + their account-coded line items.
                bank_txs = await self._get_all_pages(client, headers, "BankTransactions", "BankTransactions")
                bank_records, bank_lines = self._bank_transaction_records(bank_txs)
                await self.store.replace_records(company_id, "xero", "bank_transaction", bank_records)
                counts["bank_transactions"] = len(bank_records)

                # Manual journals give signed account-level entries and are valuable for ledger reconstruction.
                manual_journals = await self._get_all_pages(client, headers, "ManualJournals", "ManualJournals")
                manual_journal_records, manual_journal_lines = self._manual_journal_records(manual_journals)
                await self.store.replace_records(company_id, "xero", "manual_journal", manual_journal_records)
                counts["manual_journals"] = len(manual_journal_records)

                # Payments are separately useful for cash / AR / AP movement analysis.
                payments_response = await client.get(f"{API_URL}/Payments", headers=headers)
                payments_response.raise_for_status()
                payments = payments_response.json().get("Payments", [])
                payment_records = self._payment_records(payments)
                await self.store.replace_records(company_id, "xero", "payment", payment_records)
                counts["payments"] = len(payment_records)

                # One normalized source-line entity for BI and AI drill-down.
                ledger_lines = invoice_lines + bank_lines + manual_journal_lines
                await self.store.replace_records(company_id, "xero", "ledger_line", ledger_lines)
                counts["ledger_lines"] = len(ledger_lines)

            summary = (
                f"Synced Xero COA ({counts['accounts']}), "
                f"invoices/bills ({counts['invoices']}), bank transactions ({counts['bank_transactions']}), "
                f"manual journals ({counts['manual_journals']}), payments ({counts['payments']}), "
                f"and {counts['ledger_lines']} account-coded source lines."
            )
            await self.store.mark_sync(company_id, "xero", "success", summary)
            return counts

        except httpx.HTTPStatusError as exc:
            body = exc.response.text[:700]
            message = f"Xero API {exc.response.status_code}: {body}"
            await self.store.mark_sync(company_id, "xero", "failed", message)
            raise ValueError(message) from exc
        except Exception as exc:
            await self.store.mark_sync(company_id, "xero", "failed", str(exc))
            raise
