from __future__ import annotations

import asyncio
import csv
import io
import re
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import UUID, uuid4

from fastapi import UploadFile
from sqlalchemy import text
from sqlalchemy.exc import IntegrityError
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.config import settings
from app.core.exceptions import ApplicationError
from app.database.models.core.company import Company
from app.database.session import AsyncSessionLocal
from app.domain.finance.gl_amounts import canonicalise_debit_credit
from app.domain.finance.gl_csv_validator import REQUIRED_GL_COLUMNS, build_column_mapping, detect_delimiter
from app.domain.finance.gl_tabular_reader import ensure_supported_gl_filename, extension_for, xlsx_path_to_csv_path
from app.domain.finance.ingestion.gl_parser import _date, _decimal
from app.repositories.core.branch_repository import BranchRepository
from app.repositories.finance.file_upload_repository import FileUploadRepository
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository

STREAM_CHUNK_BYTES = 1024 * 1024
MAX_ISSUES = 50


def safe_filename(filename: str | None) -> str:
    try:
        return ensure_supported_gl_filename(filename)
    except ValueError as exc:
        raise ApplicationError(
            message=str(exc),
            error_code="UNSUPPORTED_FILE_TYPE",
            status_code=415,
            details={"allowed_extensions": [".csv", ".xlsx"]},
        ) from exc


async def stream_to_staging(file: UploadFile, *, company_id: UUID, job_id: UUID) -> tuple[Path, int]:
    name = safe_filename(file.filename)
    root = Path(settings.import_staging_dir).resolve()
    target_dir = root / str(company_id)
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f"{job_id}_{name}"
    written = 0
    try:
        with target.open("wb") as handle:
            while True:
                chunk = await file.read(STREAM_CHUNK_BYTES)
                if not chunk:
                    break
                written += len(chunk)
                if written > settings.import_max_upload_bytes:
                    raise ApplicationError(
                        message="The file exceeds the configured staged-upload limit.",
                        error_code="FILE_TOO_LARGE",
                        status_code=413,
                        details={"maximum_size_bytes": settings.import_max_upload_bytes},
                    )
                handle.write(chunk)
    except Exception:
        target.unlink(missing_ok=True)
        raise
    if written == 0:
        target.unlink(missing_ok=True)
        raise ApplicationError(message="The uploaded file is empty.", error_code="EMPTY_UPLOAD", status_code=422)
    return target, written


async def create_job(session: AsyncSession, *, company_id: UUID, uploaded_by: UUID, file: UploadFile, source_system: str | None, reporting_period_id: UUID | None) -> dict[str, Any]:
    job_id = uuid4()
    target, size = await stream_to_staging(file, company_id=company_id, job_id=job_id)
    result = await session.execute(text("""
        INSERT INTO public.ingestion_jobs (
            id, company_id, uploaded_by, reporting_period_id, original_file_name,
            staged_path, file_size_bytes, mime_type, source_system, status, phase
        ) VALUES (
            :id, :company_id, :uploaded_by, :reporting_period_id, :name,
            :path, :size, :mime, :source, 'queued', 'queued'
        ) RETURNING *
    """), {
        "id": job_id, "company_id": company_id, "uploaded_by": uploaded_by,
        "reporting_period_id": reporting_period_id, "name": safe_filename(file.filename),
        "path": str(target), "size": size, "mime": file.content_type or "text/csv",
        "source": source_system.strip() if source_system else None,
    })
    await session.commit()
    return dict(result.mappings().one())


async def get_job(session: AsyncSession, *, company_id: UUID, job_id: UUID) -> dict[str, Any] | None:
    result = await session.execute(text("SELECT * FROM public.ingestion_jobs WHERE id=:id AND company_id=:company_id"), {"id": job_id, "company_id": company_id})
    row = result.mappings().one_or_none()
    return dict(row) if row else None


async def list_jobs(session: AsyncSession, *, company_id: UUID, limit: int = 20) -> list[dict[str, Any]]:
    result = await session.execute(text("SELECT * FROM public.ingestion_jobs WHERE company_id=:company_id ORDER BY created_at DESC LIMIT :limit"), {"company_id": company_id, "limit": limit})
    return [dict(row) for row in result.mappings().all()]


def _mapping_for_file(path: Path) -> tuple[str, list[str], dict[str, str], list[str]]:
    with path.open("rb") as raw:
        sample = raw.read(128 * 1024)
    try:
        sample_text = sample.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("The CSV file must use UTF-8 encoding.") from exc
    delimiter = detect_delimiter(sample_text)
    reader = csv.reader(io.StringIO(sample_text), delimiter=delimiter)
    try:
        headers = [str(x or "").strip() for x in next(reader)]
    except StopIteration as exc:
        raise ValueError("CSV header is missing.") from exc
    mapping, _, mapping_issues = build_column_mapping(headers)
    detected = set(mapping.values())
    missing = sorted(REQUIRED_GL_COLUMNS - detected)
    hard_mapping_errors = [issue.message for issue in mapping_issues if issue.severity == "error"]
    if hard_mapping_errors:
        missing.extend(hard_mapping_errors)
    return delimiter, headers, mapping, missing


def _normalise_row(raw_row: dict[str, Any], mapping: dict[str, str], *, row_number: int, company: Company, upload_id: UUID, reporting_period_id: UUID | None) -> dict[str, Any]:
    row = {mapping.get(key, key): (value.strip() if isinstance(value, str) else value) for key, value in raw_row.items() if key}
    source_account_code = str(row.get("source_account_code") or "").strip()
    if not source_account_code:
        raise ValueError("Account code is missing.")
    transaction_date = _date(row.get("transaction_date"), required=True)
    raw_debit = _decimal(row.get("debit"))
    raw_credit = _decimal(row.get("credit"))
    try:
        debit, credit, signed_reversal_normalised = canonicalise_debit_credit(raw_debit, raw_credit)
    except ValueError as exc:
        raise ValueError(f"Row {row_number}: {exc}") from exc
    return {
        "company_id": company.id, "_branch_reference": str(row.get("branch") or "").strip() or None,
        "reporting_period_id": reporting_period_id, "file_upload_id": upload_id,
        "transaction_date": transaction_date, "posting_date": _date(row.get("posting_date")), "document_date": _date(row.get("document_date")),
        "document_number": row.get("document_number") or None, "journal_number": row.get("journal_code") or row.get("journal_number") or None,
        "batch_number": row.get("batch_number") or None, "source_account_code": source_account_code,
        "source_account_name": row.get("source_account_name") or None, "description": row.get("description") or None,
        "reference": row.get("reference") or None, "customer_code": row.get("customer") or row.get("customer_code") or None,
        "supplier_code": row.get("supplier") or row.get("supplier_code") or None, "project_code": row.get("project") or row.get("project_code") or None,
        "cost_centre_code": row.get("cost_centre") or row.get("cost_centre_code") or None, "department_code": row.get("department") or row.get("department_code") or None,
        "debit": debit, "credit": credit, "currency_code": str(row.get("currency_code") or company.currency_code).strip().upper(),
        "exchange_rate": _decimal(row.get("exchange_rate") or "1"), "external_reference": row.get("external_reference") or None,
        "source_row_number": row_number, "validation_status": "valid", "validation_messages": [],
        "source_metadata": {
            "raw_columns": list(raw_row.keys()),
            "signed_reversal_normalised": signed_reversal_normalised,
        },
    }


def validate_stream(path: Path, *, delimiter: str, mapping: dict[str, str], company: Company, upload_id: UUID, reporting_period_id: UUID | None) -> tuple[int, int, list[dict[str, Any]]]:
    total = valid = 0; issues: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle, delimiter=delimiter)
        for row_number, raw in enumerate(reader, start=2):
            if not any(str(v or "").strip() for v in raw.values()): continue
            total += 1
            try:
                _normalise_row(raw, mapping, row_number=row_number, company=company, upload_id=upload_id, reporting_period_id=reporting_period_id)
                valid += 1
            except (ValueError, TypeError) as exc:
                if len(issues) < MAX_ISSUES: issues.append({"row_number": row_number, "column": None, "message": str(exc), "severity": "error"})
    return total, valid, issues


async def _resolve_branches(rows: list[dict[str, Any]], *, company_id: UUID, upload_id: UUID, branch_repository: BranchRepository, mapping: dict[str, Any]) -> None:
    def code_from_value(value: str) -> str:
        return re.sub(r"[^A-Z0-9]+", "", value.upper())[:12] or "BRANCH"
    for row in rows:
        ref = row.pop("_branch_reference", None)
        if not ref: row["branch_id"] = None; continue
        key = ref.strip().lower(); branch = mapping.get(key)
        if branch is None:
            base = code_from_value(ref); code = base; suffix = 2
            while await branch_repository.find_by_code_or_name(company_id, code):
                code = f"{base[:9]}{suffix}"; suffix += 1
            branch = await branch_repository.create({"company_id": company_id, "branch_code": code, "branch_name": ref.strip(), "region": None, "review_status": "pending", "source_value": ref.strip(), "discovered_from_upload_id": upload_id, "is_active": True})
            mapping[key] = branch; mapping[code.lower()] = branch
        row["branch_id"] = branch.id


async def process_job(job_id: UUID) -> None:
    if AsyncSessionLocal is None: return
    async with AsyncSessionLocal() as session:
        try:
            claim = await session.execute(text("""
                UPDATE public.ingestion_jobs
                SET status='processing', phase='validating', progress_percent=5,
                    attempts=attempts+1, started_at=COALESCE(started_at, now()), updated_at=now(), error_message=NULL
                WHERE id=:id AND status IN ('queued','retry') RETURNING *
            """), {"id": job_id})
            job = claim.mappings().one_or_none()
            if not job: return
            await session.commit()
            company = await session.get(Company, job["company_id"])
            if company is None: raise ValueError("Company no longer exists.")
            path = Path(job["staged_path"])
            if not path.exists(): raise ValueError("The staged upload is no longer available. Re-upload the file.")
            working_path = path
            normalised_path: Path | None = None
            if extension_for(job["original_file_name"]) == ".xlsx":
                normalised_path = xlsx_path_to_csv_path(path)
                working_path = normalised_path
            delimiter, headers, mapping, missing = _mapping_for_file(working_path)
            upload_repo = FileUploadRepository(session); tx_repo = GLTransactionRepository(session)
            upload_id = uuid4()
            upload = await upload_repo.create({
                "id": upload_id, "company_id": company.id, "reporting_period_id": job["reporting_period_id"],
                "file_name": f"{upload_id}_{job['original_file_name']}", "original_file_name": job["original_file_name"],
                "storage_bucket": "staged-import", "storage_path": str(path), "mime_type": job["mime_type"], "file_size_bytes": job["file_size_bytes"],
                "document_type": "general_ledger", "source_system": job["source_system"], "processing_status": "processing", "is_active": False,
                "validation_summary": {}, "column_mapping": mapping, "processing_metadata": {"ingestion_mode": "streaming_background", "job_id": str(job_id)}, "uploaded_by": job["uploaded_by"],
            })
            await session.commit()
            if missing:
                raise ValueError("Missing/invalid required columns: " + ", ".join(missing))
            total, valid, issues = validate_stream(working_path, delimiter=delimiter, mapping=mapping, company=company, upload_id=upload_id, reporting_period_id=job["reporting_period_id"])
            invalid = total - valid
            await session.execute(text("UPDATE public.ingestion_jobs SET total_rows=:t,valid_rows=:v,invalid_rows=:i,progress_percent=35,phase='validated',file_upload_id=:u,updated_at=now() WHERE id=:id"), {"t":total,"v":valid,"i":invalid,"u":upload_id,"id":job_id})
            await upload_repo.update(upload, {"row_count": total, "valid_row_count": valid, "invalid_row_count": invalid, "validation_summary": {"total_rows":total,"valid_rows":valid,"invalid_rows":invalid,"issues":issues}, "column_mapping":mapping})
            await session.commit()
            if invalid:
                await upload_repo.update(upload, {"processing_status":"validation_failed", "processed_at":datetime.now(UTC)})
                await session.execute(text("UPDATE public.ingestion_jobs SET status='validation_failed',phase='validation_failed',progress_percent=100,error_message=:e,completed_at=now(),updated_at=now() WHERE id=:id"), {"id":job_id,"e":f"{invalid} row(s) failed validation. Review the upload record for sample issues."})
                await session.commit()
                if normalised_path is not None: normalised_path.unlink(missing_ok=True)
                return
            branch_repo=BranchRepository(session); branch_map=await branch_repo.mapping_by_code_and_name(company.id)
            inserted=0; chunk: list[dict[str,Any]]=[]
            with working_path.open("r",encoding="utf-8-sig",newline="") as handle:
                reader=csv.DictReader(handle,delimiter=delimiter)
                for row_number, raw in enumerate(reader,start=2):
                    if not any(str(v or "").strip() for v in raw.values()): continue
                    chunk.append(_normalise_row(raw,mapping,row_number=row_number,company=company,upload_id=upload_id,reporting_period_id=job["reporting_period_id"]))
                    if len(chunk)>=settings.import_chunk_rows:
                        await _resolve_branches(chunk,company_id=company.id,upload_id=upload_id,branch_repository=branch_repo,mapping=branch_map)
                        inserted += await tx_repo.bulk_create(chunk); await session.commit(); chunk=[]
                        progress=min(95,35+int((inserted/max(total,1))*60))
                        await session.execute(text("UPDATE public.ingestion_jobs SET inserted_rows=:n,progress_percent=:p,phase='importing',updated_at=now() WHERE id=:id"),{"n":inserted,"p":progress,"id":job_id}); await session.commit()
                if chunk:
                    await _resolve_branches(chunk,company_id=company.id,upload_id=upload_id,branch_repository=branch_repo,mapping=branch_map)
                    inserted += await tx_repo.bulk_create(chunk); await session.commit()
            await upload_repo.deactivate_active_datasets(company_id=company.id,document_type="general_ledger",reporting_period_id=job["reporting_period_id"],exclude_upload_id=upload_id)
            await upload_repo.update(upload,{"processing_status":"validated","is_active":True,"processed_at":datetime.now(UTC),"processing_metadata":{**upload.processing_metadata,"gl_transactions_inserted":True,"inserted_transaction_count":inserted,"dataset_status":"active"}})
            await session.execute(text("UPDATE public.ingestion_jobs SET status='completed',phase='completed',progress_percent=100,inserted_rows=:n,completed_at=now(),updated_at=now() WHERE id=:id"),{"n":inserted,"id":job_id})
            await session.commit()
            if normalised_path is not None: normalised_path.unlink(missing_ok=True)
            path.unlink(missing_ok=True)
        except Exception as exc:
            try:
                if 'normalised_path' in locals() and normalised_path is not None:
                    normalised_path.unlink(missing_ok=True)
            except Exception:
                pass
            await session.rollback()
            if isinstance(exc, IntegrityError):
                safe_message = (
                    "The import could not be completed because one or more rows violate FinCruiz canonical ledger rules. "
                    "No new General Ledger dataset was activated. Review debit/credit values and retry."
                )
            elif isinstance(exc, (ValueError, ApplicationError)):
                safe_message = str(exc)[:500] or "Import validation failed."
            else:
                safe_message = (
                    "The import could not be completed. No new General Ledger dataset was activated. "
                    "Please retry or contact support with the import job reference."
                )
            await session.execute(text("UPDATE public.ingestion_jobs SET status='failed',phase='failed',error_message=:e,completed_at=now(),updated_at=now() WHERE id=:id"), {"e":safe_message,"id":job_id})
            await session.commit()


async def claim_next_job() -> UUID | None:
    if AsyncSessionLocal is None: return None
    async with AsyncSessionLocal() as session:
        result = await session.execute(text("""
            SELECT id FROM public.ingestion_jobs
            WHERE status IN ('queued','retry')
            ORDER BY created_at
            FOR UPDATE SKIP LOCKED LIMIT 1
        """))
        row=result.first(); await session.commit(); return row[0] if row else None


async def worker_loop(stop_event: asyncio.Event) -> None:
    while not stop_event.is_set():
        try:
            job_id=await claim_next_job()
            if job_id: await process_job(job_id); continue
        except Exception:
            # Worker health is observable through job status; do not terminate the web app.
            pass
        try: await asyncio.wait_for(stop_event.wait(), timeout=settings.import_worker_poll_seconds)
        except asyncio.TimeoutError: pass
