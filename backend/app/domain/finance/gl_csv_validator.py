import csv
import io
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher


REQUIRED_GL_COLUMNS = {
    "transaction_date",
    "source_account_code",
    "debit",
    "credit",
}

STANDARD_GL_COLUMNS = {
    "transaction_date",
    "posting_date",
    "document_number",
    "reference",
    "journal_code",
    "batch_number",
    "description",
    "source_account_code",
    "source_account_name",
    "debit",
    "credit",
    "currency_code",
    "branch",
    "customer",
    "supplier",
    "project",
    "cost_centre",
}

COLUMN_ALIASES: dict[str, set[str]] = {
    "transaction_date": {
        "date",
        "transaction date",
        "transaction_date",
        "txn date",
        "txn_date",
        "entry date",
        "entry_date",
        "journal date",
        "journal_date",
        "gl date",
        "gl_date",
        "document date",
        "document_date",
        "accounting date",
        "accounting_date",
        "voucher date",
        "voucher_date",
    },
    "posting_date": {
        "posting date",
        "posting_date",
        "posted date",
        "posted_date",
        "post date",
        "post_date",
    },
    "document_number": {
        "document number",
        "document_number",
        "document no",
        "document_no",
        "doc number",
        "doc_number",
        "doc no",
        "doc_no",
        "voucher number",
        "voucher_number",
        "voucher no",
        "voucher_no",
        "invoice number",
        "invoice_number",
        "invoice no",
        "invoice_no",
    },
    "reference": {
        "reference",
        "reference number",
        "reference_number",
        "reference no",
        "reference_no",
        "ref",
        "ref number",
        "ref_number",
        "ref no",
        "ref_no",
    },
    "journal_code": {
        "journal",
        "journal code",
        "journal_code",
        "journal type",
        "journal_type",
        "journal name",
        "journal_name",
    },
    "batch_number": {
        "batch",
        "batch number",
        "batch_number",
        "batch no",
        "batch_no",
    },
    "description": {
        "description",
        "transaction description",
        "transaction_description",
        "narration",
        "narrative",
        "memo",
        "details",
        "particulars",
        "transaction details",
        "transaction_details",
    },
    "source_account_code": {
        "account code",
        "account_code",
        "source account code",
        "source_account_code",
        "gl code",
        "gl_code",
        "gl account",
        "gl_account",
        "gl account code",
        "gl_account_code",
        "ledger code",
        "ledger_code",
        "ledger account",
        "ledger_account",
        "nominal code",
        "nominal_code",
        "nominal account",
        "nominal_account",
        "account number",
        "account_number",
        "account no",
        "account_no",
        "coa code",
        "coa_code",
        "code",
    },
    "source_account_name": {
        "account name",
        "account_name",
        "source account name",
        "source_account_name",
        "gl account name",
        "gl_account_name",
        "ledger name",
        "ledger_name",
        "nominal name",
        "nominal_name",
        "coa name",
        "coa_name",
        "account description",
        "account_description",
    },
    "debit": {
        "debit",
        "debit amount",
        "debit_amount",
        "debits",
        "dr",
        "dr amount",
        "dr_amount",
        "amount dr",
        "amount_dr",
        "debit value",
        "debit_value",
    },
    "credit": {
        "credit",
        "credit amount",
        "credit_amount",
        "credits",
        "cr",
        "cr amount",
        "cr_amount",
        "amount cr",
        "amount_cr",
        "credit value",
        "credit_value",
    },
    "currency_code": {
        "currency",
        "currency code",
        "currency_code",
        "curr",
        "curr code",
        "curr_code",
        "iso currency",
        "iso_currency",
    },
    "branch": {
        "branch",
        "branch name",
        "branch_name",
        "branch code",
        "branch_code",
        "location",
        "location name",
        "location_name",
        "business unit",
        "business_unit",
        "entity",
    },
    "customer": {
        "customer",
        "customer name",
        "customer_name",
        "customer code",
        "customer_code",
        "client",
        "client name",
        "client_name",
    },
    "supplier": {
        "supplier",
        "supplier name",
        "supplier_name",
        "supplier code",
        "supplier_code",
        "vendor",
        "vendor name",
        "vendor_name",
        "vendor code",
        "vendor_code",
    },
    "project": {
        "project",
        "project name",
        "project_name",
        "project code",
        "project_code",
        "job",
        "job name",
        "job_name",
        "job code",
        "job_code",
    },
    "cost_centre": {
        "cost centre",
        "cost_centre",
        "cost center",
        "cost_center",
        "cost centre name",
        "cost_centre_name",
        "cost center name",
        "cost_center_name",
        "department",
        "department name",
        "department_name",
    },
}


@dataclass
class ValidationIssue:
    message: str
    row_number: int | None = None
    column: str | None = None
    severity: str = "error"

    def to_dict(self) -> dict:
        return {
            "row_number": self.row_number,
            "column": self.column,
            "message": self.message,
            "severity": self.severity,
        }


@dataclass
class ColumnMatch:
    original_column: str
    mapped_column: str
    method: str
    confidence: float

    def to_dict(self) -> dict:
        return {
            "original_column": self.original_column,
            "mapped_column": self.mapped_column,
            "method": self.method,
            "confidence": round(self.confidence, 4),
        }


@dataclass
class GLCSVValidationResult:
    required_columns: list[str]
    detected_columns: list[str]
    missing_columns: list[str]
    total_rows: int
    valid_rows: int
    invalid_rows: int
    issues: list[ValidationIssue] = field(default_factory=list)
    column_mapping: dict[str, str] = field(default_factory=dict)
    mapping_details: list[ColumnMatch] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return (
            not self.missing_columns
            and self.invalid_rows == 0
        )

    def to_dict(self) -> dict:
        return {
            "required_columns": self.required_columns,
            "detected_columns": self.detected_columns,
            "missing_columns": self.missing_columns,
            "total_rows": self.total_rows,
            "valid_rows": self.valid_rows,
            "invalid_rows": self.invalid_rows,
            "issues": [
                issue.to_dict()
                for issue in self.issues
            ],
            "column_mapping": self.column_mapping,
            "mapping_details": [
                detail.to_dict()
                for detail in self.mapping_details
            ],
            "is_valid": self.is_valid,
        }


def normalize_column_name(value: str) -> str:
    normalized = value.strip().lower()

    normalized = re.sub(
        r"[\-_./\\]+",
        " ",
        normalized,
    )

    normalized = re.sub(
        r"[^a-z0-9 ]+",
        "",
        normalized,
    )

    return " ".join(normalized.split())


def canonical_to_normalized(
    canonical_name: str,
) -> str:
    return normalize_column_name(
        canonical_name.replace("_", " ")
    )


def build_exact_alias_index() -> dict[str, str]:
    alias_index: dict[str, str] = {}

    for canonical_name in STANDARD_GL_COLUMNS:
        alias_index[
            canonical_to_normalized(canonical_name)
        ] = canonical_name

    for canonical_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            alias_index[
                normalize_column_name(alias)
            ] = canonical_name

    return alias_index


EXACT_ALIAS_INDEX = build_exact_alias_index()


def get_fuzzy_candidates() -> list[tuple[str, str]]:
    candidates: list[tuple[str, str]] = []

    for alias, canonical_name in EXACT_ALIAS_INDEX.items():
        candidates.append(
            (alias, canonical_name)
        )

    return candidates


FUZZY_CANDIDATES = get_fuzzy_candidates()


def find_best_fuzzy_match(
    normalized_header: str,
) -> tuple[str | None, float]:
    best_column: str | None = None
    best_score = 0.0

    for alias, canonical_name in FUZZY_CANDIDATES:
        score = SequenceMatcher(
            None,
            normalized_header,
            alias,
        ).ratio()

        if score > best_score:
            best_score = score
            best_column = canonical_name

    return best_column, best_score


def map_column(
    original_header: str,
) -> ColumnMatch:
    normalized_header = normalize_column_name(
        original_header
    )

    exact_match = EXACT_ALIAS_INDEX.get(
        normalized_header
    )

    if exact_match:
        method = (
            "standard"
            if normalized_header
            == canonical_to_normalized(exact_match)
            else "alias"
        )

        return ColumnMatch(
            original_column=original_header,
            mapped_column=exact_match,
            method=method,
            confidence=1.0,
        )

    fuzzy_match, confidence = find_best_fuzzy_match(
        normalized_header
    )

    if (
        fuzzy_match is not None
        and confidence >= 0.88
    ):
        return ColumnMatch(
            original_column=original_header,
            mapped_column=fuzzy_match,
            method="fuzzy",
            confidence=confidence,
        )

    return ColumnMatch(
        original_column=original_header,
        mapped_column=(
            normalized_header.replace(" ", "_")
            or "unnamed_column"
        ),
        method="unmapped",
        confidence=0.0,
    )


def build_column_mapping(
    headers: list[str],
) -> tuple[
    dict[str, str],
    list[ColumnMatch],
    list[ValidationIssue],
]:
    mapping: dict[str, str] = {}
    details: list[ColumnMatch] = []
    issues: list[ValidationIssue] = []

    mapped_targets: dict[str, str] = {}

    for header in headers:
        match = map_column(header)

        mapping[header] = match.mapped_column
        details.append(match)

        if match.method == "fuzzy":
            issues.append(
                ValidationIssue(
                    column=header,
                    severity="warning",
                    message=(
                        f"Column '{header}' was automatically mapped "
                        f"to '{match.mapped_column}' with "
                        f"{match.confidence:.0%} confidence."
                    ),
                )
            )

        if match.method == "unmapped":
            issues.append(
                ValidationIssue(
                    column=header,
                    severity="warning",
                    message=(
                        f"Column '{header}' was not recognised and "
                        f"will be retained as "
                        f"'{match.mapped_column}'."
                    ),
                )
            )

        previous_header = mapped_targets.get(
            match.mapped_column
        )

        if (
            previous_header is not None
            and match.mapped_column
            in STANDARD_GL_COLUMNS
        ):
            issues.append(
                ValidationIssue(
                    column=header,
                    severity="error",
                    message=(
                        f"Columns '{previous_header}' and '{header}' "
                        f"both map to '{match.mapped_column}'."
                    ),
                )
            )
        else:
            mapped_targets[
                match.mapped_column
            ] = header

    return mapping, details, issues


def parse_decimal(
    value: str | None,
) -> Decimal | None:
    if value is None:
        return None

    cleaned = value.strip()

    if cleaned == "":
        return Decimal("0")

    is_parenthesised = (
        cleaned.startswith("(")
        and cleaned.endswith(")")
    )

    cleaned = (
        cleaned
        .replace(",", "")
        .replace("$", "")
        .replace("₹", "")
        .replace("£", "")
        .replace("€", "")
        .replace("(", "")
        .replace(")", "")
        .strip()
    )

    if cleaned == "":
        return Decimal("0")

    try:
        amount = Decimal(cleaned)

        if is_parenthesised:
            amount = -amount

        return amount

    except InvalidOperation:
        return None


def detect_delimiter(text: str) -> str:
    lines = text.splitlines()

    if not lines:
        return ","

    first_line = lines[0]

    if first_line.count(",") >= 3:
        return ","

    if "\t" in first_line:
        return "\t"

    if ";" in first_line:
        return ";"

    if "|" in first_line:
        return "|"

    try:
        detected_dialect = csv.Sniffer().sniff(
            text[:8192],
            delimiters=",;\t|",
        )

        return detected_dialect.delimiter

    except csv.Error:
        return ","


def validate_gl_csv(
    file_bytes: bytes,
) -> GLCSVValidationResult:
    try:
        text = file_bytes.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError(
            "The CSV file must use UTF-8 encoding."
        ) from exc

    if not text.strip():
        raise ValueError(
            "The CSV file is empty."
        )

    delimiter = detect_delimiter(text)

    stream = io.StringIO(text)

    reader = csv.DictReader(
        stream,
        delimiter=delimiter,
    )

    if not reader.fieldnames:
        raise ValueError(
            "The CSV file does not contain a header row."
        )

    original_headers = [
        header.strip()
        for header in reader.fieldnames
        if header is not None
    ]

    (
        column_mapping,
        mapping_details,
        mapping_issues,
    ) = build_column_mapping(
        original_headers
    )

    detected_columns = list(
        dict.fromkeys(
            column_mapping.values()
        )
    )

    missing_columns = sorted(
        REQUIRED_GL_COLUMNS
        - set(detected_columns)
    )

    issues: list[ValidationIssue] = [
        *mapping_issues,
    ]

    if missing_columns:
        issues.append(
            ValidationIssue(
                message=(
                    "Missing required columns: "
                    + ", ".join(missing_columns)
                )
            )
        )

    total_rows = 0
    valid_rows = 0
    invalid_rows = 0

    has_header_errors = any(
        issue.severity == "error"
        and issue.row_number is None
        for issue in issues
    )

    for source_row_number, raw_row in enumerate(
        reader,
        start=2,
    ):
        if not any(
            str(value or "").strip()
            for value in raw_row.values()
        ):
            continue

        total_rows += 1

        row = {
            column_mapping.get(
                key,
                normalize_column_name(key).replace(
                    " ",
                    "_",
                ),
            ): (
                value.strip()
                if isinstance(value, str)
                else value
            )
            for key, value in raw_row.items()
            if key is not None
        }

        row_issues: list[ValidationIssue] = []

        account_code = row.get(
            "source_account_code"
        )

        if not account_code:
            row_issues.append(
                ValidationIssue(
                    row_number=source_row_number,
                    column="source_account_code",
                    message="Account code is required.",
                )
            )

        transaction_date = row.get(
            "transaction_date"
        )

        if not transaction_date:
            row_issues.append(
                ValidationIssue(
                    row_number=source_row_number,
                    column="transaction_date",
                    message="Transaction date is required.",
                )
            )

        debit = parse_decimal(
            row.get("debit")
        )

        credit = parse_decimal(
            row.get("credit")
        )

        if debit is None:
            row_issues.append(
                ValidationIssue(
                    row_number=source_row_number,
                    column="debit",
                    message="Debit must be numeric.",
                )
            )

        if credit is None:
            row_issues.append(
                ValidationIssue(
                    row_number=source_row_number,
                    column="credit",
                    message="Credit must be numeric.",
                )
            )

        if (
            debit is not None
            and credit is not None
        ):
            if debit < 0:
                row_issues.append(
                    ValidationIssue(
                        row_number=source_row_number,
                        column="debit",
                        message="Debit cannot be negative.",
                    )
                )

            if credit < 0:
                row_issues.append(
                    ValidationIssue(
                        row_number=source_row_number,
                        column="credit",
                        message="Credit cannot be negative.",
                    )
                )

            if debit > 0 and credit > 0:
                row_issues.append(
                    ValidationIssue(
                        row_number=source_row_number,
                        message=(
                            "A row cannot contain both "
                            "a debit and a credit amount."
                        ),
                    )
                )

        if (
            row_issues
            or missing_columns
            or has_header_errors
        ):
            invalid_rows += 1
            issues.extend(row_issues)
        else:
            valid_rows += 1

    if total_rows == 0:
        issues.append(
            ValidationIssue(
                message=(
                    "The CSV file contains no data rows."
                )
            )
        )

    return GLCSVValidationResult(
        required_columns=sorted(
            REQUIRED_GL_COLUMNS
        ),
        detected_columns=detected_columns,
        missing_columns=missing_columns,
        total_rows=total_rows,
        valid_rows=valid_rows,
        invalid_rows=invalid_rows,
        issues=issues,
        column_mapping=column_mapping,
        mapping_details=mapping_details,
    )