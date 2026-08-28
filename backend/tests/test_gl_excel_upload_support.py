from io import BytesIO
from pathlib import Path

import pytest
from openpyxl import Workbook

from app.domain.finance.gl_csv_validator import validate_gl_csv
from app.domain.finance.gl_tabular_reader import ensure_supported_gl_filename, normalise_gl_bytes, xlsx_path_to_csv_path


def _workbook_bytes(*, cover_sheet: bool = False) -> bytes:
    workbook = Workbook()
    if cover_sheet:
        cover = workbook.active
        cover.title = "Instructions"
        cover.append(["General Ledger Export"])
        sheet = workbook.create_sheet("GL")
    else:
        sheet = workbook.active
        sheet.title = "GL"
    sheet.append(["General Ledger - FY26"])
    sheet.append(["Date", "Account Code", "Account Name", "Debit", "Credit", "Description"])
    sheet.append(["2026-07-01", "1000", "Bank", 1000, 0, "Opening cash"])
    sheet.append(["2026-07-01", "4000", "Revenue", 0, 1000, "Opening revenue"])
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def test_xlsx_upload_is_normalised_into_existing_finance_truth_validator():
    content = normalise_gl_bytes("general-ledger.xlsx", _workbook_bytes())
    result = validate_gl_csv(content)
    assert result.is_valid is True
    assert result.total_rows == 2
    assert set(result.required_columns) <= set(result.detected_columns)


def test_xlsx_header_detection_can_skip_cover_sheet_and_title_row():
    content = normalise_gl_bytes("ledger.xlsx", _workbook_bytes(cover_sheet=True))
    first_line = content.decode("utf-8-sig").splitlines()[0]
    assert "Date" in first_line
    assert "Account Code" in first_line
    assert validate_gl_csv(content).is_valid is True


def test_xlsx_staging_conversion_works_from_disk(tmp_path: Path):
    source = tmp_path / "ledger.xlsx"
    source.write_bytes(_workbook_bytes())
    target = xlsx_path_to_csv_path(source)
    try:
        assert target.exists()
        assert validate_gl_csv(target.read_bytes()).is_valid is True
    finally:
        target.unlink(missing_ok=True)


def test_renamed_excel_binary_is_not_misread_as_csv():
    with pytest.raises(ValueError, match="Excel workbook"):
        normalise_gl_bytes("ledger.csv", _workbook_bytes())


def test_manual_gl_accepts_csv_and_xlsx_but_rejects_legacy_xls():
    assert ensure_supported_gl_filename("ledger.csv") == "ledger.csv"
    assert ensure_supported_gl_filename("ledger.xlsx") == "ledger.xlsx"
    with pytest.raises(Exception, match="CSV and Excel"):
        ensure_supported_gl_filename("ledger.xls")
