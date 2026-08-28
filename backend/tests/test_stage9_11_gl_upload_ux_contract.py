from pathlib import Path


def test_frontend_manual_gl_upload_accepts_xlsx_without_binary_text_preview():
    source = Path("../frontend/app/dashboard/uploads/page.tsx").read_text(encoding="utf-8")
    assert '.csv,.xlsx' in source
    assert 'extension === ".xlsx"' in source
    assert 'Excel workbook ready' in source
    assert 'looksLikeXlsx' in source
    assert 'Choose a CSV or Excel file' in source
