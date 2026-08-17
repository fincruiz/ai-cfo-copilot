from pathlib import Path
import pytest
from app.services.finance.ingestion_job_service import safe_filename, _mapping_for_file

def test_safe_filename_rejects_non_csv():
    assert safe_filename("ledger.csv") == "ledger.csv"
    with pytest.raises(Exception): safe_filename("ledger.xlsx")

def test_stream_mapping_detects_required_columns(tmp_path: Path):
    p=tmp_path/"ledger.csv"
    p.write_text("transaction_date,source_account_code,debit,credit\n2026-01-01,4000,0,100\n",encoding="utf-8")
    delimiter,headers,mapping,missing=_mapping_for_file(p)
    assert delimiter == ","
    assert missing == []
    assert mapping["source_account_code"] == "source_account_code"
