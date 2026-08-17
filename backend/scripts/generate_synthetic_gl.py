"""Generate a balanced synthetic GL CSV for FinCruiz scale testing.

This script only writes a CSV. It never connects to the database.

Examples:
  python scripts/generate_synthetic_gl.py --rows 100000 --output synthetic_100k.csv
  python scripts/generate_synthetic_gl.py --rows 1000000 --branches 5 --output synthetic_1m.csv

Rows are generated as balanced journal pairs. If an odd row count is requested,
one row is dropped so debit == credit for the generated ledger.

Important:
- Use only in a synthetic/test workspace.
- FinCruiz's current browser GL uploader has a 10 MB upload ceiling, so large
  files are intended for benchmark preparation, not immediate browser upload.
"""
from __future__ import annotations

import argparse
import csv
from datetime import date, timedelta
from pathlib import Path
import random


ACCOUNTS = [
    ("4000", "Product Revenue", "credit"),
    ("4010", "Service Revenue", "credit"),
    ("5000", "Cost of Sales", "debit"),
    ("6100", "Payroll Expense", "debit"),
    ("6200", "Rent Expense", "debit"),
    ("6300", "Marketing Expense", "debit"),
    ("6400", "Freight Expense", "debit"),
    ("1000", "Bank Account", "debit"),
    ("1100", "Accounts Receivable", "debit"),
    ("2000", "Accounts Payable", "credit"),
]

HEADERS = [
    "transaction_date",
    "document_number",
    "source_account_code",
    "source_account_name",
    "description",
    "debit",
    "credit",
    "currency_code",
    "branch",
]


def paired_accounts(rng: random.Random):
    operating = rng.choice(ACCOUNTS[:7])
    if operating[2] == "credit":
        counterpart = rng.choice([ACCOUNTS[7], ACCOUNTS[8]])
    else:
        counterpart = rng.choice([ACCOUNTS[7], ACCOUNTS[9]])
    return operating, counterpart


def amounts(account, counterpart, amount: float):
    if account[2] == "credit":
        return (0.0, amount), (amount, 0.0)
    return (amount, 0.0), (0.0, amount)


def generate(*, rows: int, branches: int, output: Path, seed: int, currency: str) -> int:
    if rows < 2:
        raise ValueError("rows must be at least 2")
    rows -= rows % 2
    branches = max(1, branches)
    rng = random.Random(seed)
    branch_names = [f"Branch {i + 1:02d}" for i in range(branches)]
    start = date(2024, 1, 1)

    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS)
        writer.writeheader()

        journals = rows // 2
        for i in range(journals):
            txn_date = start + timedelta(days=i % 730)
            branch = branch_names[i % branches]
            account, counterpart = paired_accounts(rng)
            amount = round(rng.uniform(25.0, 25000.0), 2)
            first, second = amounts(account, counterpart, amount)
            document = f"SYN-{i + 1:09d}"

            writer.writerow({
                "transaction_date": txn_date.isoformat(),
                "document_number": document,
                "source_account_code": account[0],
                "source_account_name": account[1],
                "description": "Synthetic FinCruiz benchmark transaction",
                "debit": f"{first[0]:.2f}",
                "credit": f"{first[1]:.2f}",
                "currency_code": currency,
                "branch": branch,
            })
            writer.writerow({
                "transaction_date": txn_date.isoformat(),
                "document_number": document,
                "source_account_code": counterpart[0],
                "source_account_name": counterpart[1],
                "description": "Synthetic FinCruiz benchmark counterpart",
                "debit": f"{second[0]:.2f}",
                "credit": f"{second[1]:.2f}",
                "currency_code": currency,
                "branch": branch,
            })

    return rows


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, required=True)
    parser.add_argument("--branches", type=int, default=3)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--seed", type=int, default=20260817)
    parser.add_argument("--currency", default="AUD")
    args = parser.parse_args()

    written = generate(
        rows=args.rows,
        branches=args.branches,
        output=args.output,
        seed=args.seed,
        currency=args.currency.upper(),
    )
    print(f"Created {args.output} with {written:,} balanced GL rows.")
