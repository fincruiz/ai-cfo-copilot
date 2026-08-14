"""FinCruiz Tally bridge push helper (P6 Phase 1).

This helper sends a normalized JSON export to the secure FinCruiz Tally endpoint.
It is intentionally transport-only in P6. A later signed Windows bridge will query
TallyPrime directly using the customer's approved Tally XML/JSON/TDL configuration.
"""
from __future__ import annotations
import argparse, json
from pathlib import Path
import httpx


def main() -> None:
    parser=argparse.ArgumentParser()
    parser.add_argument('--api',required=True,help='Example: https://api.example.com/api/v1')
    parser.add_argument('--token',required=True,help='FinCruiz Tally bridge token')
    parser.add_argument('--input',required=True,help='JSON file containing {"records": [...]}')
    args=parser.parse_args()
    payload=json.loads(Path(args.input).read_text(encoding='utf-8'))
    response=httpx.post(args.api.rstrip('/')+'/integrations/tally/push',headers={'Authorization':f'Bearer {args.token}'},json=payload,timeout=60)
    response.raise_for_status()
    print(json.dumps(response.json(),indent=2))

if __name__=='__main__': main()
