# Stage 8.5B security launch checks

1. Run `python scripts/audit_rls.py` against staging first. It is read-only and flags tenant tables where RLS or policies need review.
2. Do not automatically enable generic policies from a script. Existing application authorization and database policies must be reconciled table-by-table.
3. Confirm service-role keys exist only on the backend/hosting secret store. Never expose them through `NEXT_PUBLIC_*` variables.
4. Verify owner/admin/viewer access against every finance write route using the existing permission tests.
5. Use the Support ID returned on failed API responses to correlate customer reports with backend request logs; do not ask customers to email ledger data to diagnose infrastructure failures.
6. Run the Stage 8.4 load test only with synthetic data and a staging/test company.
