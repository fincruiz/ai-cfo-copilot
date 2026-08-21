# Stage 10 — Paid Launch Certification

This runbook is deliberately fail-closed. A green local test suite is **not** proof that the deployed paid product is ready. Operator-evidence timestamps must only be populated after the corresponding real check has passed.

## 1. Freeze the release candidate

- Deploy the exact backend/frontend commit intended for launch.
- Apply all database migrations, including reporting indexes and `20260821_p9_stage9_9_sales_leads.sql`.
- Set `ENVIRONMENT=production` only in the production deployment.
- Keep `BILLING_ALLOW_LIVE_PAYMENTS=false` while the remaining gates are being certified.

## 2. Stripe sandbox lifecycle

Using Stripe **test** credentials and a synthetic company, verify the complete lifecycle end to end through the signed webhook endpoint:

1. checkout session completes;
2. invoice is paid and subscription becomes active;
3. payment-failure event makes the subscription past due;
4. subscription update is reflected correctly;
5. cancellation/deletion makes the subscription cancelled;
6. replay the same webhook and confirm idempotency (no duplicate state transition/event row).

The Stage 10 regression suite validates the canonical transition contract, but the deployed sandbox flow must still be executed. Set `STRIPE_SANDBOX_CERTIFIED_AT` only after the external sandbox flow passes.

## 3. Razorpay sandbox lifecycle

Using Razorpay **test** credentials and a synthetic company, verify:

1. subscription activation;
2. successful charge;
3. halted/payment-failure state;
4. cancellation;
5. payment signature verification;
6. signed webhook verification and duplicate/idempotent handling.

Set `RAZORPAY_SANDBOX_CERTIFIED_AT` only after the external sandbox flow passes.

## 4. Production billing configuration and live switch

- Replace test credentials with the intended live credentials for every provider in `PAID_LAUNCH_PROVIDERS`.
- Configure webhook secrets and all live price/plan identifiers.
- Set `BILLING_FRONTEND_URL` to the public HTTPS frontend.
- Rerun certification while `BILLING_ALLOW_LIVE_PAYMENTS=false` and resolve every non-billing gate first.
- Deliberately set `BILLING_ALLOW_LIVE_PAYMENTS=true` only as the final activation step, then rerun certification. The backend blocks unsafe live credential/switch combinations.

## 5. Production region and performance

- Set `DEPLOYMENT_REGION` and `DATABASE_REGION` to the actual production locations. Do not enter matching values merely to make the gate green.
- Run the authenticated load certification against the deployed production-region API with a synthetic company:

```bash
python -m scripts.load_test_finance \
  --base-url https://<api-host>/api/v1 \
  --token <synthetic-company-jwt> \
  --concurrency 20 \
  --requests 200 \
  --success-target-percent 99 \
  --p95-target-ms 1500
```

- Investigate report/API latency and database path if the target is missed.
- Set `PRODUCTION_PERFORMANCE_CERTIFIED_AT` only after the deployed test passes.

## 6. Persistent ingestion

- Put `IMPORT_STAGING_DIR` on storage that survives process restart/redeploy.
- Upload a representative large GL, interrupt/restart the worker or deployment safely, and verify the queued file remains available and processing can resume/recover.
- Only then set `IMPORT_STAGING_PERSISTENT=true`.

## 7. Backup and restore drill

Do not certify backups by checking that a provider says “backup enabled.” Perform a restore drill:

1. restore the production backup into an isolated non-production target;
2. validate tenant/company isolation;
3. compare transaction counts and `Data as of` for a golden company;
4. run P&L, balance sheet and trial balance checks;
5. confirm integration/import metadata needed for source traceability is present;
6. document restore duration and any manual recovery steps.

Set `BACKUP_RESTORE_VERIFIED_AT` only after the restored data is validated.

## 8. Monitoring and support

- Configure `ERROR_MONITORING_DSN` for the production service.
- Trigger a safe synthetic exception and confirm the alert reaches the monitored workflow.
- Set `ERROR_MONITORING_VERIFIED_AT` only after that delivery/alert check succeeds.
- Configure `SUPPORT_CONTACT_EMAIL` to a monitored address.
- Set `SUPPORT_RUNBOOK_URL` to the current incident/support runbook location.
- The support runbook should cover billing disputes, failed imports, stale ingestion jobs, integration sync failures, tenant/security incidents and finance-output escalation.

## 9. Final certification

Run the full regression suite and the Stage 10 release gate:

```bash
python -m pytest -q
python -m scripts.launch_certify --company-id <golden-production-certification-company-uuid> --frontend-root ../frontend
```

The final script reruns authentication, frontend source surface, tenant/security baseline, finance reliability, ingestion/operations, runtime database latency, billing runtime configuration and all Stage 10 operator-evidence gates. It does not switch billing on and does not mutate customer finance data.

A **GO** is only valid for the exact deployed configuration tested. Any material region, database, billing, authentication or ingestion-storage change requires re-certification.
