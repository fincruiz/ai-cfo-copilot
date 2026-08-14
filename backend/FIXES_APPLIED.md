# FinCruiz Backend – Customer-readiness repair pack

Applied fixes:

1. Secured company listing and company-by-ID routes so authenticated users can only see companies where they have an active membership.
2. Added user+company membership lookup and scoped company listing repository methods.
3. Fixed `CompanyService.update_logo` and `CompanyService.update_company` so they are actual class methods.
4. Added `/auth/refresh` using the Supabase refresh-token flow.
5. Made CORS environment-driven and disabled debug/docs by default in production.
6. Made database initialization tolerate an unset `DATABASE_URL` until a database-backed endpoint is used.
7. Converted `requirements.txt` from UTF-16 to UTF-8.
8. Added `.env.example` and excluded real `.env` from the deliverable.
9. Added tenant/security and company-service regression tests.

Validation performed in the repair environment:

`7 passed` across the original finance/app tests plus the new company security/service tests.

Before production deployment, set real environment variables and run the full test suite in a clean Linux virtual environment.
