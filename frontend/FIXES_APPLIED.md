# FinCruiz Frontend – Customer-readiness repair pack

Applied fixes:

1. Added automatic access-token refresh using `/auth/refresh`.
2. Added single-flight refresh behavior so concurrent 401s do not trigger multiple refresh calls.
3. Retries the original failed request after a successful refresh.
4. Clears the session and returns to login when refresh fails.
5. Added a dashboard-layout authorization guard covering every `/dashboard/*` page.
6. Redirects authenticated users without a company membership to onboarding.
7. Added `.env.example` for `NEXT_PUBLIC_API_URL`.

Run `npm ci`, then `npm run lint` and `npm run build` in your normal development environment before deployment.
