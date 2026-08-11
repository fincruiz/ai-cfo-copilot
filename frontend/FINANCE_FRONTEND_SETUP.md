# FinCruiz finance frontend

## Local setup

1. Copy `.env.local.example` to `.env.local`.
2. Start backend on `http://127.0.0.1:8000`.
3. Run `npm install` and `npm run dev` in the frontend folder.
4. Open `http://localhost:3000/login`.

## End-to-end flow

1. Login.
2. Upload a general ledger CSV under **Upload GL**.
3. Review and save AI suggestions under **Mapping**.
4. Open **Reports** and **KPIs**.

The frontend uses the bearer token stored after login and calls the FastAPI routes under `/api/v1`.
