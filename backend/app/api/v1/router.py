from fastapi import APIRouter
from app.api.v1.auth.router import router as auth_router
from app.api.v1.core.companies.router import router as companies_router
from app.api.v1.core.branches.router import router as branches_router
from app.api.v1.core.profile.router import router as profile_router
from app.api.v1.finance.forecasts.router import router as finance_forecasts_router
from app.api.v1.finance.mappings.router import router as finance_mappings_router
from app.api.v1.finance.reports.router import router as finance_reports_router
from app.api.v1.finance.uploads.router import router as finance_uploads_router
from app.api.v1.finance.imports.router import router as finance_imports_router
from app.api.v1.finance.analytics.router import router as finance_analytics_router
from app.api.v1.finance.ai_cfo.router import router as finance_ai_cfo_router
from app.api.v1.finance.planning.router import router as finance_planning_router
from app.api.v1.finance.advanced_forecasting.router import router as advanced_forecasting_router
from app.api.v1.finance.native_planning.router import router as native_planning_router
from app.api.v1.finance.board_packs.router import router as board_packs_router
from app.api.v1.health import router as health_router
from app.api.v1.core.workspace.router import router as workspace_router
from app.api.v1.account.router import router as account_router
from app.api.v1.core.audit.router import router as audit_router
from app.api.v1.integrations.router import router as integrations_router
from app.api.v1.intelligence.router import router as intelligence_router
from app.api.v1.usage.router import router as usage_router
api_router=APIRouter()
api_router.include_router(health_router)
api_router.include_router(auth_router)
api_router.include_router(companies_router)
api_router.include_router(branches_router)
api_router.include_router(profile_router)
api_router.include_router(workspace_router)
api_router.include_router(account_router)
api_router.include_router(audit_router)
api_router.include_router(finance_uploads_router)
api_router.include_router(finance_mappings_router)
api_router.include_router(finance_reports_router)
api_router.include_router(finance_forecasts_router)

api_router.include_router(finance_imports_router)
api_router.include_router(finance_analytics_router)
api_router.include_router(finance_ai_cfo_router)
api_router.include_router(finance_planning_router)

api_router.include_router(advanced_forecasting_router)
api_router.include_router(native_planning_router)
api_router.include_router(board_packs_router)
api_router.include_router(integrations_router)
api_router.include_router(intelligence_router)
api_router.include_router(usage_router)
