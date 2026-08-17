from __future__ import annotations

from typing import Any


def build_launch_readiness(*, company: Any, workspace: dict, connections: list[dict]) -> dict:
    profile_complete = bool(
        getattr(company, "legal_name", None)
        and getattr(company, "country_code", None)
        and getattr(company, "currency_code", None)
        and getattr(company, "industry", None)
        and getattr(company, "business_model", None)
    )
    connected = [item for item in connections if item.get("status") == "connected"]
    healthy_connected = [item for item in connected if item.get("last_sync_status") not in {"failed", "error"}]
    has_source = bool(workspace.get("has_financial_data") or connected)
    mapping_ready = bool(workspace.get("mapping_count", 0) > 0)
    insights_ready = bool(workspace.get("transaction_count", 0) > 0 and mapping_ready)

    checks = [
        {"key": "profile", "label": "Business profile", "ready": profile_complete, "detail": "Country, currency, industry and business model are set.", "path": "/dashboard/profile"},
        {"key": "source", "label": "Business data", "ready": has_source, "detail": "Upload finance data or connect an accounting source.", "path": "/dashboard/integrations" if not workspace.get("upload_count") else "/dashboard/uploads"},
        {"key": "mapping", "label": "Account mapping", "ready": mapping_ready, "detail": "Mapped accounts make reports, BI and AI financially consistent.", "path": "/dashboard/mapping"},
        {"key": "intelligence", "label": "Management intelligence", "ready": insights_ready, "detail": "Enough mapped transaction data is available for management insights.", "path": "/dashboard/intelligence"},
    ]
    completed = sum(1 for check in checks if check["ready"])
    next_check = next((check for check in checks if not check["ready"]), None)
    return {
        "score": int(round(completed / len(checks) * 100)),
        "completed_steps": completed,
        "total_steps": len(checks),
        "checks": checks,
        "next_path": next_check["path"] if next_check else "/dashboard/intelligence",
        "next_label": next_check["label"] if next_check else "Explore management intelligence",
        "connected_sources": len(connected),
        "healthy_sources": len(healthy_connected),
        "ready_for_management_use": insights_ready,
    }
