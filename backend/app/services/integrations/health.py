from __future__ import annotations

from datetime import datetime, timezone


def integration_health(item: dict) -> dict:
    status = item.get("status") or "disconnected"
    sync_status = item.get("last_sync_status")
    last_synced = item.get("last_synced_at")
    metadata = item.get("metadata") or {}
    finance_truth = metadata.get("finance_truth") or {}

    if item.get("configured") is False:
        return {
            "health_status": "configuration_required",
            "health_message": "Server credentials are not configured.",
            "recommended_action": "Ask the workspace administrator to configure this provider.",
        }
    if status == "selection_required":
        return {
            "health_status": "setup_required",
            "health_message": "Connection authorization succeeded but an organisation still needs to be selected.",
            "recommended_action": "Choose the organisation to complete setup.",
        }
    if status == "awaiting_bridge":
        return {
            "health_status": "setup_required",
            "health_message": "The secure bridge has been created but no Tally data has arrived yet.",
            "recommended_action": "Start the Tally bridge on the network where TallyPrime is running.",
        }
    if status != "connected":
        return {
            "health_status": "disconnected",
            "health_message": "This source is not connected.",
            "recommended_action": "Connect the source when you are ready.",
        }
    if sync_status in {"failed", "error"}:
        return {
            "health_status": "failed",
            "health_message": item.get("last_sync_message") or "The last synchronization failed.",
            "recommended_action": "Retry sync. If it fails again, reconnect the source or open Support & diagnostics.",
        }
    if not last_synced:
        return {
            "health_status": "needs_sync",
            "health_message": "Connected, but FinCruiz has not completed a data sync yet.",
            "recommended_action": "Run the first sync.",
        }

    try:
        synced = (
            last_synced
            if isinstance(last_synced, datetime)
            else datetime.fromisoformat(str(last_synced).replace("Z", "+00:00"))
        )
        if synced.tzinfo is None:
            synced = synced.replace(tzinfo=timezone.utc)
        age_hours = (datetime.now(timezone.utc) - synced).total_seconds() / 3600
        if age_hours > 72:
            return {
                "health_status": "stale",
                "health_message": f"Last successful sync was {int(age_hours)} hours ago.",
                "recommended_action": "Sync now so management views use current source data.",
            }
    except (TypeError, ValueError):
        pass

    truth_status = finance_truth.get("status")
    if truth_status == "blocked":
        return {
            "health_status": "finance_blocked",
            "health_message": finance_truth.get("message")
            or "Source sync succeeded but the ledger snapshot did not pass finance-truth checks.",
            "recommended_action": "Review the source ledger/reconciliation. FinCruiz has kept the previous active GL unchanged.",
        }
    if truth_status == "source_only":
        return {
            "health_status": "source_only",
            "health_message": finance_truth.get("message")
            or "Source data is synchronized, but this connection is not yet driving financial reports.",
            "recommended_action": "Enable journal-grade ledger access for this provider or load a validated GL file.",
        }
    if truth_status == "collecting":
        return {
            "health_status": "syncing_snapshot",
            "health_message": finance_truth.get("message")
            or "FinCruiz is collecting a complete ledger snapshot.",
            "recommended_action": "Allow the bridge to finish the snapshot before relying on updated reports.",
        }
    if truth_status == "activated":
        rows = finance_truth.get("canonical_rows")
        data_through = finance_truth.get("data_through")
        details = []
        if rows is not None:
            details.append(f"{int(rows):,} GL lines")
        if data_through:
            details.append(f"data through {data_through}")
        suffix = f" ({', '.join(details)})" if details else ""
        return {
            "health_status": "healthy",
            "health_message": f"Connection is healthy and is driving the active FinCruiz ledger{suffix}.",
            "recommended_action": "No action required.",
        }

    return {
        "health_status": "healthy",
        "health_message": "Connection and latest sync are healthy.",
        "recommended_action": "No action required.",
    }
