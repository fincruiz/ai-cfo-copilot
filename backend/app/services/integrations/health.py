from __future__ import annotations

from datetime import datetime, timezone


def integration_health(item: dict) -> dict:
    status = item.get("status") or "disconnected"
    sync_status = item.get("last_sync_status")
    last_synced = item.get("last_synced_at")
    if item.get("configured") is False:
        return {"health_status": "configuration_required", "health_message": "Server credentials are not configured.", "recommended_action": "Ask the workspace administrator to configure this provider."}
    if status == "selection_required":
        return {"health_status": "setup_required", "health_message": "Connection authorization succeeded but an organisation still needs to be selected.", "recommended_action": "Choose the organisation to complete setup."}
    if status == "awaiting_bridge":
        return {"health_status": "setup_required", "health_message": "The secure bridge has been created but no Tally data has arrived yet.", "recommended_action": "Start the Tally bridge on the network where TallyPrime is running."}
    if status != "connected":
        return {"health_status": "disconnected", "health_message": "This source is not connected.", "recommended_action": "Connect the source when you are ready."}
    if sync_status in {"failed", "error"}:
        return {"health_status": "failed", "health_message": item.get("last_sync_message") or "The last synchronization failed.", "recommended_action": "Retry sync. If it fails again, reconnect the source or open Support & diagnostics."}
    if not last_synced:
        return {"health_status": "needs_sync", "health_message": "Connected, but FinCruiz has not completed a data sync yet.", "recommended_action": "Run the first sync."}
    try:
        synced = last_synced if isinstance(last_synced, datetime) else datetime.fromisoformat(str(last_synced).replace("Z", "+00:00"))
        if synced.tzinfo is None:
            synced = synced.replace(tzinfo=timezone.utc)
        age_hours = (datetime.now(timezone.utc) - synced).total_seconds() / 3600
        if age_hours > 72:
            return {"health_status": "stale", "health_message": f"Last successful sync was {int(age_hours)} hours ago.", "recommended_action": "Sync now so management views use current source data."}
    except (TypeError, ValueError):
        pass
    return {"health_status": "healthy", "health_message": "Connection and latest sync are healthy.", "recommended_action": "No action required."}
