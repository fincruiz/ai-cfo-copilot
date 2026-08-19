from pathlib import Path


def test_dashboard_uses_global_copilot_without_contextual_duplicate():
    contextual = Path(
        "../frontend/components/contextual-ai-bar.tsx"
    ).read_text(encoding="utf-8")

    dashboard_ai = Path(
        "../frontend/components/ask-fincruiz-dashboard.tsx"
    ).read_text(encoding="utf-8")

    # Dashboard must not show the page-scoped contextual assistant.
    assert 'if (pathname === "/dashboard") return null;' in contextual
    assert "Ask about this page" in contextual

    # Dashboard keeps the company-wide global copilot.
    assert "Your global AI CFO copilot" in dashboard_ai
    assert "What should I focus on today?" in dashboard_ai
    assert "Which branch is underperforming?" in dashboard_ai

    # The global copilot must keep the collapsed/docked preference behaviour.
    assert "COLLAPSE_STORAGE_KEY" in dashboard_ai
    assert "localStorage" in dashboard_ai