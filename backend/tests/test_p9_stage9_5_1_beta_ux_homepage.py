from pathlib import Path


def test_feedback_list_avoids_ambiguous_null_filters():
    text = Path("app/services/beta_feedback_service.py").read_text(
        encoding="utf-8"
    )
    assert ":status IS NULL" not in text
    assert ":severity IS NULL" not in text
    assert "clauses.append" in text


def test_usage_funnel_uses_typed_interval():
    text = Path("app/api/v1/usage/router.py").read_text(encoding="utf-8")
    assert "make_interval(days => :days)" in text


def test_marketing_telemetry_does_not_accept_free_text_fields():
    text = Path("app/services/marketing_event_service.py").read_text(
        encoding="utf-8"
    )
    lowered = text.lower()

    assert "allowed_events" in lowered

    # These names may appear only as explicit deny/sanitisation keys. They
    # must never be persisted as arbitrary anonymous marketing properties.
    assert '"question"' in lowered
    assert '"email"' in lowered


def test_public_homepage_no_longer_uses_absolute_demo_cards():
    text = Path("../frontend/app/page.tsx").read_text(encoding="utf-8")
    assert "absolute -left" not in text
    assert "absolute -right" not in text


def test_homepage_contains_multiple_capture_points():
    text = Path("../frontend/app/page.tsx").read_text(encoding="utf-8")

    expected_events = {
        "homepage_hero_demo_clicked",
        "homepage_hero_signup_clicked",
        "homepage_ai_question_submitted",
        "homepage_ai_signup_clicked",
        "homepage_reporting_cta_clicked",
        "homepage_forecasting_cta_clicked",
        "homepage_pricing_cta_clicked",
        "homepage_final_demo_clicked",
        "homepage_final_signup_clicked",
    }

    for event_name in expected_events:
        assert event_name in text


def test_ask_fincruiz_collapsed_preference_is_persisted():
    text = Path(
        "../frontend/components/ask-fincruiz-dashboard.tsx"
    ).read_text(encoding="utf-8")

    assert "localStorage" in text
    assert "COLLAPSE_STORAGE_KEY" in text
    assert "setCollapsedPreference" in text
    assert 'useState(true)' in text
    assert "Collapse" in text
    assert "Open" in text


def test_feedback_reporter_is_globally_available():
    layout = Path("../frontend/app/dashboard/layout.tsx").read_text(
        encoding="utf-8"
    )
    reporter = Path(
        "../frontend/components/beta-feedback-button.tsx"
    ).read_text(encoding="utf-8")

    assert "BetaFeedbackButton" in layout
    assert "Feedback" in reporter
