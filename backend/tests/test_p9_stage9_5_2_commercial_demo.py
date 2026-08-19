from pathlib import Path

from app.services.demo_ai_service import deterministic_demo_answer


def test_homepage_is_commercially_positioned_and_keeps_conversion_events():
    text = Path("../frontend/app/page.tsx").read_text(encoding="utf-8")

    assert "Turn financial data into" in text
    assert "management decisions" in text
    assert "Owner / CEO" in text
    assert "CFO / Finance" in text
    assert "Accountant / Advisor" in text
    assert "Evidence before AI narrative" in text

    for event_name in {
        "homepage_hero_demo_clicked",
        "homepage_hero_signup_clicked",
        "homepage_ai_question_submitted",
        "homepage_ai_signup_clicked",
        "homepage_reporting_cta_clicked",
        "homepage_forecasting_cta_clicked",
        "homepage_pricing_cta_clicked",
        "homepage_final_demo_clicked",
        "homepage_final_signup_clicked",
    }:
        assert event_name in text


def test_public_metadata_no_longer_uses_create_next_app_defaults():
    text = Path("../frontend/app/layout.tsx").read_text(encoding="utf-8")
    assert "Create Next App" not in text
    assert "AI CFO & Management Intelligence" in text
    assert "management reporting" in text


def test_demo_has_sales_presenter_mode_and_role_specific_story():
    text = Path("../frontend/app/demo/page.tsx").read_text(encoding="utf-8")

    assert "Presenter mode" in text
    assert "Presenter talk track" in text
    assert "Owner / CEO" in text
    assert "CFO / Finance" in text
    assert "Accountant / Advisor" in text
    assert "Pick the problem the prospect already recognises" in text
    assert "demo_presenter_mode_toggled" in text
    assert "demo_scenario_clicked" in text
    assert "demo_question_submitted" in text


def test_demo_growth_scenario_is_modelled_instead_of_falling_into_generic_revenue():
    result = deterministic_demo_answer("What happens if revenue grows 10%?")
    assert "₹27.28M" in result["answer"]
    assert "₹7.42M" in result["answer"]
    assert result["action"]["demo_anchor"] == "forecasting"
    assert result["evidence"]


def test_demo_forecast_question_has_prepared_evidence():
    result = deterministic_demo_answer("Build a 12-month forecast.")
    assert "₹26.04M" in result["answer"]
    assert "₹6.78M" in result["answer"]
    assert result["action"]["demo_anchor"] == "forecasting"


def test_demo_board_question_routes_to_board_story():
    result = deterministic_demo_answer("What should the board discuss next?")
    assert "board" in result["answer"].lower()
    assert result["action"]["demo_anchor"] == "board"


def test_demo_refuses_unprepared_growth_percentage():
    result = deterministic_demo_answer("What happens if revenue grows 23%?")
    assert "will not fabricate" in result["answer"]
    assert not result["evidence"]


def test_marketing_allowlist_matches_new_homepage_and_demo_events():
    text = Path("app/services/marketing_event_service.py").read_text(encoding="utf-8")
    for event_name in {
        "homepage_pricing_cta_clicked",
        "homepage_reporting_cta_clicked",
        "homepage_forecasting_cta_clicked",
        "homepage_persona_changed",
        "demo_viewed",
        "demo_audience_changed",
        "demo_presenter_mode_toggled",
        "demo_guided_scene_clicked",
        "demo_scenario_clicked",
        "demo_question_submitted",
        "demo_signup_clicked",
    }:
        assert event_name in text

    lowered = text.lower()
    assert '"question"' in lowered
    assert '"email"' in lowered
