from app.schemas.finance.ai_cfo import AIVisualization
from app.services.finance.ai_cfo_service import AICFOService


def _context():
    return {
        "company": {"currency": "AUD"},
        "monthly_actuals": [
            {
                "month": "2026-06-01",
                "revenue": 100000,
                "cost_of_sales": 40000,
                "gross_profit": 60000,
                "operating_expenses": 25000,
                "depreciation": 5000,
                "finance_costs": 2000,
                "tax": 6000,
                "net_profit": 22000,
            },
            {
                "month": "2026-07-01",
                "revenue": 120000,
                "cost_of_sales": 45000,
                "gross_profit": 75000,
                "operating_expenses": 28000,
                "depreciation": 5000,
                "finance_costs": 2000,
                "tax": 8000,
                "net_profit": 32000,
            },
        ],
        "branch_comparison": [
            {"branch_name": "North", "revenue": 70000, "net_profit": 20000},
            {"branch_name": "South", "revenue": 50000, "net_profit": 12000},
        ],
        "balance_sheet": {"total_assets": 500000, "total_liabilities": 220000, "equity": 280000},
    }


def test_profit_question_uses_deterministic_waterfall():
    visual = AICFOService._visualization_for_question("Why did profit change this month?", _context())
    assert visual is not None
    assert visual["type"] == "waterfall"
    assert visual["labels"][0] == "Revenue"
    assert visual["series"][0]["data"][0] == 120000
    AIVisualization(**visual)


def test_branch_question_uses_branch_comparison():
    visual = AICFOService._visualization_for_question("Compare branch performance", _context())
    assert visual is not None
    assert visual["type"] == "bar"
    assert visual["labels"] == ["North", "South"]
    assert visual["series"][0]["data"] == [70000, 50000]
    AIVisualization(**visual)


def test_cost_mix_question_uses_stacked_bar():
    visual = AICFOService._visualization_for_question("Show the expense mix trend", _context())
    assert visual is not None
    assert visual["type"] == "stacked_bar"
    assert len(visual["series"]) == 4
    AIVisualization(**visual)
