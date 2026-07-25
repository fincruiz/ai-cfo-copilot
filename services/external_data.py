import json
from urllib.request import Request, urlopen
from urllib.parse import urlencode
import pandas as pd

def get_country_code(country: str) -> str:
    mapping = {
        "Australia": "AUS", "India": "IND", "United States": "USA", "United Kingdom": "GBR",
        "Canada": "CAN", "New Zealand": "NZL",
    }
    return mapping.get(str(country).strip(), "")


def currency_for_country(country: str) -> str:
    mapping = {
        "Australia": "AUD", "India": "INR", "United States": "USD", "United Kingdom": "GBP",
        "Canada": "CAD", "New Zealand": "NZD",
    }
    return mapping.get(str(country).strip(), "USD")


def fetch_json_url(url: str, timeout: int = 12):
    req = Request(url, headers={"User-Agent": "AI-CFO-Copilot/1.0"})
    with urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def fetch_fx_rate(base_currency: str, target_currency: str, date_value: str = "latest") -> dict:
    base_currency = str(base_currency).upper().strip()
    target_currency = str(target_currency).upper().strip()
    date_value = str(date_value or "latest").strip()
    if not base_currency or not target_currency:
        raise ValueError("Base currency and target currency are required.")
    if base_currency == target_currency:
        return {"Base": base_currency, "Target": target_currency, "Rate": 1.0, "Date": date_value, "Source": "Same currency"}
    path_date = "latest" if date_value.lower() in ["", "latest"] else date_value
    url = f"https://api.frankfurter.app/{path_date}?{urlencode({'from': base_currency, 'to': target_currency})}"
    data = fetch_json_url(url)
    rate = data.get("rates", {}).get(target_currency)
    if rate is None:
        raise ValueError(f"FX rate not returned for {base_currency} to {target_currency}.")
    return {"Base": base_currency, "Target": target_currency, "Rate": float(rate), "Date": data.get("date", date_value), "Source": "Frankfurter / ECB"}


def fetch_world_bank_indicator(country_code: str, indicator_code: str, indicator_name: str) -> dict:
    if not country_code:
        raise ValueError("Country code is missing.")
    url = f"https://api.worldbank.org/v2/country/{country_code}/indicator/{indicator_code}?format=json&per_page=8"
    data = fetch_json_url(url)
    rows = data[1] if isinstance(data, list) and len(data) > 1 else []
    for row in rows:
        value = row.get("value")
        if value is not None:
            return {"Indicator": indicator_name, "Code": indicator_code, "Year": row.get("date"), "Value": float(value), "Source": "World Bank"}
    return {"Indicator": indicator_name, "Code": indicator_code, "Year": "N/A", "Value": None, "Source": "World Bank"}


def fetch_country_indicators(country: str) -> pd.DataFrame:
    country_code = get_country_code(country)
    indicators = [
        ("NY.GDP.MKTP.KD.ZG", "GDP growth %"),
        ("FP.CPI.TOTL.ZG", "Inflation %"),
        ("SL.UEM.TOTL.ZS", "Unemployment %"),
        ("NE.EXP.GNFS.ZS", "Exports % of GDP"),
        ("NE.IMP.GNFS.ZS", "Imports % of GDP"),
        ("NV.IND.TOTL.ZS", "Industry value added % of GDP"),
    ]
    rows = [fetch_world_bank_indicator(country_code, code, name) for code, name in indicators]
    return pd.DataFrame(rows)


def get_builtin_industry_benchmarks(industry: str, country: str) -> pd.DataFrame:
    """Starter benchmark set. Users can override with uploaded benchmarks.
    These are broad placeholders for app testing, not official industry benchmarks.
    """
    base = {
        "Manufacturing": {"Gross Margin %": 30, "Operating Margin %": 10, "Opex as % of Revenue": 20, "AR Overdue %": 25, "AP Overdue %": 25},
        "Wholesale / Distribution": {"Gross Margin %": 22, "Operating Margin %": 6, "Opex as % of Revenue": 16, "AR Overdue %": 30, "AP Overdue %": 30},
        "Retail": {"Gross Margin %": 35, "Operating Margin %": 8, "Opex as % of Revenue": 28, "AR Overdue %": 10, "AP Overdue %": 25},
        "Professional Services": {"Gross Margin %": 55, "Operating Margin %": 18, "Opex as % of Revenue": 35, "AR Overdue %": 30, "AP Overdue %": 20},
        "Construction": {"Gross Margin %": 20, "Operating Margin %": 6, "Opex as % of Revenue": 14, "AR Overdue %": 35, "AP Overdue %": 35},
        "Logistics": {"Gross Margin %": 28, "Operating Margin %": 8, "Opex as % of Revenue": 22, "AR Overdue %": 30, "AP Overdue %": 25},
        "Hospitality": {"Gross Margin %": 60, "Operating Margin %": 10, "Opex as % of Revenue": 45, "AR Overdue %": 10, "AP Overdue %": 20},
        "Healthcare": {"Gross Margin %": 40, "Operating Margin %": 12, "Opex as % of Revenue": 30, "AR Overdue %": 25, "AP Overdue %": 20},
        "Technology": {"Gross Margin %": 65, "Operating Margin %": 20, "Opex as % of Revenue": 42, "AR Overdue %": 25, "AP Overdue %": 15},
        "Other": {"Gross Margin %": 30, "Operating Margin %": 10, "Opex as % of Revenue": 25, "AR Overdue %": 25, "AP Overdue %": 25},
    }
    values = base.get(industry, base["Other"])
    return pd.DataFrame([{"Metric": k, "Benchmark Value": v, "Country": country, "Industry": industry, "Source": "Starter benchmark - user should verify/replace"} for k, v in values.items()])


def merge_benchmark_sources(uploaded_df: pd.DataFrame | None, external_df: pd.DataFrame | None) -> pd.DataFrame | None:
    frames = []
    if uploaded_df is not None and not uploaded_df.empty:
        frames.append(uploaded_df[["Metric", "Benchmark Value"]].copy())
    if external_df is not None and not external_df.empty:
        frames.append(external_df[["Metric", "Benchmark Value"]].copy())
    if not frames:
        return None
    merged = pd.concat(frames, ignore_index=True)
    merged = merged.drop_duplicates(subset=["Metric"], keep="first")
    return merged

