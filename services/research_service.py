from __future__ import annotations

import json
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from typing import Any
from urllib.parse import urlparse
from urllib.request import Request, urlopen

import pandas as pd

from services.external_data import fetch_country_indicators, get_builtin_industry_benchmarks

TAVILY_ENDPOINT = "https://api.tavily.com/search"
TAVILY_USAGE_ENDPOINT = "https://api.tavily.com/usage"

COUNTRY_DOMAINS: dict[str, list[str]] = {
    "australia": ["abs.gov.au", "rba.gov.au", "treasury.gov.au", "ato.gov.au", "industry.gov.au", "asic.gov.au", "austrade.gov.au"],
    "india": ["rbi.org.in", "mospi.gov.in", "finmin.gov.in", "mca.gov.in", "gst.gov.in", "commerce.gov.in"],
    "united states": ["bea.gov", "bls.gov", "federalreserve.gov", "sec.gov", "census.gov", "irs.gov"],
    "united kingdom": ["ons.gov.uk", "bankofengland.co.uk", "gov.uk", "fca.org.uk"],
    "canada": ["statcan.gc.ca", "bankofcanada.ca", "canada.ca", "osfi-bsif.gc.ca"],
    "new zealand": ["stats.govt.nz", "rbnz.govt.nz", "ird.govt.nz", "mbie.govt.nz"],
}

GLOBAL_AUTHORITY_DOMAINS = {
    "worldbank.org", "imf.org", "oecd.org", "bis.org", "wto.org", "ilo.org",
    "unctad.org", "iea.org", "fao.org", "un.org", "europa.eu",
}

RESEARCH_SCAN_LIBRARY: dict[str, list[dict[str, str]]] = {
    "Executive scan": [
        {"label": "Industry outlook", "intent": "industry outlook demand pricing capacity margins"},
        {"label": "Economic environment", "intent": "inflation interest rates GDP labour market business confidence"},
        {"label": "Cost pressure watch", "intent": "wages freight energy raw materials insurance rent cost pressures"},
        {"label": "Risk and opportunity radar", "intent": "regulation supply chain technology opportunities risks"},
    ],
    "Board pack refresh": [
        {"label": "Market performance", "intent": "market growth demand outlook customer spending latest"},
        {"label": "Peer performance", "intent": "gross margin EBITDA margin working capital benchmarks peer companies"},
        {"label": "Regulatory update", "intent": "latest regulation tax compliance reporting changes"},
        {"label": "Forward indicators", "intent": "leading indicators forecast next 12 months industry"},
    ],
    "Risk watch": [
        {"label": "Supply chain risk", "intent": "supply chain disruption freight logistics availability risk"},
        {"label": "Labour and wage risk", "intent": "labour shortages wage inflation hiring costs"},
        {"label": "Commodity and input risk", "intent": "commodity input prices energy raw material outlook"},
        {"label": "Regulatory and tax risk", "intent": "regulatory tax compliance cyber privacy changes"},
    ],
    "Benchmark scan": [
        {"label": "Profitability benchmarks", "intent": "gross margin EBITDA margin net margin benchmark"},
        {"label": "Working capital benchmarks", "intent": "DSO DPO inventory days cash conversion cycle benchmark"},
        {"label": "Productivity benchmarks", "intent": "revenue per employee labour productivity operating cost benchmark"},
        {"label": "Growth benchmarks", "intent": "revenue growth market growth forecast benchmark"},
    ],
}


def _secret(name: str) -> str:
    value = os.getenv(name, "")
    if value:
        return value
    try:
        import streamlit as st
        return str(st.secrets.get(name, ""))
    except Exception:
        return ""


def tavily_is_configured() -> bool:
    return bool(_secret("TAVILY_API_KEY"))


def tavily_usage() -> dict[str, Any]:
    api_key = _secret("TAVILY_API_KEY")
    if not api_key:
        return {"ok": False, "error": "TAVILY_API_KEY missing"}
    request = Request(
        TAVILY_USAGE_ENDPOINT,
        headers={"Authorization": f"Bearer {api_key}"},
        method="GET",
    )
    try:
        with urlopen(request, timeout=15) as response:
            raw = json.loads(response.read().decode("utf-8"))
        return {"ok": True, **raw}
    except Exception as exc:
        return {"ok": False, "error": str(exc)}


def _domain(url: str) -> str:
    try:
        return urlparse(url).netloc.lower().removeprefix("www.")
    except Exception:
        return ""


def _authority(domain: str, country: str = "") -> str:
    country_domains = COUNTRY_DOMAINS.get(country.lower(), [])
    if any(domain == d or domain.endswith("." + d) for d in country_domains):
        return "Government / regulator"
    if any(domain == d or domain.endswith("." + d) for d in GLOBAL_AUTHORITY_DOMAINS):
        return "Multilateral / official"
    if domain.endswith(".gov") or ".gov." in domain or domain.endswith(".org"):
        return "Institutional"
    return "Market / media"


def search_web(
    query: str,
    *,
    max_results: int = 6,
    search_depth: str = "basic",
    country: str | None = None,
    include_domains: list[str] | None = None,
    exclude_domains: list[str] | None = None,
    topic: str = "general",
    time_range: str | None = None,
) -> dict[str, Any]:
    """Search Tavily and preserve source, authority and retrieval metadata."""
    api_key = _secret("TAVILY_API_KEY")
    retrieved_at = datetime.now(timezone.utc).isoformat()
    if not api_key:
        return {
            "ok": False,
            "query": query,
            "answer": "Live research is not configured. Add TAVILY_API_KEY to .streamlit/secrets.toml and Streamlit Cloud secrets.",
            "results": [],
            "retrieved_at": retrieved_at,
            "error": "TAVILY_API_KEY missing",
        }

    payload: dict[str, Any] = {
        "query": query,
        "search_depth": search_depth if search_depth in {"basic", "advanced", "fast", "ultra-fast"} else "basic",
        "include_answer": "advanced" if search_depth == "advanced" else "basic",
        "include_raw_content": False,
        "max_results": max(1, min(int(max_results), 12)),
        "topic": topic if topic in {"general", "news", "finance"} else "general",
    }
    if country and topic == "general":
        payload["country"] = str(country).lower()
    if include_domains:
        payload["include_domains"] = include_domains[:100]
    if exclude_domains:
        payload["exclude_domains"] = exclude_domains[:100]
    if time_range in {"day", "week", "month", "year", "d", "w", "m", "y"}:
        payload["time_range"] = time_range

    request = Request(
        TAVILY_ENDPOINT,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
        method="POST",
    )

    try:
        with urlopen(request, timeout=45) as response:
            raw = json.loads(response.read().decode("utf-8"))
        normalized = []
        for item in raw.get("results", [])[: payload["max_results"]]:
            url = str(item.get("url") or "")
            domain = _domain(url)
            normalized.append(
                {
                    "title": str(item.get("title") or "Untitled source"),
                    "url": url,
                    "domain": domain,
                    "authority": _authority(domain, country or ""),
                    "content": str(item.get("content") or "")[:2200],
                    "score": float(item.get("score") or 0),
                    "published_date": item.get("published_date") or "",
                }
            )
        return {
            "ok": True,
            "query": query,
            "answer": str(raw.get("answer") or ""),
            "results": normalized,
            "retrieved_at": retrieved_at,
            "response_time": raw.get("response_time"),
            "request_id": raw.get("request_id"),
            "usage": raw.get("usage", {}),
            "error": "",
        }
    except Exception as exc:
        return {
            "ok": False,
            "query": query,
            "answer": "",
            "results": [],
            "retrieved_at": retrieved_at,
            "error": str(exc),
        }


def build_research_plan(profile: dict[str, Any], scan_type: str = "Executive scan") -> list[dict[str, Any]]:
    industry = str(profile.get("Industry") or "business")
    country = str(profile.get("Country") or "Australia")
    company = str(profile.get("Company Name") or "the company")
    period = str(profile.get("Report Period") or "current period")
    items = RESEARCH_SCAN_LIBRARY.get(scan_type, RESEARCH_SCAN_LIBRARY["Executive scan"])
    plan: list[dict[str, Any]] = []
    for item in items:
        plan.append(
            {
                "label": item["label"],
                "query": f"{country} {industry} {item['intent']} {period}",
                "topic": "finance" if "benchmark" in item["label"].lower() or "peer" in item["label"].lower() else "general",
                "time_range": "year",
            }
        )
    if company and company.lower() not in {"the company", "sample manufacturing co."}:
        plan.append(
            {
                "label": "Company and competitor signals",
                "query": f"{company} {country} competitors market news customer demand latest",
                "topic": "news",
                "time_range": "month",
            }
        )
    return plan


def research_company_environment(
    profile: dict[str, Any],
    *,
    search_depth: str = "basic",
    scan_type: str = "Executive scan",
    prefer_authoritative: bool = True,
) -> dict[str, Any]:
    country = str(profile.get("Country") or "Australia")
    plan = build_research_plan(profile, scan_type)
    sections: list[dict[str, Any]] = []

    def run(item: dict[str, Any]) -> dict[str, Any]:
        domains = None
        if prefer_authoritative and item["label"] in {"Economic environment", "Regulatory update", "Regulatory and tax risk"}:
            domains = COUNTRY_DOMAINS.get(country.lower(), []) + sorted(GLOBAL_AUTHORITY_DOMAINS)
        pack = search_web(
            item["query"],
            max_results=7,
            search_depth=search_depth,
            country=country if item["topic"] == "general" else None,
            include_domains=domains,
            topic=item["topic"],
            time_range=item["time_range"],
        )
        pack["label"] = item["label"]
        return pack

    with ThreadPoolExecutor(max_workers=min(4, len(plan))) as pool:
        futures = {pool.submit(run, item): item for item in plan}
        for future in as_completed(futures):
            item = futures[future]
            try:
                sections.append(future.result())
            except Exception as exc:
                sections.append({"ok": False, "label": item["label"], "query": item["query"], "answer": "", "results": [], "error": str(exc)})

    order = {item["label"]: index for index, item in enumerate(plan)}
    sections.sort(key=lambda section: order.get(section.get("label", ""), 999))

    try:
        macro = fetch_country_indicators(country)
    except Exception:
        macro = pd.DataFrame()
    try:
        starter = get_builtin_industry_benchmarks(str(profile.get("Industry") or "Other"), country)
    except Exception:
        starter = pd.DataFrame()

    sources = research_sources_dataframe({"sections": sections})
    authority_count = int(sources["Authority"].isin(["Government / regulator", "Multilateral / official"]).sum()) if not sources.empty else 0
    return {
        "profile": profile,
        "scan_type": scan_type,
        "sections": sections,
        "country_indicators": macro,
        "starter_benchmarks": starter,
        "retrieved_at": datetime.now(timezone.utc).isoformat(),
        "live_search_enabled": tavily_is_configured(),
        "source_count": len(sources),
        "authoritative_source_count": authority_count,
    }


def research_pack_to_context(pack: dict[str, Any] | None, max_chars: int = 24000) -> str:
    if not pack:
        return "No external research snapshot has been generated."
    lines = [
        f"External research retrieved at: {pack.get('retrieved_at', '')}",
        f"Research scan: {pack.get('scan_type', '')}",
        "Important: external sources are contextual research and are not audited financial benchmarks.",
        "Cite source titles and URLs when using external claims.",
    ]
    for section in pack.get("sections", []):
        lines.append(f"\n## {section.get('label', 'Research')}\nQuery: {section.get('query', '')}")
        if section.get("answer"):
            lines.append(f"Search synthesis: {section['answer']}")
        for source in section.get("results", [])[:6]:
            lines.append(
                f"Source: {source.get('title')} | {source.get('url')} | Authority: {source.get('authority')}\n"
                f"Extract: {source.get('content', '')[:850]}"
            )
    macro = pack.get("country_indicators")
    if isinstance(macro, pd.DataFrame) and not macro.empty:
        lines.append("\nCountry indicators:\n" + macro.to_string(index=False))
    starter = pack.get("starter_benchmarks")
    if isinstance(starter, pd.DataFrame) and not starter.empty:
        lines.append("\nStarter benchmarks (directional and must be verified):\n" + starter.to_string(index=False))
    return "\n".join(lines)[:max_chars]


def research_sources_dataframe(pack: dict[str, Any] | None) -> pd.DataFrame:
    rows = []
    for section in (pack or {}).get("sections", []):
        for source in section.get("results", []):
            rows.append(
                {
                    "Research Area": section.get("label", "Research"),
                    "Source": source.get("title", ""),
                    "Domain": source.get("domain", ""),
                    "Authority": source.get("authority", "Market / media"),
                    "URL": source.get("url", ""),
                    "Relevance": round(float(source.get("score") or 0), 3),
                    "Published": source.get("published_date", ""),
                }
            )
    return pd.DataFrame(rows)
