"""Authenticated FinCruiz API performance certification.

Example:
  python -m scripts.load_test_finance \
    --base-url https://api.example.com/api/v1 \
    --token <JWT> \
    --concurrency 20 \
    --requests 200 \
    --p95-target-ms 1500 \
    --success-target-percent 99

Use a synthetic/non-production company. No GL data is generated.
Exit codes:
  0 = targets met
  1 = attention (one target missed)
  2 = blocked (severe failure or <95% success)
"""
from __future__ import annotations

import argparse
import asyncio
import statistics
import time

import httpx

DEFAULT_PATHS = [
    "/reports/profit-and-loss",
    "/reports/balance-sheet",
    "/reports/monthly-actuals",
    "/workspace/status",
]


def percentile(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    if not ordered:
        return 0.0
    index = min(len(ordered) - 1, max(0, int((len(ordered) - 1) * fraction)))
    return ordered[index]


def certification_status(*, success_percent: float, p95_ms: float, success_target: float, p95_target_ms: float) -> str:
    if success_percent < 95:
        return "blocked"
    if success_percent < success_target or p95_ms > p95_target_ms:
        return "attention"
    return "ready"


async def hit(client: httpx.AsyncClient, path: str, token: str):
    started = time.perf_counter()
    try:
        response = await client.get(path, headers={"Authorization": f"Bearer {token}"})
        return path, response.status_code, (time.perf_counter() - started) * 1000
    except Exception:
        return path, 0, (time.perf_counter() - started) * 1000


async def run(args) -> int:
    limits = httpx.Limits(
        max_connections=args.concurrency,
        max_keepalive_connections=args.concurrency,
    )
    timeout = httpx.Timeout(args.timeout)
    semaphore = asyncio.Semaphore(args.concurrency)

    async with httpx.AsyncClient(
        base_url=args.base_url.rstrip("/"),
        limits=limits,
        timeout=timeout,
    ) as client:
        async def one(index: int):
            async with semaphore:
                return await hit(client, DEFAULT_PATHS[index % len(DEFAULT_PATHS)], args.token)

        results = await asyncio.gather(*(one(i) for i in range(args.requests)))

    latencies = [item[2] for item in results]
    successes = sum(1 for _, status, _ in results if 200 <= status < 300)
    success_percent = successes / max(len(results), 1) * 100
    p50 = percentile(latencies, 0.50)
    p95 = percentile(latencies, 0.95)
    p99 = percentile(latencies, 0.99)
    state = certification_status(
        success_percent=success_percent,
        p95_ms=p95,
        success_target=args.success_target_percent,
        p95_target_ms=args.p95_target_ms,
    )

    print("\nFinCruiz Performance Certification")
    print("=" * 92)
    print(
        f"requests={len(results)} success={successes} failures={len(results)-successes} "
        f"success_rate={success_percent:.2f}% concurrency={args.concurrency}"
    )
    print(
        f"latency_ms mean={statistics.mean(latencies):.1f} "
        f"p50={p50:.1f} p95={p95:.1f} p99={p99:.1f} max={max(latencies):.1f}"
    )
    print(
        f"targets success>={args.success_target_percent:.2f}% "
        f"p95<={args.p95_target_ms:.0f}ms"
    )
    print("-" * 92)

    for path in DEFAULT_PATHS:
        values = [ms for p, _, ms in results if p == path]
        statuses = sorted({status for p, status, _ in results if p == path})
        successful_path = sum(
            1 for p, status, _ in results if p == path and 200 <= status < 300
        )
        path_rate = successful_path / max(len(values), 1) * 100
        print(
            f"{path:34} n={len(values):4} "
            f"success={path_rate:6.2f}% p95={percentile(values,.95):8.1f}ms "
            f"codes={statuses}"
        )

    print("-" * 92)
    print("PERFORMANCE CERTIFICATION:", state.upper())
    return 0 if state == "ready" else 1 if state == "attention" else 2


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--base-url", required=True)
    parser.add_argument("--token", required=True)
    parser.add_argument("--concurrency", type=int, default=10)
    parser.add_argument("--requests", type=int, default=100)
    parser.add_argument("--timeout", type=float, default=30)
    parser.add_argument("--p95-target-ms", type=float, default=1500)
    parser.add_argument("--success-target-percent", type=float, default=99)
    raise SystemExit(asyncio.run(run(parser.parse_args())))
