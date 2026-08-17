"""Lightweight authenticated FinCruiz API load test.

Example:
  python scripts/load_test_finance.py --base-url https://api.example.com/api/v1 --token <JWT> --concurrency 20 --requests 200

Use a synthetic/non-production company. This measures API latency; it does not generate GL data.
"""
from __future__ import annotations
import argparse, asyncio, statistics, time
import httpx

DEFAULT_PATHS = ["/reports/profit-and-loss", "/reports/balance-sheet", "/reports/monthly-actuals", "/workspace/status"]

async def hit(client, path, token):
    started=time.perf_counter()
    try:
        r=await client.get(path, headers={"Authorization": f"Bearer {token}"})
        return path, r.status_code, (time.perf_counter()-started)*1000
    except Exception:
        return path, 0, (time.perf_counter()-started)*1000

async def run(args):
    limits=httpx.Limits(max_connections=args.concurrency, max_keepalive_connections=args.concurrency)
    timeout=httpx.Timeout(args.timeout)
    semaphore=asyncio.Semaphore(args.concurrency)
    async with httpx.AsyncClient(base_url=args.base_url.rstrip('/'), limits=limits, timeout=timeout) as client:
        async def one(i):
            async with semaphore: return await hit(client, DEFAULT_PATHS[i % len(DEFAULT_PATHS)], args.token)
        results=await asyncio.gather(*(one(i) for i in range(args.requests)))
    lat=[r[2] for r in results]; ok=sum(1 for _,s,_ in results if 200<=s<300)
    ordered=sorted(lat)
    def pct(p): return ordered[min(len(ordered)-1, int((len(ordered)-1)*p))]
    print(f"requests={len(results)} success={ok} failures={len(results)-ok} concurrency={args.concurrency}")
    print(f"latency_ms mean={statistics.mean(lat):.1f} p50={pct(.50):.1f} p95={pct(.95):.1f} p99={pct(.99):.1f} max={max(lat):.1f}")
    for path in DEFAULT_PATHS:
        vals=[ms for p,_,ms in results if p==path]
        codes={s for p,s,_ in results if p==path}
        print(f"{path}: n={len(vals)} mean={statistics.mean(vals):.1f}ms codes={sorted(codes)}")

if __name__=='__main__':
    ap=argparse.ArgumentParser(); ap.add_argument('--base-url',required=True); ap.add_argument('--token',required=True); ap.add_argument('--concurrency',type=int,default=10); ap.add_argument('--requests',type=int,default=100); ap.add_argument('--timeout',type=float,default=30)
    asyncio.run(run(ap.parse_args()))
