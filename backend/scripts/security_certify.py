from __future__ import annotations
import argparse,asyncio
from pathlib import Path
import sys
from uuid import UUID
from sqlalchemy import text
ROOT=Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:sys.path.insert(0,str(ROOT))
from app.database.session import AsyncSessionLocal
TABLES=["companies","company_members","company_invitations","branches","file_uploads","gl_transactions","finance_account_mappings","planning_versions","native_plan_lines","integration_connections","integration_records","audit_events","company_subscriptions","billing_events"]
async def main(cid:UUID):
 if AsyncSessionLocal is None: print("BLOCKED: DATABASE_URL is not configured.");return 2
 async with AsyncSessionLocal() as s:
  rows=(await s.execute(text("SELECT c.relname table_name,c.relrowsecurity rls FROM pg_class c JOIN pg_namespace n ON n.oid=c.relnamespace WHERE n.nspname='public' AND c.relname=ANY(:t)"),{"t":TABLES})).mappings().all()
  state={x["table_name"]:x["rls"] for x in rows};missing=[x for x in TABLES if x not in state];off=[x for x,v in state.items() if not v]
  dup=int((await s.execute(text("SELECT count(*) FROM (SELECT company_id,user_id FROM public.company_members GROUP BY company_id,user_id HAVING count(*)>1)x"))).scalar_one() or 0)
  premature=int((await s.execute(text("SELECT count(*) FROM public.company_invitations i JOIN public.company_members m ON m.company_id=i.company_id AND m.user_id=i.accepted_by WHERE i.company_id=:c AND i.status='accepted' AND m.is_active=true"),{"c":cid})).scalar_one() or 0)
 checks=[("tenant tables present",not missing,str(missing or "all present")),("RLS enabled",not off,str(off or "all enabled")),("membership uniqueness",dup==0,f"{dup} duplicate pair(s)"),("profile-before-access gate",premature==0,f"{premature} premature active membership(s)")]
 print("\nFinCruiz Security Certification\n"+"="*90)
 for label,ok,detail in checks:print(f"{'PASS' if ok else 'FAIL':8} {label:32} {detail}")
 ready=all(x[1] for x in checks);print("-"*90);print("SECURITY CERTIFICATION:","READY" if ready else "BLOCKED");return 0 if ready else 2
if __name__=="__main__":
 p=argparse.ArgumentParser();p.add_argument("--company-id",type=UUID,required=True);raise SystemExit(asyncio.run(main(p.parse_args().company_id)))
