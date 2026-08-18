from __future__ import annotations

import json
import re
import httpx

from app.core.config import settings

DEMO_CONTEXT={
 "company":"Nova Retail","currency":"INR","period":"FY2026 synthetic demo",
 "revenue":24800000,"net_profit":4120000,"cash":6210000,"gross_margin":42.4,
 "overdue_ar":1180000,"overdue_ar_percent":28.0,"debtor_days":54,
 "branches":[
   {"name":"Central","revenue":11800000,"gp_percent":44.8,"net_profit":2050000},
   {"name":"North","revenue":7400000,"gp_percent":41.6,"net_profit":1240000},
   {"name":"West","revenue":5600000,"gp_percent":36.7,"net_profit":830000},
 ],
 "monthly_revenue":[1.62,1.68,1.75,1.82,1.89,1.94,2.02,2.08,2.15,2.22,2.31,2.42],
 "monthly_cash":[5.9,5.7,5.5,5.4,5.1,4.9,4.8,4.7,4.9,5.2,5.7,6.21],
 "scenario":{"hire_3":{"closing_cash":4080000,"profit":3460000,"buffer":3500000},"downside":{"closing_cash":2920000}},
}

SUGGESTIONS=["Why is profit up but cash tight?","Which branch needs attention?","Can we afford to hire 3 people?","What are the biggest risks management should discuss?","Show the revenue trend","How can we improve working capital?"]

def _chart(title:str,labels:list[str],series:list[dict],fmt:str="currency"):
 return {"type":"line","title":title,"subtitle":"Synthetic demo data","labels":labels,"series":series,"value_format":fmt,"currency":"INR"}

def deterministic_demo_answer(question:str)->dict:
 q=question.lower(); ev=[]; action=None; viz=None
 def evidence(label,value,source="Synthetic management data"):ev.append({"label":label,"value":value,"source":source})
 if any(x in q for x in ("cash","working capital","receivable","debtor","collection")):
  answer="Cash is tighter than profit suggests because receivables are absorbing working capital. The synthetic business has ₹1.18M overdue in AR and debtor days of 54, so reported profit is converting to cash more slowly than management would want."
  evidence("Overdue AR","₹1.18M","Synthetic AR ageing");evidence("Debtor days","54 days","Synthetic working-capital KPI");evidence("Cash","₹6.21M","Synthetic balance sheet")
  viz=_chart("Revenue vs cash conversion",[f"M{i}" for i in range(1,13)],[{"name":"Revenue","data":DEMO_CONTEXT["monthly_revenue"]},{"name":"Cash","data":DEMO_CONTEXT["monthly_cash"]}])
  action={"label":"Explore Working Capital","demo_anchor":"working-capital"}
 elif any(x in q for x in ("branch","west","location","site")):
  b=DEMO_CONTEXT["branches"]; weak=min(b,key=lambda x:x["gp_percent"])
  answer=f"West needs the closest management attention. It is growing, but its {weak['gp_percent']:.1f}% gross margin is materially below Central at 44.8%, so the issue is margin quality rather than simply revenue volume."
  evidence("West GP%","36.7%","Synthetic branch P&L");evidence("Central GP%","44.8%","Synthetic branch P&L")
  viz={"type":"bar","title":"Branch gross margin","subtitle":"Synthetic branch comparison","labels":[x["name"] for x in b],"series":[{"name":"GP%","data":[x["gp_percent"] for x in b]}],"value_format":"percent","currency":"INR"}
  action={"label":"Compare branch drivers","demo_anchor":"branches"}
 elif any(x in q for x in ("hire","can we afford","afford to hire","add staff","new staff","additional people")):
  answer="In the base hiring scenario, three additional people are affordable: projected closing cash is ₹4.08M, still above the ₹3.5M management buffer. The downside case is the risk—slower collections would push closing cash to about ₹2.92M."
  evidence("Hiring closing cash","₹4.08M","Synthetic three-way forecast");evidence("Management buffer","₹3.50M","Synthetic policy assumption");evidence("Downside closing cash","₹2.92M","Synthetic downside scenario")
  action={"label":"Explore Decision Simulation","demo_anchor":"decision"}
 elif any(x in q for x in ("revenue","sales","growth","trend")):
  answer="Revenue is trending upward across the synthetic 12-month period, reaching about ₹2.42M in the latest month. Management should still monitor margin and cash conversion so growth does not hide deteriorating quality."
  evidence("Annual revenue","₹24.80M","Synthetic P&L");evidence("Latest monthly revenue","₹2.42M","Synthetic monthly actuals")
  viz=_chart("Revenue trend",[f"M{i}" for i in range(1,13)],[{"name":"Revenue (₹M)","data":DEMO_CONTEXT["monthly_revenue"]}])
 elif any(x in q for x in ("profit","margin","gross")):
  answer="The synthetic company is profitable at ₹4.12M net profit with a 42.4% gross margin. The main management question is quality by branch: West's lower margin is diluting an otherwise healthy group result."
  evidence("Net profit","₹4.12M","Synthetic P&L");evidence("Gross margin","42.4%","Synthetic P&L");evidence("West GP%","36.7%","Synthetic branch P&L")
 elif any(x in q for x in ("risk","priority","focus","board","management")):
  answer="The three highest-priority issues are working-capital conversion, West branch margin leakage, and downside cash resilience if collections slow. Growth is positive, but management should protect cash quality before accelerating fixed-cost commitments."
  evidence("Overdue AR","₹1.18M","Synthetic AR ageing");evidence("West GP%","36.7%","Synthetic branch P&L");evidence("Downside cash","₹2.92M","Synthetic scenario")
 else:
  answer="The demo can analyse profitability, cash and working capital, branch performance, trends, risks, and decision scenarios from the synthetic Nova Retail dataset. This question needs evidence outside that demo dataset, so I won't invent an answer. Try a management question about revenue, margin, cash, branches, hiring, working capital or risk."
 return {"answer":answer,"mode":"synthetic_demo","evidence":ev,"confidence":"high" if ev else "medium","confidence_reason":"This answer is grounded only in the fixed synthetic Nova Retail demo dataset; no customer data is used.","suggested_questions":SUGGESTIONS[:4],"visualization":viz,"action":action}

async def answer_demo_question(question:str)->dict:
 fallback=deterministic_demo_answer(question)
 if not settings.openai_api_key: return fallback
 # Public demo AI receives only synthetic aggregates. It never receives customer/company context.
 prompt={"question":question,"synthetic_context":DEMO_CONTEXT,"rules":["Use only synthetic_context.","Do not claim access to live customer data.","If evidence is missing, say so.","Keep the answer under 130 words.","Use management language, not accounting jargon."]}
 try:
  payload={"model":settings.openai_model,"input":[{"role":"system","content":[{"type":"input_text","text":"You are the public FinCruiz product demo. Answer only from the supplied synthetic company context. Never invent missing evidence."}]},{"role":"user","content":[{"type":"input_text","text":json.dumps(prompt)}]}],"max_output_tokens":260}
  async with httpx.AsyncClient(timeout=18) as client:
   r=await client.post("https://api.openai.com/v1/responses",json=payload,headers={"Authorization":f"Bearer {settings.openai_api_key}","Content-Type":"application/json"})
  if r.status_code>=400:return fallback
  data=r.json(); parts=[]
  for item in data.get("output",[]):
   for content in item.get("content",[]):
    if content.get("type") in {"output_text","text"} and content.get("text"):parts.append(str(content["text"]))
  answer="\n".join(parts).strip()
  if answer: return {**fallback,"answer":answer,"mode":"synthetic_demo_ai"}
 except Exception: pass
 return fallback
