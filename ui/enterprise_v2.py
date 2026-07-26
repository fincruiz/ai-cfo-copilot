from __future__ import annotations

from datetime import datetime
from html import escape
from typing import Iterable

import streamlit as st


PAGE_CONFIG = [
    ("dashboard", "Dashboard", "⌂", "Executive command centre"),
    ("upload", "Import Centre", "⇧", "Load and validate finance data"),
    ("home", "Company & Setup", "⚙", "Branding, profile and workspace controls"),
    ("analytics", "Analytics", "◫", "Trends, variance and drivers"),
    ("reports", "Reports", "▤", "Statements and KPI packs"),
    ("ratios", "Ratios", "◒", "Liquidity, leverage and productivity"),
    ("board_report", "Board Report", "▣", "AI-supported board and management pack"),
    ("board_meeting", "Board Meeting", "▶", "Interactive director presentation mode"),
    ("working_capital", "Working Capital", "◎", "Receivables and payables"),
    ("insights", "AI Insights", "✦", "Commentary and recommendations"),
    ("research", "Market Intelligence", "⌁", "External signals and benchmarks"),
    ("downloads", "Downloads", "⇩", "Export management packs"),
]

def inject_enterprise_v2_css() -> None:
    st.markdown(
        """
<style>
:root{
 --bg:#111827;--panel:#182235;--panel2:#202c42;--panel3:#293750;
 --text:#f7f9ff;--muted:#91a0ba;--line:rgba(148,163,184,.16);
 --blue:#5b8cff;--cyan:#35d4e8;--violet:#8a6dff;--green:#39d98a;
 --amber:#ffbf5b;--red:#ff6b7a;--radius:20px;
}
html,body,[class*="css"]{font-family:Inter,ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif!important}
html,body{overflow-x:hidden;background:var(--bg)}
.stApp{
 background:
 radial-gradient(circle at 80% -10%,rgba(91,140,255,.13),transparent 34%),
 radial-gradient(circle at 15% 105%,rgba(53,212,232,.08),transparent 28%),
 linear-gradient(180deg,#162135 0%,#111827 100%)!important;
 color:var(--text);
}
[data-testid="stHeader"]{background:transparent!important;height:2.2rem}
[data-testid="stToolbar"]{display:none!important}
[data-testid="stSidebar"]{
 display:block!important;width:260px!important;min-width:260px!important;
 background:linear-gradient(180deg,rgba(28,39,58,.99),rgba(20,29,45,.99))!important;
 border-right:1px solid var(--line)!important;
}
[data-testid="stSidebar"]>div:first-child{padding:1.1rem .85rem 1rem!important}
[data-testid="collapsedControl"], [data-testid="stSidebarCollapseButton"], button[kind="header"]{display:none!important}
/* Keep navigation mounted after Streamlit reruns and page clicks. */
section[data-testid="stSidebar"], [data-testid="stSidebar"]{
 display:block!important;visibility:visible!important;opacity:1!important;
 transform:none!important;left:0!important;margin-left:0!important;
 width:260px!important;min-width:260px!important;max-width:260px!important;
}
section[data-testid="stSidebar"][aria-expanded="false"]{transform:none!important;min-width:260px!important;max-width:260px!important;}
[data-testid="stAppViewContainer"]{margin-left:0!important;}
.main .block-container{max-width:1500px!important;padding:1.2rem 2rem 5rem!important}
h1,h2,h3,h4{color:var(--text)!important;letter-spacing:-.035em}
h1{font-size:2.05rem!important;font-weight:800!important} h2{font-size:1.5rem!important} h3{font-size:1.1rem!important}
p,label,.stCaption{color:var(--muted)!important}

/* global motion */
.main .block-container>div{animation:v2PageIn .45s cubic-bezier(.2,.8,.2,1) both}
@keyframes v2PageIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
@keyframes v2Glow{0%,100%{opacity:.6;transform:scale(.96)}50%{opacity:1;transform:scale(1.04)}}
@keyframes v2Shimmer{from{background-position:200% 0}to{background-position:-200% 0}}
@keyframes v2Pulse{0%,100%{box-shadow:0 0 0 0 rgba(91,140,255,.25)}50%{box-shadow:0 0 0 12px rgba(91,140,255,0)}}

/* sidebar */
.v2-brand{display:flex;align-items:center;gap:.75rem;padding:.25rem .25rem 1rem;border-bottom:1px solid var(--line);margin-bottom:.85rem}
.v2-brand-mark{width:42px;height:42px;border-radius:14px;display:grid;place-items:center;background:linear-gradient(135deg,var(--cyan),var(--blue),var(--violet));font-weight:900;color:white;box-shadow:0 12px 30px rgba(91,140,255,.25)}
.v2-brand-name{font-size:1rem;font-weight:850;color:white}.v2-brand-sub{font-size:.69rem;color:#71809a;margin-top:.08rem}
.v2-workspace{padding:.78rem;border-radius:15px;background:rgba(255,255,255,.035);border:1px solid var(--line);margin-bottom:.8rem}
.v2-workspace-label{font-size:.63rem;text-transform:uppercase;letter-spacing:.1em;color:#71809a;font-weight:800}
.v2-workspace-name{font-weight:800;color:#f6f8ff;margin-top:.22rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.v2-workspace-period{font-size:.72rem;color:#8290a9;margin-top:.15rem}
.v2-nav-heading{font-size:.62rem;text-transform:uppercase;letter-spacing:.11em;color:#66748d;font-weight:850;margin:.95rem .35rem .4rem}
[data-testid="stSidebar"] div[data-testid="stButton"]{margin:.12rem 0}
[data-testid="stSidebar"] div[data-testid="stButton"]>button{
 min-height:2.68rem!important;border-radius:12px!important;border:1px solid transparent!important;
 background:transparent!important;color:#a9b5c9!important;justify-content:flex-start!important;
 padding:.52rem .72rem!important;font-weight:700!important;box-shadow:none!important;transition:.18s ease!important;
}
[data-testid="stSidebar"] div[data-testid="stButton"]>button:hover{background:rgba(91,140,255,.09)!important;color:#fff!important;border-color:rgba(91,140,255,.18)!important;transform:translateX(2px)}
[data-testid="stSidebar"] div[data-testid="stButton"].v2-active-nav>button,
[data-testid="stSidebar"] .st-key-v2_active_nav button{background:linear-gradient(90deg,rgba(91,140,255,.20),rgba(138,109,255,.10))!important;color:#fff!important;border-color:rgba(91,140,255,.30)!important;box-shadow:inset 3px 0 0 var(--blue)!important}
.v2-sidebar-foot{margin-top:1rem;padding:.8rem;border-radius:14px;border:1px solid var(--line);background:rgba(255,255,255,.025)}
.v2-status-row{display:flex;align-items:center;justify-content:space-between;font-size:.72rem;color:#8c9ab2}.v2-status-dot{width:7px;height:7px;border-radius:50%;background:var(--green);box-shadow:0 0 12px rgba(57,217,138,.7)}

/* topbar and headers */
.v2-topbar{display:flex;justify-content:space-between;align-items:center;gap:1rem;padding:.72rem .9rem;margin-bottom:1.15rem;border-radius:16px;border:1px solid var(--line);background:rgba(29,40,59,.86);backdrop-filter:blur(18px);box-shadow:0 12px 35px rgba(0,0,0,.14);position:relative;overflow:hidden}
.v2-topbar:after{content:"";position:absolute;inset:0;background:linear-gradient(90deg,transparent,rgba(255,255,255,.025),transparent);background-size:200% 100%;animation:v2Shimmer 7s linear infinite;pointer-events:none}
.v2-breadcrumb{font-size:.73rem;color:#7887a0;font-weight:700}.v2-breadcrumb b{color:#cbd5e1}
.v2-top-actions{display:flex;gap:.55rem;align-items:center}.v2-live{display:flex;gap:.4rem;align-items:center;padding:.38rem .62rem;border:1px solid rgba(57,217,138,.2);border-radius:999px;background:rgba(57,217,138,.07);font-size:.7rem;color:#8fe8bb;font-weight:800}
.v2-avatar{width:34px;height:34px;border-radius:11px;display:grid;place-items:center;background:linear-gradient(135deg,#26395f,#5849a7);color:#fff;font-size:.78rem;font-weight:850}
.v2-page-head{display:flex;align-items:flex-end;justify-content:space-between;gap:1rem;margin:.3rem 0 1rem}.v2-page-title{font-size:2rem;font-weight:850;color:#fff;letter-spacing:-.05em}.v2-page-sub{font-size:.88rem;color:#8290aa;margin-top:.22rem}.v2-period-pill{padding:.45rem .68rem;border:1px solid var(--line);border-radius:10px;background:rgba(255,255,255,.03);font-size:.73rem;color:#9aa8bf;font-weight:750}

.v6-welcome{display:flex;justify-content:space-between;align-items:center;gap:1rem;margin:.2rem 0 1.1rem;padding:1.25rem 1.35rem;border-radius:22px;background:linear-gradient(120deg,rgba(62,91,137,.50),rgba(50,63,112,.34),rgba(61,112,132,.26));border:1px solid rgba(172,196,230,.28);box-shadow:0 18px 45px rgba(2,8,20,.18);position:relative;overflow:hidden;animation:v6Welcome .55s cubic-bezier(.2,.8,.2,1) both}.v6-welcome:after{content:"";position:absolute;width:260px;height:260px;border-radius:50%;right:-90px;top:-150px;background:rgba(98,140,255,.20);filter:blur(20px);animation:v6Orb 7s ease-in-out infinite alternate}.v6-greeting{font-size:1.75rem;color:#fff;font-weight:900;letter-spacing:-.045em}.v6-welcome-copy{color:#d7e1ef;margin-top:.25rem}.v6-date{position:relative;z-index:2;color:#dbe8fb;font-weight:800;padding:.5rem .75rem;border-radius:12px;background:rgba(5,12,24,.24);border:1px solid rgba(203,218,239,.18)}@keyframes v6Welcome{from{opacity:0;transform:translateY(-12px) scale(.985)}to{opacity:1;transform:none}}@keyframes v6Orb{from{transform:translate(0,0)}to{transform:translate(-40px,35px)}}
/* cards */
.v2-card,.main-card,.aicfo-card,.feature-card{background:linear-gradient(180deg,rgba(35,48,70,.97),rgba(25,35,53,.97))!important;border:1px solid var(--line)!important;border-radius:var(--radius)!important;box-shadow:0 14px 38px rgba(0,0,0,.16)!important;color:var(--text)!important;transition:transform .22s ease,border-color .22s ease,box-shadow .22s ease!important}
.v2-card:hover,.main-card:hover,.feature-card:hover{transform:translateY(-2px);border-color:rgba(91,140,255,.26)!important;box-shadow:0 20px 48px rgba(0,0,0,.22)!important}
.v2-kpi-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.85rem;margin:.75rem 0 1rem}
.v2-kpi{position:relative;overflow:hidden;padding:1rem;border:1px solid var(--line);border-radius:17px;background:linear-gradient(145deg,rgba(37,51,74,.98),rgba(27,38,57,.97));min-height:126px;transition:.22s ease}
.v2-kpi:hover{transform:translateY(-3px);border-color:rgba(91,140,255,.32)}.v2-kpi:before{content:"";position:absolute;width:90px;height:90px;border-radius:50%;right:-35px;top:-40px;background:var(--accent,rgba(91,140,255,.18));filter:blur(3px);animation:v2Glow 5s ease-in-out infinite}
.v2-kpi-icon{font-size:.9rem;color:#adc3ff}.v2-kpi-label{font-size:.68rem;text-transform:uppercase;letter-spacing:.08em;color:#74829a;font-weight:850;margin-top:.65rem}.v2-kpi-value{font-size:1.55rem;font-weight:900;color:#fff;letter-spacing:-.045em;margin-top:.16rem}.v2-kpi-delta{font-size:.7rem;color:#72dea7;margin-top:.2rem;font-weight:750}
.v2-two-col{display:grid;grid-template-columns:minmax(0,1.55fr) minmax(300px,.75fr);gap:.9rem}.v2-panel{padding:1rem;border:1px solid var(--line);border-radius:18px;background:linear-gradient(180deg,rgba(34,47,69,.96),rgba(24,34,52,.97))}.v2-panel-title{display:flex;align-items:center;justify-content:space-between;color:#eef2ff;font-weight:820;font-size:.93rem;margin-bottom:.8rem}.v2-panel-kicker{font-size:.65rem;color:#71809a;text-transform:uppercase;letter-spacing:.1em;font-weight:850}
.v2-ai-brief{padding:1rem;border-radius:17px;background:linear-gradient(135deg,rgba(91,140,255,.14),rgba(138,109,255,.08));border:1px solid rgba(91,140,255,.22);position:relative;overflow:hidden}.v2-ai-brief:after{content:"✦";position:absolute;right:15px;top:5px;font-size:3.4rem;color:rgba(155,171,255,.08)}
.v2-action{display:flex;gap:.7rem;padding:.7rem 0;border-bottom:1px solid rgba(148,163,184,.10)}.v2-action:last-child{border-bottom:0}.v2-action-num{width:25px;height:25px;flex:0 0 25px;border-radius:8px;display:grid;place-items:center;background:rgba(91,140,255,.12);color:#a9c1ff;font-size:.68rem;font-weight:850}.v2-action-title{font-size:.79rem;color:#e5eaf5;font-weight:760}.v2-action-sub{font-size:.68rem;color:#71809a;margin-top:.12rem}

/* native Streamlit */
div[data-testid="stButton"]>button,div[data-testid="stDownloadButton"]>button{
 border-radius:11px!important;border:1px solid rgba(148,163,184,.20)!important;background:rgba(21,30,53,.88)!important;color:#dfe6f4!important;min-height:2.45rem!important;font-weight:720!important;transition:.18s ease!important;box-shadow:none!important}
div[data-testid="stButton"]>button:hover,div[data-testid="stDownloadButton"]>button:hover{border-color:rgba(91,140,255,.55)!important;background:linear-gradient(135deg,rgba(53,84,173,.95),rgba(94,66,171,.95))!important;color:white!important;transform:translateY(-1px);box-shadow:0 10px 25px rgba(67,97,209,.22)!important}
button[kind="primary"]{background:linear-gradient(135deg,#3979ee,#7055e8)!important;border-color:#779aff!important;color:white!important;box-shadow:0 10px 28px rgba(91,140,255,.24)!important}
[data-baseweb="input"]>div,[data-baseweb="select"]>div,[data-baseweb="textarea"]>div{background:#0d1427!important;border-color:rgba(148,163,184,.18)!important;border-radius:11px!important;color:#eaf0fb!important}
input,textarea{color:#eaf0fb!important}.stTabs [data-baseweb="tab-list"]{gap:.35rem;background:rgba(9,14,28,.72);padding:.3rem;border-radius:12px;border:1px solid var(--line)}.stTabs [data-baseweb="tab"]{border-radius:9px;padding:.45rem .78rem;color:#8795ad}.stTabs [aria-selected="true"]{background:rgba(91,140,255,.13)!important;color:#fff!important}.stTabs [data-baseweb="tab-highlight"]{display:none}
[data-testid="stMetric"]{background:linear-gradient(180deg,rgba(16,23,42,.95),rgba(10,15,30,.95));border:1px solid var(--line);padding:.85rem;border-radius:15px}[data-testid="stMetricValue"]{color:#fff!important;font-size:1.45rem!important}[data-testid="stMetricLabel"]{color:#7f8da5!important}
[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:14px;overflow:hidden;background:#0b1121}.stAlert{border-radius:13px!important;border:1px solid var(--line)!important;background:#0e162a!important}
hr{border-color:var(--line)!important}

/* AI copilot */
.st-key-open_ai_cfo_global{right:0!important;top:44vh!important;bottom:auto!important;width:54px!important;height:132px!important;transform:translateY(-50%)!important}
.st-key-open_ai_cfo_global button{width:54px!important;height:132px!important;min-height:132px!important;border-radius:18px 0 0 18px!important;background:linear-gradient(180deg,#4f8cff,#745ee8)!important;border:1px solid rgba(255,255,255,.30)!important;border-right:0!important;box-shadow:0 16px 42px rgba(55,91,190,.34)!important;animation:v2Pulse 3.2s ease-in-out infinite!important;font-size:20px!important}
.st-key-ai_cfo_overlay_panel{right:64px!important;top:50%!important;bottom:auto!important;transform:translateY(-50%)!important;width:min(420px,calc(100vw - 28px))!important;max-height:calc(100vh - 120px)!important;border-radius:20px!important;background:linear-gradient(180deg,rgba(13,19,36,.99),rgba(7,11,23,.99))!important;border:1px solid rgba(91,140,255,.28)!important;box-shadow:0 30px 90px rgba(0,0,0,.55)!important;animation:v2DrawerIn .28s cubic-bezier(.2,.8,.2,1) both!important}
@keyframes v2DrawerIn{from{opacity:0;transform:translateX(22px) scale(.98)}to{opacity:1;transform:none}}

/* login upgrade */
.st-key-login_shell{max-width:1420px;margin:auto}.st-key-login_shell>[data-testid="stHorizontalBlock"]{border:1px solid rgba(148,163,184,.16);border-radius:28px;overflow:hidden;background:#090f1e;box-shadow:0 35px 110px rgba(0,0,0,.5)}
.login-visual{min-height:760px!important}.login-headline{background:linear-gradient(120deg,#fff,#b7d1ff 45%,#beaaff);-webkit-background-clip:text;background-clip:text;color:transparent!important}.login-proof{gap:.55rem!important}.proof-item{transition:.22s ease}.proof-item:hover{transform:translateY(-3px);border-color:rgba(91,140,255,.35)}

@media(max-width:1050px){[data-testid="stSidebar"]{display:block!important;width:228px!important;min-width:228px!important;max-width:228px!important}.main .block-container{padding:1rem 1rem 5rem!important}.v2-kpi-grid{grid-template-columns:repeat(2,1fr)}.v2-two-col{grid-template-columns:1fr}}
@media(max-width:650px){.v2-kpi-grid{grid-template-columns:1fr}.v2-page-head{display:block}.v2-period-pill{display:inline-flex;margin-top:.6rem}.main .block-container{padding:.75rem .7rem 5rem!important}.v2-topbar{padding:.6rem}.v2-live{display:none}}


/* =====================================================================
   ENTERPRISE V3 ACCESSIBILITY + VISUAL POLISH OVERRIDES
   Brighter slate surfaces, strong text hierarchy and readable controls.
   ===================================================================== */
:root{
 --bg:#1d293d!important;
 --panel:#26364d!important;
 --panel2:#2d405b!important;
 --panel3:#354b69!important;
 --text:#f8fbff!important;
 --text-strong:#ffffff!important;
 --body:#e4ebf5!important;
 --muted:#bac6d8!important;
 --subtle:#94a5bd!important;
 --line:rgba(203,213,225,.22)!important;
 --blue:#6c9cff!important;
 --cyan:#54dbea!important;
 --violet:#9b82ff!important;
}
html,body,.stApp,[data-testid="stAppViewContainer"]{
 background:#1d293d!important;
 color:var(--body)!important;
}
.stApp{
 background:
 radial-gradient(circle at 76% -12%,rgba(82,132,235,.20),transparent 34%),
 radial-gradient(circle at 12% 108%,rgba(50,196,212,.11),transparent 30%),
 linear-gradient(145deg,#263753 0%,#1f2d43 46%,#1a2639 100%)!important;
}

/* Force a clear, accessible text hierarchy across native Streamlit widgets. */
.main h1,.main h2,.main h3,.main h4,.main h5,.main h6,
.main [data-testid="stMarkdownContainer"] h1,
.main [data-testid="stMarkdownContainer"] h2,
.main [data-testid="stMarkdownContainer"] h3,
.main [data-testid="stMarkdownContainer"] h4{
 color:#fff!important;opacity:1!important;text-shadow:0 1px 0 rgba(0,0,0,.10);
}
.main p,.main li,.main label,.main .stCaption,
.main [data-testid="stMarkdownContainer"] p,
.main [data-testid="stMarkdownContainer"] li,
[data-testid="stWidgetLabel"] p{
 color:#dbe5f2!important;opacity:1!important;
}
small,.small,.caption,[data-testid="stCaptionContainer"]{
 color:#b7c4d6!important;opacity:1!important;
}

/* Sidebar: stable, brighter and easier to scan. */
[data-testid="stSidebar"]{
 background:linear-gradient(180deg,#293b56 0%,#223249 58%,#1e2c40 100%)!important;
 border-right:1px solid rgba(213,223,238,.24)!important;
 box-shadow:14px 0 34px rgba(4,10,20,.15)!important;
}
[data-testid="stSidebar"]>div:first-child{overflow-y:auto!important;scrollbar-color:#7187a5 transparent!important}
.v2-brand-sub,.v2-workspace-period{color:#b6c3d6!important}
.v2-workspace-label,.v2-nav-heading{color:#a8bad2!important}
.v2-workspace{background:rgba(255,255,255,.075)!important;border-color:rgba(220,228,240,.28)!important}
.v2-workspace-name{color:#fff!important}
[data-testid="stSidebar"] div[data-testid="stButton"]>button{
 background:rgba(255,255,255,.045)!important;
 border-color:rgba(213,223,238,.16)!important;
 color:#dce5f2!important;
 opacity:1!important;
}
[data-testid="stSidebar"] div[data-testid="stButton"]>button p,
[data-testid="stSidebar"] div[data-testid="stButton"]>button span{
 color:inherit!important;opacity:1!important;
}
[data-testid="stSidebar"] div[data-testid="stButton"]>button:hover{
 background:rgba(108,156,255,.18)!important;color:#fff!important;
 border-color:rgba(139,177,255,.42)!important;transform:translateX(3px);
 box-shadow:0 8px 22px rgba(4,14,30,.15)!important;
}
[data-testid="stSidebar"] .st-key-v2_active_nav button{
 background:linear-gradient(90deg,rgba(74,128,244,.38),rgba(132,103,238,.20))!important;
 border-color:rgba(137,176,255,.60)!important;color:#fff!important;
 box-shadow:inset 4px 0 0 #70a3ff,0 10px 26px rgba(20,55,125,.22)!important;
}
.v2-sidebar-foot{background:rgba(255,255,255,.065)!important;border-color:rgba(220,228,240,.22)!important}
.v2-status-row{color:#c7d2e2!important}

/* Main shell and panels. */
.v2-topbar{
 background:linear-gradient(180deg,rgba(54,73,101,.96),rgba(43,59,84,.96))!important;
 border-color:rgba(219,228,240,.30)!important;
 box-shadow:0 13px 34px rgba(6,13,25,.18)!important;
}
.v2-breadcrumb{color:#b9c7d9!important}.v2-breadcrumb b{color:#fff!important}
.v2-page-title{color:#fff!important}.v2-page-sub{color:#c5d1e1!important}
.v2-period-pill{color:#dce5f2!important;background:rgba(255,255,255,.075)!important;border-color:rgba(220,228,240,.30)!important}
.v2-live{color:#aef0ce!important;background:rgba(33,185,116,.12)!important;border-color:rgba(89,228,163,.30)!important}

.v2-card,.main-card,.aicfo-card,.feature-card,.v2-panel,.v2-kpi{
 background:linear-gradient(180deg,rgba(52,71,99,.98),rgba(39,55,79,.98))!important;
 border-color:rgba(211,222,237,.25)!important;
 box-shadow:0 16px 38px rgba(7,15,28,.18)!important;
}
.v2-card:hover,.main-card:hover,.feature-card:hover,.v2-kpi:hover{
 border-color:rgba(128,170,255,.52)!important;
 box-shadow:0 22px 48px rgba(5,14,30,.25)!important;
}
.v2-kpi-label,.v2-panel-kicker{color:#b5c5da!important}
.v2-kpi-value,.v2-panel-title,.v2-action-title{color:#fff!important}
.v2-kpi-delta{color:#9aefc2!important}.v2-action-sub{color:#b7c4d5!important}
.v2-ai-brief{background:linear-gradient(135deg,rgba(82,132,235,.23),rgba(133,104,239,.14))!important;border-color:rgba(137,176,255,.40)!important}
.v2-ai-brief p{color:#dce7f5!important}

/* Tabs now read like controls, not disabled labels. */
[data-testid="stTabs"] [data-baseweb="tab-list"]{gap:.25rem;border-bottom:1px solid rgba(213,223,238,.24)!important}
[data-testid="stTabs"] button[role="tab"]{
 color:#c7d3e4!important;opacity:1!important;font-weight:700!important;
 padding:.6rem .78rem!important;border-radius:9px 9px 0 0!important;
}
[data-testid="stTabs"] button[role="tab"]:hover{color:#fff!important;background:rgba(255,255,255,.06)!important}
[data-testid="stTabs"] button[role="tab"][aria-selected="true"]{color:#fff!important;background:rgba(108,156,255,.14)!important}
[data-testid="stTabs"] [data-baseweb="tab-highlight"]{background:#68a0ff!important;height:3px!important}

/* Buttons and disabled states: disabled remains visibly disabled, never washed out. */
.main div[data-testid="stButton"]>button,
.main div[data-testid="stDownloadButton"]>button,
.main div[data-testid="stFormSubmitButton"]>button{
 color:#eef4ff!important;background:linear-gradient(180deg,#405573,#344963)!important;
 border:1px solid rgba(215,225,238,.28)!important;opacity:1!important;
 box-shadow:0 7px 18px rgba(5,12,24,.14)!important;
}
.main div[data-testid="stButton"]>button p,
.main div[data-testid="stDownloadButton"]>button p,
.main div[data-testid="stFormSubmitButton"]>button p{color:inherit!important;opacity:1!important}
.main div[data-testid="stButton"]>button:hover,
.main div[data-testid="stDownloadButton"]>button:hover,
.main div[data-testid="stFormSubmitButton"]>button:hover{
 color:#fff!important;background:linear-gradient(180deg,#5378b7,#405f93)!important;
 border-color:rgba(142,180,255,.60)!important;transform:translateY(-1px);
}
button:disabled,[aria-disabled="true"]{
 opacity:.68!important;filter:saturate(.70)!important;color:#bcc8d9!important;
}
button:disabled p,button:disabled span,[aria-disabled="true"] p,[aria-disabled="true"] span{color:#bcc8d9!important;opacity:1!important}

/* Inputs, selects and uploaders. */
[data-baseweb="input"]>div,[data-baseweb="select"]>div,[data-baseweb="textarea"]>div,
[data-testid="stFileUploaderDropzone"]{
 background:#31445f!important;border-color:rgba(214,225,240,.28)!important;color:#f4f7fb!important;
}
input,textarea,[data-baseweb="select"] span{color:#f3f7fc!important;opacity:1!important}
input::placeholder,textarea::placeholder{color:#aebdd1!important;opacity:1!important}
[data-testid="stFileUploaderDropzone"] p,[data-testid="stFileUploaderDropzone"] small{color:#d2dce9!important}

/* Metrics, alerts, expanders and data displays. */
[data-testid="stMetric"]{background:linear-gradient(180deg,#354a68,#2b3e59)!important;border:1px solid rgba(213,223,238,.25)!important;border-radius:16px!important;padding:.85rem!important}
[data-testid="stMetricLabel"] p{color:#bdcadb!important}
[data-testid="stMetricValue"]{color:#fff!important}
[data-testid="stMetricDelta"]{opacity:1!important}
[data-testid="stExpander"]{background:rgba(48,66,93,.82)!important;border-color:rgba(213,223,238,.24)!important}
[data-testid="stExpander"] summary p{color:#eef4fb!important}
[data-testid="stAlert"]{background:#324761!important;border-color:rgba(213,223,238,.27)!important;color:#eef4fb!important}
[data-testid="stAlert"] p{color:#eef4fb!important}
[data-testid="stDataFrame"],[data-testid="stTable"]{border:1px solid rgba(213,223,238,.24)!important;border-radius:13px!important;overflow:hidden!important}

/* AI edge launcher: lively but unobtrusive and always reachable. */
.st-key-open_ai_cfo_global{top:50%!important;right:0!important;bottom:auto!important;transform:translateY(-50%)!important;z-index:999998!important}
.st-key-open_ai_cfo_global button{
 background:linear-gradient(180deg,#63a0ff 0%,#7469ef 58%,#8e69ee 100%)!important;
 color:#fff!important;opacity:1!important;
 box-shadow:0 16px 45px rgba(66,101,205,.40)!important;
 transition:width .22s ease,transform .22s ease,filter .22s ease!important;
}
.st-key-open_ai_cfo_global button:hover{width:62px!important;transform:translateX(-2px)!important;filter:brightness(1.10)!important}
.st-key-open_ai_cfo_global button p{color:#fff!important;opacity:1!important}
.st-key-ai_cfo_overlay_panel{
 background:linear-gradient(180deg,rgba(45,61,86,.995),rgba(28,40,59,.995))!important;
 border-color:rgba(137,176,255,.48)!important;
 box-shadow:0 32px 90px rgba(2,8,20,.48)!important;
}

/* Login is lighter and consistent with the workspace. */
.st-key-login_shell>[data-testid="stHorizontalBlock"]{
 background:linear-gradient(145deg,#2d405c,#22334c)!important;
 border-color:rgba(216,226,239,.28)!important;
}

/* Streamlit sometimes places faded text inside nested blocks. Cancel it. */
.main [style*="opacity"], [data-testid="stSidebar"] [style*="opacity"]{opacity:1!important}

@media(max-width:650px){
 [data-testid="stSidebar"]{box-shadow:none!important}
 .st-key-open_ai_cfo_global{top:auto!important;right:14px!important;bottom:16px!important;transform:none!important;width:auto!important;height:auto!important}
 .st-key-open_ai_cfo_global button{width:auto!important;height:48px!important;min-height:48px!important;border-radius:14px!important;padding:0 .9rem!important}
}

</style>
        """,
        unsafe_allow_html=True,
    )

    compact = bool(st.session_state.get("v2_sidebar_compact", False))
    sidebar_width = "88px" if compact else "260px"
    compact_css = ""
    if compact:
        compact_css = """
        [data-testid="stSidebar"] .v2-brand > div:last-child,
        [data-testid="stSidebar"] .v2-workspace,
        [data-testid="stSidebar"] .v2-nav-heading,
        [data-testid="stSidebar"] .v2-sidebar-foot{display:none!important;}
        [data-testid="stSidebar"] .v2-brand{justify-content:center;border-bottom:0;margin-bottom:.35rem;padding-bottom:.25rem;}
        [data-testid="stSidebar"] div[data-testid="stButton"]>button{justify-content:center!important;padding:.5rem!important;font-size:1.05rem!important;}
        [data-testid="stSidebar"] div[data-testid="stButton"]>button p{font-size:1.05rem!important;}
        """
    st.markdown(
        f"""
        <style>
        section[data-testid="stSidebar"],[data-testid="stSidebar"]{{
          width:{sidebar_width}!important;min-width:{sidebar_width}!important;max-width:{sidebar_width}!important;
          transition:width .28s cubic-bezier(.2,.8,.2,1),min-width .28s cubic-bezier(.2,.8,.2,1),max-width .28s cubic-bezier(.2,.8,.2,1)!important;
        }}
        [data-testid="stSidebar"]>div:first-child{{padding:{'.72rem .5rem' if compact else '1.1rem .85rem 1rem'}!important;}}
        {compact_css}
        /* Native chart toolbar remains available, including Show data / Show chart. */
        /* Extra motion with restraint. */
        .v2-kpi:nth-child(1){{animation:v4CardIn .45s .04s both}}
        .v2-kpi:nth-child(2){{animation:v4CardIn .45s .10s both}}
        .v2-kpi:nth-child(3){{animation:v4CardIn .45s .16s both}}
        .v2-kpi:nth-child(4){{animation:v4CardIn .45s .22s both}}
        .v2-panel{{animation:v4CardIn .52s .12s both}}
        @keyframes v4CardIn{{from{{opacity:0;transform:translateY(16px) scale(.985)}}to{{opacity:1;transform:none}}}}
        [data-testid="stSidebar"] div[data-testid="stButton"]>button{{color:#eef4ff!important;opacity:1!important;font-weight:760!important;letter-spacing:.005em!important;}}
        [data-testid="stSidebar"] div[data-testid="stButton"]>button p{{color:inherit!important;opacity:1!important;}}
        [data-testid="stSidebar"] div[data-testid="stButton"]>button:hover{{transform:translateX(4px) scale(1.012)!important;box-shadow:0 10px 28px rgba(44,92,190,.20)!important;}}
        .v2-brand-mark{{animation:v5LogoFloat 4s ease-in-out infinite!important;}}
        .v2-status-dot{{animation:v5StatusPulse 2s ease-in-out infinite!important;}}
        @keyframes v5LogoFloat{{0%,100%{{transform:translateY(0) rotate(0)}}50%{{transform:translateY(-4px) rotate(2deg)}}}}
        @keyframes v5StatusPulse{{0%,100%{{box-shadow:0 0 0 0 rgba(57,217,138,.25)}}50%{{box-shadow:0 0 0 8px rgba(57,217,138,0)}}}}
        </style>
        """, unsafe_allow_html=True
    )


def render_enterprise_sidebar(selected_slug: str, profile: dict, period_label: str, score: int) -> str:
    company = escape(str(profile.get("Company Name") or "Sample Company"))
    period = escape(str(period_label or "Current period"))
    with st.sidebar:
        compact = bool(st.session_state.get("v2_sidebar_compact", False))
        toggle_label = "→" if compact else "☰  Collapse"
        if st.button(toggle_label, key="v2_sidebar_toggle", use_container_width=True, help="Expand navigation" if compact else "Collapse navigation"):
            st.session_state["v2_sidebar_compact"] = not compact
            st.rerun()
        st.markdown(
            f"""
<div class="v2-brand"><div class="v2-brand-mark">A</div><div><div class="v2-brand-name">AI CFO Copilot</div><div class="v2-brand-sub">Finance intelligence workspace</div></div></div>
<div class="v2-workspace"><div class="v2-workspace-label">Active workspace</div><div class="v2-workspace-name">{company}</div><div class="v2-workspace-period">{period}</div></div>
<div class="v2-nav-heading">Finance command centre</div>
            """,
            unsafe_allow_html=True,
        )
        for idx, (slug, label, icon, _) in enumerate(PAGE_CONFIG):
            if slug == "upload":
                st.markdown('<div class="v2-nav-heading">Setup & data</div>', unsafe_allow_html=True)
            elif slug == "analytics":
                st.markdown('<div class="v2-nav-heading">Performance & reporting</div>', unsafe_allow_html=True)
            elif slug == "working_capital":
                st.markdown('<div class="v2-nav-heading">Decision intelligence</div>', unsafe_allow_html=True)
            prefix = "●" if slug == selected_slug else icon
            key = "v2_active_nav" if slug == selected_slug else f"v2_nav_{slug}"
            nav_label = prefix if compact else f"{prefix}  {label}"
            if st.button(nav_label, key=key, use_container_width=True, help=f"{label} — {dict((x[0], x[3]) for x in PAGE_CONFIG)[slug]}"):
                st.query_params["page"] = slug
                st.rerun()
        st.markdown(
            f"""
<div class="v2-sidebar-foot"><div class="v2-status-row"><span>Workspace readiness</span><b style="color:#dce5f5">{int(score)}/100</b></div><div style="height:6px;background:#172036;border-radius:999px;margin:.55rem 0;overflow:hidden"><div style="height:100%;width:{max(0,min(100,int(score)))}%;background:linear-gradient(90deg,#35d4e8,#5b8cff,#8a6dff);border-radius:999px"></div></div><div class="v2-status-row"><span>Services online</span><span class="v2-status-dot"></span></div></div>
            """,
            unsafe_allow_html=True,
        )
    return selected_slug


def render_enterprise_topbar(selected_slug: str, profile: dict, period_label: str) -> None:
    page = next((item for item in PAGE_CONFIG if item[0] == selected_slug), PAGE_CONFIG[0])
    company = escape(str(profile.get("Company Name") or "Finance Workspace"))
    initials = "".join(part[0] for part in company.split()[:2]).upper() or "FC"
    now_label = datetime.now().strftime("%d %b %Y")
    st.markdown(
        f"""
<div class="v2-topbar"><div class="v2-breadcrumb">{company} &nbsp;/&nbsp; <b>{page[1]}</b></div><div class="v2-top-actions"><div class="v2-live"><span class="v2-status-dot"></span> Live workspace</div><div class="v2-period-pill">{escape(str(period_label))}</div><div class="v2-avatar" title="{now_label}">{initials}</div></div></div>
<div class="v2-page-head"><div><div class="v2-page-title">{page[2]}&nbsp; {page[1]}</div><div class="v2-page-sub">{page[3]}</div></div><div class="v2-period-pill">Reporting period · {escape(str(period_label))}</div></div>
        """,
        unsafe_allow_html=True,
    )


def render_dashboard_v2_summary(profile: dict, score: int, flags: Iterable[bool], state: dict) -> None:
    complete = sum(bool(x) for x in flags)
    progress = int(complete / 5 * 100)
    pnl = state.get("consolidated_pnl")
    mapped = state.get("mapped")
    ai_ready = bool(state.get("ai_commentary"))
    rows = len(mapped) if getattr(mapped, "empty", True) is False else 0
    company = escape(str(profile.get("Company Name") or "Your company"))
    currency = escape(str(profile.get("Currency") or "USD"))
    first_name = escape(str(profile.get("Contact Name") or profile.get("User Name") or "Finance team").split()[0])
    greeting = "Good morning" if datetime.now().hour < 12 else ("Good afternoon" if datetime.now().hour < 18 else "Good evening")
    date_label = datetime.now().strftime("%A, %d %B %Y")
    st.markdown(
        f"""
<div class="v6-welcome"><div><div class="v6-greeting">{greeting}, {first_name}</div><div class="v6-welcome-copy">Here is what needs your attention across {company} today.</div></div><div class="v6-date">{date_label}</div></div>
<div class="v2-kpi-grid">
 <div class="v2-kpi" style="--accent:rgba(53,212,232,.16)"><div class="v2-kpi-icon">◈</div><div class="v2-kpi-label">Close readiness</div><div class="v2-kpi-value">{progress}%</div><div class="v2-kpi-delta">{complete} of 5 workflow stages complete</div></div>
 <div class="v2-kpi" style="--accent:rgba(91,140,255,.18)"><div class="v2-kpi-icon">▦</div><div class="v2-kpi-label">Data processed</div><div class="v2-kpi-value">{rows:,}</div><div class="v2-kpi-delta">Validated ledger records</div></div>
 <div class="v2-kpi" style="--accent:rgba(138,109,255,.18)"><div class="v2-kpi-icon">✦</div><div class="v2-kpi-label">AI CFO status</div><div class="v2-kpi-value">{'Ready' if ai_ready else 'Standby'}</div><div class="v2-kpi-delta">Market-aware finance commentary</div></div>
 <div class="v2-kpi" style="--accent:rgba(57,217,138,.16)"><div class="v2-kpi-icon">✓</div><div class="v2-kpi-label">Validation score</div><div class="v2-kpi-value">{int(score)}/100</div><div class="v2-kpi-delta">Reporting confidence indicator</div></div>
</div>
<div class="v2-two-col">
 <div class="v2-panel"><div class="v2-panel-title"><span>Executive finance workspace</span><span class="v2-panel-kicker">{company} · {currency}</span></div><div class="v2-ai-brief"><div class="v2-panel-kicker">AI CFO morning brief</div><h3 style="margin:.35rem 0;color:#fff">Your finance operating rhythm is {progress}% complete.</h3><p style="font-size:.8rem;line-height:1.55;margin:0;color:#aebbd0">Use the command centre to finish validation, review performance drivers and generate board-ready commentary. External market context is automatically connected when research credentials are configured.</p></div></div>
 <div class="v2-panel"><div class="v2-panel-title"><span>Priority actions</span><span class="v2-panel-kicker">Today</span></div>
  <div class="v2-action"><div class="v2-action-num">1</div><div><div class="v2-action-title">Complete data validation</div><div class="v2-action-sub">Resolve mapping gaps before publishing reports</div></div></div>
  <div class="v2-action"><div class="v2-action-num">2</div><div><div class="v2-action-title">Review variance drivers</div><div class="v2-action-sub">Focus on margin, overhead and cash conversion</div></div></div>
  <div class="v2-action"><div class="v2-action-num">3</div><div><div class="v2-action-title">Generate management brief</div><div class="v2-action-sub">Turn finance results into clear actions</div></div></div>
 </div>
</div>
        """,
        unsafe_allow_html=True,
    )
