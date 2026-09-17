from __future__ import annotations

import html
from typing import Iterable

import pandas as pd
import streamlit as st

WFP_BLUE = "#007DBC"
WFP_DARK = "#075B93"
NAV_DARK = "#07578B"
TEXT = "#143E67"
MUTED = "#647C93"
BORDER = "#DDE7F0"
BG = "#F7F9FC"

KPI_PALETTES = {
    "blue": ("#EAF5FF", "#0B6FB7"),
    "green": ("#ECFAF2", "#148A4D"),
    "purple": ("#F1EDFF", "#6D3DE8"),
    "orange": ("#FFF5E8", "#E66B00"),
    "teal": ("#ECFBF8", "#008B83"),
    "red": ("#FFF0F2", "#D92D43"),
    "yellow": ("#FFF9E6", "#B77900"),
    "gray": ("#F3F6F8", "#586B7C"),
}


def inject_styles() -> None:
    st.markdown(
        f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;500;600;700;800&display=swap');
:root {{ --wfp:{WFP_BLUE}; --wfp-dark:{WFP_DARK}; --nav:{NAV_DARK}; --text:{TEXT}; --muted:{MUTED}; --border:{BORDER}; --bg:{BG}; }}
html, body, [class*="css"] {{ font-family: "Open Sans", "Segoe UI", Arial, sans-serif; color:var(--text); }}
.stApp {{ background:var(--bg); }}
[data-testid="stHeader"] {{ display:none; }}
#MainMenu, footer {{ visibility:hidden; }}
.block-container {{ padding:0 1.0rem 1.5rem 1.0rem; max-width:100%; }}

/* Sidebar */
[data-testid="stSidebar"] {{
  background:linear-gradient(180deg,#07588C 0%,#006EB2 58%,#0877B9 100%);
  border-right:none; min-width:220px; max-width:220px;
}}
[data-testid="stSidebar"] > div:first-child {{ padding-top:.7rem; }}
[data-testid="stSidebar"] label, [data-testid="stSidebar"] p,
[data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] span {{ color:#F4FAFF !important; }}
[data-testid="stSidebar"] hr {{ border-color:rgba(255,255,255,.18); }}
[data-testid="stSidebar"] [data-baseweb="select"] > div,
[data-testid="stSidebar"] input {{ background:rgba(255,255,255,.98) !important; color:#173A5B !important; }}
[data-testid="stSidebar"] [role="radiogroup"] {{ gap:1px; }}
[data-testid="stSidebar"] [data-baseweb="radio"] {{
  padding:7px 8px; margin:1px -6px; border-radius:6px; transition:.15s;
}}
[data-testid="stSidebar"] [data-baseweb="radio"]:hover {{ background:rgba(255,255,255,.10); }}
[data-testid="stSidebar"] [data-baseweb="radio"] > div:first-child {{ display:none; }}
.sidebar-logo {{ padding:6px 6px 12px 6px; border-bottom:1px solid rgba(255,255,255,.18); margin-bottom:8px; }}
.sidebar-wfp {{ display:flex;align-items:center;gap:9px;color:white; }}
.sidebar-emblem {{ width:42px;height:42px;border:2px solid white;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:13px; }}
.sidebar-org {{ font-weight:700;font-size:13px;line-height:1.15; }}
.sidebar-slogan {{ font-size:9px;opacity:.88;margin-top:3px;letter-spacing:.25px; }}
.sidebar-footer {{ margin-top:24px; padding:14px 6px; color:#E5F4FF; font-size:11px; line-height:1.4; }}

/* Top header */
.app-header {{
  background:white; border-bottom:1px solid var(--border); min-height:78px;
  margin:0 -1rem 8px -1rem; padding:10px 20px 8px 20px;
  display:flex; align-items:center; justify-content:space-between; gap:24px;
}}
.app-name {{ color:#0A4F84; font-size:31px; font-weight:800; letter-spacing:-.8px; line-height:1; }}
.app-name .ver {{ color:#69AEE8; font-size:24px; font-weight:700; }}
.app-sub {{ color:#0B5F9B; font-size:14px; font-weight:700; margin-top:4px; }}
.app-tag {{ color:#6C8297; font-size:11px; margin-top:2px; }}
.header-right {{ display:flex;align-items:center;gap:22px; }}
.region-block {{ text-align:right; }}
.region-title {{ font-weight:700;color:#0A4F84;font-size:13px; }}
.region-sub {{ color:#7A8FA3;font-size:10px;margin-top:3px; }}
.user-chip {{ display:flex;align-items:center;gap:8px;border-left:1px solid var(--border);padding-left:18px; }}
.user-avatar {{ width:37px;height:37px;border-radius:50%;background:#3D98DE;color:white;font-weight:700;display:flex;align-items:center;justify-content:center; }}
.user-name {{ font-size:12px;color:#0A4F84;font-weight:700; }}
.user-role {{ font-size:10px;color:#7A8FA3; }}

/* Titles */
.dashboard-title {{ color:#073D70;font-size:28px;font-weight:800;margin:4px 0 0 0;letter-spacing:-.4px; }}
.dashboard-subtitle {{ color:#657C92;font-size:12px;margin:0 0 10px 0; }}
.section-title {{ color:#0C497D;font-size:15px;font-weight:800;margin:8px 0 7px 0; }}

/* Native widget polish */
[data-testid="stSelectbox"] > div > div, [data-testid="stMultiSelect"] > div > div,
[data-testid="stTextInput"] > div > div, [data-testid="stDateInput"] > div > div {{ border-radius:6px; }}
[data-testid="stButton"] button, [data-testid="stDownloadButton"] button {{ border-radius:6px; font-weight:700; }}

/* Generic cards */
.panel {{ background:white;border:1px solid var(--border);border-radius:7px;padding:12px 13px;box-shadow:0 1px 2px rgba(17,62,100,.035); }}
.panel-title {{ color:#0B4D83;font-weight:800;font-size:14px;margin-bottom:8px; }}
.card { background:white;border:1px solid var(--border);border-radius:7px;padding:13px;box-shadow:0 1px 2px rgba(17,62,100,.035); }
.step-num { display:inline-flex;width:24px;height:24px;border-radius:50%;background:#E6F3FF;color:#1377BF;align-items:center;justify-content:center;font-weight:800;margin-right:7px; }
.muted { color:#6D8295;font-size:10px;line-height:1.45; }
.kpi-card {{ border:1px solid var(--border);border-radius:7px;padding:11px 12px;min-height:78px;display:flex;align-items:center;gap:11px; }}
.kpi-icon {{ width:42px;height:42px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:24px;font-weight:800;flex:0 0 auto; }}
.kpi-value {{ font-size:24px;font-weight:800;line-height:1;color:#104C7B; }}
.kpi-label {{ font-size:11px;color:#5E768D;margin-top:4px;white-space:nowrap; }}
.kpi-note {{ font-size:9px;color:#7890A5;margin-top:3px; }}

/* Filter frame */
.filter-strip {{ background:white;border-bottom:1px solid var(--border);padding:4px 0 7px 0;margin-bottom:4px; }}

/* tabs */
.stTabs [data-baseweb="tab-list"] {{ gap:0; border-bottom:1px solid var(--border); }}
.stTabs [data-baseweb="tab"] {{ background:white;border:1px solid var(--border);border-bottom:none;padding:7px 18px;height:38px; }}
.stTabs [aria-selected="true"] {{ background:var(--wfp) !important;color:white !important; }}

/* HTML tables */
.jm-table-wrap {{ background:white;border:1px solid var(--border);border-radius:7px;overflow:hidden; }}
.jm-table-title {{ padding:9px 12px;color:#0A4B80;font-weight:800;font-size:14px;border-bottom:1px solid var(--border); }}
table.jm-table {{ width:100%;border-collapse:collapse;font-size:10px; }}
.jm-table th {{ background:#F5F8FC;color:#365A79;text-align:left;padding:7px 8px;border-bottom:1px solid var(--border);font-weight:700; }}
.jm-table td {{ padding:6px 8px;border-bottom:1px solid #E9EFF4;color:#31526D;vertical-align:middle; }}
.jm-table tr:last-child td {{ border-bottom:none; }}
.badge {{ display:inline-block;padding:2px 8px;border-radius:4px;font-size:9px;font-weight:700;white-space:nowrap; }}
.badge-critical {{ background:#F04452;color:white; }}
.badge-high {{ background:#FF7A1A;color:white; }}
.badge-medium-high {{ background:#FF9E2A;color:white; }}
.badge-medium {{ background:#F5C928;color:#624B00; }}
.badge-low {{ background:#20A75A;color:white; }}
.badge-proposed {{ background:#FCE9EC;color:#D02A40; }}
.badge-in-progress {{ background:#E8F3FE;color:#1477C9; }}
.badge-closed {{ background:#E8F7ED;color:#148244; }}
.badge-pending {{ background:#EEE7FF;color:#7346D9; }}
.badge-agreed {{ background:#DFF4FF;color:#0874AE; }}
.badge-assigned {{ background:#EEE7FF;color:#7346D9; }}
.badge-neutral {{ background:#EDF1F4;color:#566C7E; }}

/* right rail */
.rail-card {{ background:white;border:1px solid var(--border);border-radius:7px;margin-bottom:10px;overflow:hidden; }}
.rail-head {{ color:#0A4B80;font-size:14px;font-weight:800;padding:10px 12px;border-bottom:1px solid var(--border); }}
.rail-item {{ padding:9px 12px;border-bottom:1px solid #ECF1F5;display:flex;gap:8px; }}
.rail-item:last-child {{ border-bottom:none; }}
.rail-dot {{ width:9px;height:9px;border-radius:50%;background:#E43C4F;margin-top:4px;flex:0 0 auto; }}
.rail-main {{ color:#285271;font-size:10px;line-height:1.3;flex:1; }}
.rail-meta {{ color:#8295A6;font-size:9px;margin-top:2px; }}
.rail-count {{ font-weight:800;color:#0A4B80;font-size:11px; }}

/* alerts */
.status-ok {{background:#ECFDF3;border:1px solid #ABEFC6;padding:9px 12px;border-radius:7px;color:#176B45;}}
.status-warn {{background:#FFFAEB;border:1px solid #FEDF89;padding:9px 12px;border-radius:7px;color:#8A5A00;}}
.status-info {{background:#EFF8FF;border:1px solid #B2DDFF;padding:9px 12px;border-radius:7px;color:#175C8D;}}
</style>
""",
        unsafe_allow_html=True,
    )


def render_sidebar_brand() -> None:
    st.markdown(
        """
<div class="sidebar-logo">
  <div class="sidebar-wfp">
    <div class="sidebar-emblem">WFP</div>
    <div>
      <div class="sidebar-org">World Food<br>Programme</div>
      <div class="sidebar-slogan">SAVING LIVES · CHANGING LIVES</div>
    </div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


def render_header(version: str = "v1.8", user_name: str = "Mohamed Mussie", user_role: str = "Jijiga AO") -> None:
    initials = "".join(x[:1] for x in user_name.split()[:2]).upper() or "JM"
    st.markdown(
        f"""
<div class="app-header">
  <div>
    <div class="app-name">J-MAAP <span class="ver">{html.escape(version)}</span></div>
    <div class="app-sub">Jijiga Area Office – Monitoring Automation and Analysis Platform</div>
    <div class="app-tag">From Data to Decisions for Greater Impact</div>
  </div>
  <div class="header-right">
    <div class="region-block"><div class="region-title">Somali Region, Ethiopia</div><div class="region-sub">People &nbsp;|&nbsp; Evidence &nbsp;|&nbsp; Accountability</div></div>
    <div class="user-chip"><div class="user-avatar">{initials}</div><div><div class="user-name">{html.escape(user_name)}</div><div class="user-role">{html.escape(user_role)}</div></div></div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


def page_title(title: str, subtitle: str) -> None:
    st.markdown(f'<div class="dashboard-title">{html.escape(title)}</div><div class="dashboard-subtitle">{html.escape(subtitle)}</div>', unsafe_allow_html=True)


def kpi(label: str, value, icon: str = "●", palette: str = "blue", note: str | None = None) -> None:
    bg, fg = KPI_PALETTES.get(palette, KPI_PALETTES["blue"])
    note_html = f'<div class="kpi-note">{html.escape(str(note))}</div>' if note else ""
    st.markdown(
        f"""
<div class="kpi-card" style="background:{bg};border-color:{fg}22;">
  <div class="kpi-icon" style="background:{fg}16;color:{fg};">{html.escape(str(icon))}</div>
  <div><div class="kpi-value" style="color:{fg};">{html.escape(str(value))}</div><div class="kpi-label">{html.escape(label)}</div>{note_html}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def severity_badge(value: str) -> str:
    s = str(value or "").strip()
    low = s.casefold().replace(" ", "-")
    cls = {
        "critical": "critical",
        "high": "high",
        "medium-high": "medium-high",
        "medium": "medium",
        "low": "low",
    }.get(low, "neutral")
    return f'<span class="badge badge-{cls}">{html.escape(s or "—")}</span>'


def status_badge(value: str) -> str:
    s = str(value or "").strip()
    low = s.casefold()
    if "closed" in low or "complete" in low:
        cls = "closed"
    elif "progress" in low:
        cls = "in-progress"
    elif "proposed" in low or "not started" in low:
        cls = "proposed"
    elif "pending" in low:
        cls = "pending"
    elif "agreed" in low:
        cls = "agreed"
    elif "assign" in low:
        cls = "assigned"
    else:
        cls = "neutral"
    return f'<span class="badge badge-{cls}">{html.escape(s or "—")}</span>'


def findings_action_table(df: pd.DataFrame, max_rows: int = 12) -> str:
    if df is None or df.empty:
        return '<div class="jm-table-wrap"><div class="jm-table-title">Action Tracker</div><div style="padding:16px;color:#7890A5;">No action records for the current filters.</div></div>'
    rows = []
    view = df.head(max_rows)
    for _, r in view.iterrows():
        action = r.get("Aggregated_Action_ID", r.get("Action_ID", ""))
        finding = r.get("Aggregated_Action_Title", r.get("Finding", ""))
        activity = r.get("Activities", r.get("Activity", ""))
        woreda = r.get("Woredas", r.get("Woreda", ""))
        if isinstance(woreda, str) and "," in woreda:
            woreda = ", ".join([x.strip() for x in woreda.split(",")[:2]]) + ("…" if len(woreda.split(",")) > 2 else "")
        severity = r.get("Priority", r.get("Severity", ""))
        owner = r.get("Responsible_Person") or r.get("Suggested_Responsible_Unit", "") or "—"
        due = r.get("Due_Date", "")
        due = "—" if pd.isna(due) or str(due).strip() in {"", "NaT", "nan"} else str(due)[:10]
        status = r.get("Status", "Proposed")
        validation = r.get("Validation_Status", "Pending field validation")
        rows.append(
            "<tr>"
            f"<td>{html.escape(str(action))}</td>"
            f"<td style='max-width:280px'>{html.escape(str(finding))}</td>"
            f"<td>{html.escape(str(activity))}</td>"
            f"<td>{html.escape(str(woreda))}</td>"
            f"<td>{severity_badge(str(severity))}</td>"
            f"<td>{html.escape(str(owner))}</td>"
            f"<td>{html.escape(due)}</td>"
            f"<td>{status_badge(str(status))}</td>"
            f"<td>{status_badge(str(validation))}</td>"
            "</tr>"
        )
    return (
        '<div class="jm-table-wrap"><div class="jm-table-title">Action Tracker</div>'
        '<table class="jm-table"><thead><tr><th>#</th><th>Issue / Finding</th><th>Activity</th><th>Woreda</th><th>Severity</th><th>Owner</th><th>Due Date</th><th>Status</th><th>Validation Stage</th></tr></thead><tbody>'
        + "".join(rows) + "</tbody></table>"
        + f'<div style="padding:7px 10px;color:#7890A5;font-size:9px;">Showing {min(len(df),max_rows)} of {len(df)} action records</div></div>'
    )


def priority_actions_html(df: pd.DataFrame, max_items: int = 5) -> str:
    if df is None or df.empty:
        items = '<div class="rail-item"><div class="rail-main">No priority actions for the current filters.</div></div>'
    else:
        rank = {"Critical": 5, "High": 4, "Medium-High": 3, "Medium": 2, "Low": 1}
        tmp = df.copy()
        pcol = "Priority" if "Priority" in tmp.columns else "Severity"
        tmp["_rank"] = tmp[pcol].map(rank).fillna(0)
        tmp = tmp.sort_values(["_rank"], ascending=False).head(max_items)
        parts = []
        for _, r in tmp.iterrows():
            title = r.get("Aggregated_Action_Title", r.get("Finding", ""))
            meta = r.get("Activities", r.get("Activity", ""))
            woreda = r.get("Woredas", r.get("Woreda", ""))
            if woreda:
                meta = f"{str(woreda).split(',')[0].strip()} | {str(meta).split(',')[0].strip()}"
            parts.append(f'<div class="rail-item"><div class="rail-dot"></div><div class="rail-main">{html.escape(str(title))}<div class="rail-meta">{html.escape(str(meta))}</div></div></div>')
        items = "".join(parts)
    return f'<div class="rail-card"><div class="rail-head">Top Priority Actions</div>{items}</div>'


def recurring_issues_html(ddf: pd.DataFrame, max_items: int = 5) -> str:
    if ddf is None or ddf.empty or "Finding" not in ddf.columns:
        body = '<div class="rail-item"><div class="rail-main">No recurring issues identified.</div></div>'
    else:
        vc = ddf["Finding"].fillna("Unspecified").value_counts().head(max_items)
        body = "".join(
            f'<div class="rail-item"><div class="rail-main">{html.escape(str(k))}</div><div class="rail-count">{int(v)}</div></div>' for k, v in vc.items()
        )
    return f'<div class="rail-card"><div class="rail-head">Recurring Issues (This Period)</div>{body}</div>'
