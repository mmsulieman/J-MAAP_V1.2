from __future__ import annotations

import json
import html
import tempfile
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

from io_utils import save_uploads, list_xlsx_sheets, preview_file, preview_moda_core
from dq_matching import (
    profile_file,
    pairwise_matching,
    moda_file_overlap,
    unmatched_sites,
    scoped_source_site_matching,
)
from pipeline import run_analysis, available_periods
from analytics import rows_to_df, overview_metrics, cfm_journey, cfm_woreda_table, protection_summary
from outputs import detailed_tracker_xlsx, aggregated_tracker_xlsx, management_report_docx, dq_report_docx, bundle_zip
from ui_components import (
    inject_styles, render_sidebar_brand, render_header, page_title, kpi,
    findings_action_table, priority_actions_html, recurring_issues_html,
)
from mapping import build_site_points, apply_map_filters, apply_photo_filters, load_admin_assets, leaflet_html, leaflet_dashboard_html, ACTIVITY_LABELS

# -----------------------------------------------------------------------------
# PAGE / BRANDING
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="J-MAAP | Monitoring Automation and Analysis Platform",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_styles()


def branded_header(reporting_month: str) -> None:
    # The reporting period is shown in the filter strip; keep the top header clean.
    render_header(version="v1.8.1", user_name="Mohamed Mussie", user_role="Jijiga AO")


def page_heading(title: str, subtitle: str) -> None:
    page_title(title, subtitle)


def fmt_pct(x):
    return f"{x:.1%}" if pd.notna(x) else ""


def clean_plot(fig, height=None, legend=True):
    fig.update_layout(
        paper_bgcolor="white", plot_bgcolor="white", margin=dict(l=12,r=12,t=46,b=12),
        font=dict(family="Open Sans, Segoe UI, Arial", size=10, color="#345873"),
        title_font=dict(size=14, color="#0A4B80"),
        legend=dict(font=dict(size=9), orientation="v") if legend else dict(visible=False),
    )
    if height:
        fig.update_layout(height=height)
    fig.update_xaxes(gridcolor="#E7EDF3", zeroline=False)
    fig.update_yaxes(gridcolor="#E7EDF3", zeroline=False)
    return fig


def dashboard_filter_row(rows, prefix: str, show_monitoring_type: bool = True):
    rows = rows or []
    activities = sorted({ACTIVITY_LABELS.get(str(r.get("activity") or "").strip(), str(r.get("activity") or "").strip()) for r in rows if str(r.get("activity") or "").strip()})
    providers = sorted({("TPM" if ("third party" in str(r.get("provider") or "").casefold() or "wfma" in str(r.get("provider") or "").casefold()) else "WFP" if "wfp" in str(r.get("provider") or "").casefold() else str(r.get("provider") or "").strip()) for r in rows if str(r.get("provider") or "").strip()})
    zones = sorted({str(r.get("zone") or "").strip() for r in rows if str(r.get("zone") or "").strip()})
    c = st.columns([1.0,1.0,1.0,1.0,1.0,1.0,1.15,0.7])
    with c[0]: st.text_input("Reporting Month", value=reporting_month, disabled=True, key=f"{prefix}_month")
    with c[1]: st.text_input("Sub-office", value=sub_office, disabled=True, key=f"{prefix}_so")
    with c[2]: activity = st.selectbox("Activity", ["All"]+activities, key=f"{prefix}_activity")
    with c[3]: provider = st.selectbox("Provider", ["All"]+providers, key=f"{prefix}_provider")
    with c[4]: zone = st.selectbox("Zone", ["All"]+zones, key=f"{prefix}_zone")
    zone_rows = rows if zone=="All" else [r for r in rows if str(r.get("zone") or "").strip()==zone]
    woredas = sorted({str(r.get("woreda") or "").strip() for r in zone_rows if str(r.get("woreda") or "").strip()})
    with c[5]: woreda = st.selectbox("Woreda", ["All"]+woredas, key=f"{prefix}_woreda")
    monitoring_types = sorted({str(r.get("monitoring_type") or "").strip() for r in rows if str(r.get("monitoring_type") or "").strip()})
    with c[6]:
        mtype = st.selectbox("Monitoring Type", ["All"]+monitoring_types, key=f"{prefix}_mtype") if show_monitoring_type else "All"
    with c[7]:
        st.markdown("<div style='height:27px'></div>", unsafe_allow_html=True)
        if st.button("↻ Reset", key=f"{prefix}_reset", use_container_width=True):
            for k in ["activity","provider","zone","woreda","mtype"]:
                st.session_state.pop(f"{prefix}_{k}", None)
            st.rerun()
    return activity, provider, zone, woreda, mtype


def row_matches_filters(r, activity="All", provider="All", zone="All", woreda="All", mtype="All"):
    a = ACTIVITY_LABELS.get(str(r.get("activity") or "").strip(), str(r.get("activity") or "").strip())
    p_raw = str(r.get("provider") or "").strip(); p_low=p_raw.casefold()
    p = "TPM" if ("third party" in p_low or "wfma" in p_low) else "WFP" if "wfp" in p_low else p_raw
    return (activity=="All" or a==activity) and (provider=="All" or p==provider) and (zone=="All" or str(r.get("zone") or "").strip()==zone) and (woreda=="All" or str(r.get("woreda") or "").strip()==woreda) and (mtype=="All" or str(r.get("monitoring_type") or "").strip()==mtype)


@st.cache_data(show_spinner=False)
def cached_raw_preview(path: str, sheet: str | None, source_type: str, sub_office: str | None = None):
    if source_type == "MoDa":
        return preview_moda_core(path, 1000, sub_office=sub_office)
    return preview_file(path, sheet, 1000)


@st.cache_data(show_spinner=False)
def cached_available_periods(moda_paths: tuple[str, ...]):
    return pd.DataFrame(available_periods(list(moda_paths)))


# -----------------------------------------------------------------------------
# STATE / SIDEBAR
# -----------------------------------------------------------------------------
if "workdir" not in st.session_state:
    st.session_state.workdir = tempfile.mkdtemp(prefix="jijiga_monitoring_")
if "analysis" not in st.session_state:
    st.session_state.analysis = None
if "upload_signature" not in st.session_state:
    st.session_state.upload_signature = None
if "generated_outputs" not in st.session_state:
    st.session_state.generated_outputs = None

# Build the global Sub-office filter from the currently uploaded MoDa source(s).
# Before any upload, show the three Somali Region operational offices as sensible defaults.
persisted_profiles = st.session_state.get("profiles", [])
persisted_moda_paths = tuple(p.get("path") for p in persisted_profiles if p.get("source_type") == "MoDa" and p.get("path"))
detected_suboffices = []
if persisted_moda_paths:
    try:
        _avail_sidebar = cached_available_periods(persisted_moda_paths)
        if not _avail_sidebar.empty:
            detected_suboffices = sorted(_avail_sidebar["Sub-office"].dropna().astype(str).str.strip().loc[lambda x: x.ne("")].unique().tolist())
    except Exception:
        detected_suboffices = []

if not detected_suboffices:
    detected_suboffices = ["Jijiga", "Gode", "Dollo Addo"]
else:
    # Keep Somali Region offices at the top while retaining all offices detected in an all-Ethiopia MoDa export.
    preferred = ["Jijiga", "Gode", "Dollo Addo"]
    ordered = [x for x in preferred if x in detected_suboffices]
    ordered.extend([x for x in detected_suboffices if x not in ordered])
    detected_suboffices = ordered

_previous_so = st.session_state.get("sub_office_filter", "Jijiga")
_default_so = _previous_so if _previous_so in detected_suboffices else ("Jijiga" if "Jijiga" in detected_suboffices else detected_suboffices[0])

with st.sidebar:
    render_sidebar_brand()
    # Keep the reporting context in a compact sidebar expander. The principal
    # analytical pages repeat the filters horizontally to match the dashboard UI.
    with st.expander("Reporting context", expanded=False):
        reporting_month = st.text_input("Reporting month", value=st.session_state.get("reporting_month_ui", "2026-08"), help="YYYY-MM; reporting period is derived from monitoring date.", key="reporting_month_ui")
        sub_office = st.selectbox(
            "Sub-office", detected_suboffices, index=detected_suboffices.index(_default_so),
            key="sub_office_filter", help="Options are detected from uploaded MoDa file(s)."
        )
    nav = st.radio(
        "Navigation",
        [
            "Home",
            "Data Upload & Setup",
            "Data Quality & Matching",
            "Raw Data Explorer",
            "Monitoring Overview",
            "Visited Sites Map",
            "AAP / CFM",
            "Protection, Safety & Dignity",
            "Activity Analysis",
            "Findings & Actions",
            "Photo Evidence",
            "Reports & Exports",
            "About",
        ],
        index=0, label_visibility="collapsed",
    )
    st.markdown('<div class="sidebar-footer"><b>Zero Hunger<br>for a Better Tomorrow</b><br><br>J-MAAP v1.8.1</div>', unsafe_allow_html=True)

# Changing the selected reporting cohort invalidates any previously calculated run.
current_params = (reporting_month.strip(), sub_office.strip())
if st.session_state.get("analysis") is not None and st.session_state.get("analysis_params") != current_params:
    st.session_state.analysis = None
    st.session_state.generated_outputs = None

branded_header(reporting_month)

# -----------------------------------------------------------------------------
# UPLOAD / PROFILING (persistent control at top)
# -----------------------------------------------------------------------------
uploads = st.file_uploader(
    "Upload monthly raw files",
    type=["xlsx", "xlsm", "csv"],
    accept_multiple_files=True,
    help="Upload one or more MoDa exports. RBMF, FRN, logistics/handover and previous action trackers are optional and will be profiled separately.",
    label_visibility="collapsed" if nav != "Data Upload & Setup" else "visible",
)

# Restore persisted upload/profile state on every Streamlit rerun.
# Page navigation and widget changes rerun the whole script; the file uploader
# can still contain files while the upload signature is unchanged.
profiles = st.session_state.get("profiles", [])
paths = st.session_state.get("paths", [])

if uploads:
    signature = tuple((f.name, getattr(f, "size", len(f.getbuffer()))) for f in uploads)
    if signature != st.session_state.upload_signature:
        st.session_state.workdir = tempfile.mkdtemp(prefix="jijiga_monitoring_")
        paths = save_uploads(uploads, st.session_state.workdir)
        profiles = []
        with st.spinner("Profiling uploaded files and building matching diagnostics..."):
            for p in paths:
                try:
                    profiles.append(profile_file(p))
                except Exception as e:
                    st.error(f"Could not profile {Path(p).name}: {e}")
            st.session_state.profiles = profiles
            st.session_state.paths = paths
            st.session_state.pairs = pairwise_matching(profiles)
            st.session_state.uuid_mat, st.session_state.site_mat = moda_file_overlap(profiles)
            st.session_state.unmatched = unmatched_sites(profiles)
        st.session_state.analysis = None
        st.session_state.analysis_params = None
        st.session_state.generated_outputs = None
        st.session_state.upload_signature = signature
        # Refresh once so the sidebar Sub-office dropdown immediately reflects the uploaded MoDa file(s).
        st.rerun()

    # Explicitly reload persisted values even when the signature did not change.
    profiles = st.session_state.get("profiles", [])
    paths = st.session_state.get("paths", [])

# -----------------------------------------------------------------------------
# HOME
# -----------------------------------------------------------------------------
if nav == "Home":
    page_heading("Monitoring Analytics Control Centre", "A transparent monthly workflow from raw submissions to validated management action.")
    if profiles:
        moda_count = sum(p["source_type"] == "MoDa" for p in profiles)
        a = st.session_state.analysis
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Files loaded", len(profiles))
        c2.metric("MoDa exports", moda_count)
        c3.metric("Sub-office", sub_office)
        c4.metric("Analysis status", "Ready" if a else "Analysis pending")
        c5.metric("Reporting month", reporting_month)
        if a:
            dq = a["dq"]
            st.markdown('<div class="status-ok">Monthly analysis is available. Review the DQ exceptions and field-validation status before generating final outputs.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-info">Files are loaded. Continue to Data Quality & Matching before running the monthly analysis.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="status-info">Upload the monthly raw files to begin. MoDa exports drive the monitoring analysis; operational files are matched diagnostically.</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Monthly workflow</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    steps = [
        ("1", "Upload", "Load raw MoDa and optional operational sources."),
        ("2", "Validate", "Inspect schema, duplicates, date logic and source matching."),
        ("3", "Analyse", "Calculate indicators, thematic signals and site-level findings."),
        ("4", "Act", "Validate findings, aggregate actions and generate management outputs."),
    ]
    for col, (n, t, d) in zip(cols, steps):
        with col:
            st.markdown(f'<div class="card"><span class="step-num">{n}</span><b>{t}</b><div class="muted" style="margin-top:8px">{d}</div></div>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# DATA UPLOAD
# -----------------------------------------------------------------------------
elif nav == "Data Upload & Setup":
    page_heading("Data Upload & Setup", "Load all monthly source files and confirm that the expected inputs have been recognised.")
    if not profiles:
        st.info("Upload one or more raw files using the uploader above.")
    else:
        inv = pd.DataFrame([{k: p.get(k) for k in ["file", "source_type", "sheet", "preview_rows", "columns", "site_column", "uuid_unique_preview"]} for p in profiles])
        moda_count = sum(p["source_type"] == "MoDa" for p in profiles)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Files uploaded", len(profiles))
        c2.metric("MoDa exports", moda_count)
        c3.metric("Other sources", len(profiles) - moda_count)
        c4.metric("Reporting month", reporting_month)
        st.dataframe(inv, use_container_width=True, hide_index=True)
        if moda_count:
            moda_paths_upload = tuple(p["path"] for p in profiles if p["source_type"] == "MoDa")
            try:
                _av = cached_available_periods(moda_paths_upload)
                if not _av.empty:
                    _so = sorted(_av["Sub-office"].dropna().astype(str).unique().tolist())
                    st.caption("Detected sub-offices: " + ", ".join(_so))
            except Exception:
                pass
            st.markdown('<div class="status-ok">MoDa source detected. Continue to Data Quality & Matching before running analysis.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-warn">No MoDa source detected. DQ profiling is available, but findings generation requires MoDa data.</div>', unsafe_allow_html=True)
        with st.expander("Column / schema inventory"):
            for p in profiles:
                st.markdown(f"**{p['file']} — {p['source_type']}**")
                st.write(p.get("column_names", []))

# -----------------------------------------------------------------------------
# DQ & MATCHING
# -----------------------------------------------------------------------------
elif nav == "Data Quality & Matching":
    page_heading("Data Quality & Matching", "Check source integrity, overlap and match quality before allowing the monthly analysis to proceed.")
    if not profiles:
        st.warning("Upload source files first.")
    else:
        moda_paths = [p["path"] for p in profiles if p["source_type"] == "MoDa"]
        availability = cached_available_periods(tuple(moda_paths)) if moda_paths else pd.DataFrame()
        eligible = 0
        available_months = []
        if not availability.empty:
            so_avail = availability[availability["Sub-office"].astype(str).str.casefold() == sub_office.strip().casefold()].copy()
            available_months = so_avail["Reporting month"].astype(str).tolist()
            match = so_avail[so_avail["Reporting month"] == reporting_month.strip()]
            eligible = int(match["Records"].sum()) if not match.empty else 0

            st.markdown('<div class="section-title">Analysis cohort availability</div>', unsafe_allow_html=True)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Selected sub-office", sub_office)
            c2.metric("Selected reporting month", reporting_month)
            c3.metric("Eligible MoDa records", eligible)
            c4.metric("Detected months", len(available_months))
            st.dataframe(so_avail, use_container_width=True, hide_index=True)

            if eligible == 0:
                if available_months:
                    labels = ", ".join(available_months)
                    st.markdown(f'<div class="status-warn"><b>No eligible {sub_office} records were found for {reporting_month}.</b> Available actual-date periods in the uploaded MoDa file(s): {labels}. Change the Reporting month in the sidebar before running the analysis.</div>', unsafe_allow_html=True)
                else:
                    offices = ", ".join(sorted(availability["Sub-office"].astype(str).unique().tolist())[:20])
                    st.markdown(f'<div class="status-warn"><b>No records were found for sub-office “{sub_office}”.</b> Detected sub-offices include: {offices}.</div>', unsafe_allow_html=True)
            else:
                row = match.iloc[0]
                st.markdown(f'<div class="status-ok">Cohort is valid: <b>{eligible:,}</b> {sub_office} records dated {row["First date"]} to {row["Last date"]} are eligible for this run.</div>', unsafe_allow_html=True)
                try:
                    per = pd.Period(reporting_month.strip(), freq="M")
                    month_start = per.start_time.date().isoformat()
                    month_end = per.end_time.date().isoformat()
                    if str(row["First date"]) > month_start or str(row["Last date"]) < month_end:
                        st.markdown(f'<div class="status-warn"><b>Possible partial-month source:</b> the selected cohort covers {row["First date"]} to {row["Last date"]}, while the calendar month is {month_start} to {month_end}. If this is a completed reporting month, upload the additional MoDa export/form version covering the missing dates. J-MAAP will combine and UUID-deduplicate multiple exports.</div>', unsafe_allow_html=True)
                except Exception:
                    pass

        pairs = st.session_state.get("pairs", pd.DataFrame())
        if not pairs.empty:
            st.markdown('<div class="section-title">Cross-file reconciliation</div>', unsafe_allow_html=True)
            st.dataframe(pairs, use_container_width=True, hide_index=True)
        uuid_mat, site_mat = st.session_state.get("uuid_mat", pd.DataFrame()), st.session_state.get("site_mat", pd.DataFrame())
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="section-title">MoDa UUID overlap</div>', unsafe_allow_html=True)
            st.dataframe(uuid_mat, use_container_width=True) if not uuid_mat.empty else st.caption("No comparable MoDa UUID matrix available.")
        with c2:
            st.markdown('<div class="section-title">Normalized site overlap</div>', unsafe_allow_html=True)
            st.dataframe(site_mat, use_container_width=True) if not site_mat.empty else st.caption("No comparable site matrix available.")

        um = st.session_state.get("unmatched", pd.DataFrame())
        if not um.empty:
            st.markdown('<div class="section-title">Optional-source site matching against MoDa</div>', unsafe_allow_html=True)
            view = um.copy()
            if "Match %" in view.columns:
                view["Match %"] = view["Match %"].map(fmt_pct)
            st.dataframe(view, use_container_width=True, hide_index=True)
        st.caption("Exact site-name matching is intentionally conservative. Production reconciliation should use the canonical Site Crosswalk for spelling variants and corporate IDs.")

        can_run = bool(moda_paths) and (availability.empty or eligible > 0)
        if not can_run and moda_paths:
            st.button("Run controlled monthly analysis", type="primary", use_container_width=True, disabled=True, help="Select a reporting month/sub-office with eligible records first.")
        elif moda_paths and st.button("Run controlled monthly analysis", type="primary", use_container_width=True):
            with st.spinner("Running controlled analysis and DQ rules..."):
                st.session_state.analysis = run_analysis(moda_paths, reporting_month.strip(), sub_office.strip())
                st.session_state.analysis_params = (reporting_month.strip(), sub_office.strip())
                st.session_state.generated_outputs = None
            st.success("Analysis completed. Review DQ exceptions below before relying on findings.")

        if st.session_state.analysis:
            dq = st.session_state.analysis["dq"]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Accepted submissions", dq["rows"])
            c2.metric("Unique UUIDs", dq["uuid_unique"])
            c3.metric("Duplicates removed", dq["duplicate_uuid_rows"])
            c4.metric("DQ issues", len(dq["issues"]))
            st.markdown('<div class="section-title">Automated DQ exceptions</div>', unsafe_allow_html=True)
            st.dataframe(pd.DataFrame(dq["issues"]), use_container_width=True, hide_index=True)
            scoped = scoped_source_site_matching(profiles, st.session_state.analysis["rows"])
            if not scoped.empty:
                scoped_view = scoped.copy()
                scoped_view["Match %"] = scoped_view["Match %"].map(fmt_pct)
                st.markdown('<div class="section-title">Matching to selected month / sub-office cohort</div>', unsafe_allow_html=True)
                st.dataframe(scoped_view, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# RAW EXPLORER
# -----------------------------------------------------------------------------
elif nav == "Raw Data Explorer":
    page_heading("Raw Data Explorer", "Explore uploaded records before analysis. Filters and previews are designed for transparency and spot-checking.")
    if not profiles:
        st.warning("Upload source files first.")
    else:
        a = st.session_state.analysis
        if a:
            rows = a["rows"]
            m = overview_metrics(rows)
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Total submissions", m["submissions"])
            c2.metric("Sites monitored", m["sites"])
            c3.metric("Woredas", m["woredas"])
            c4.metric("Activities", m["activities"])
            df_a = rows_to_df(rows)
            c5.metric("WFP submissions", int((df_a.get("provider", pd.Series(dtype=str)).astype(str).str.upper() == "WFP").sum()))
            c6.metric("TPM submissions", int((df_a.get("provider", pd.Series(dtype=str)).astype(str).str.upper() == "TPM").sum()))

        chosen = st.selectbox("Source file", [p["file"] for p in profiles])
        p = next(x for x in profiles if x["file"] == chosen)
        sheets = list_xlsx_sheets(p["path"])
        sheet = st.selectbox("Worksheet", sheets, index=(sheets.index("data") if "data" in sheets else 0)) if sheets else None
        dfraw = cached_raw_preview(p["path"], sheet, p["source_type"], sub_office=sub_office).copy()

        q = st.text_input("Search displayed columns", value="", placeholder="Type a value, site, woreda, activity or UUID fragment...")
        if q and not dfraw.empty:
            mask = dfraw.astype(str).apply(lambda s: s.str.contains(q, case=False, na=False)).any(axis=1)
            dfraw = dfraw[mask]

        if p["source_type"] == "MoDa":
            st.caption(f"Preview filtered to Sub-office: {sub_office}. For performance, the preview exposes core raw fields only; the complete workbook remains the source used by the analysis engine.")

        if not dfraw.empty:
            # Lightweight visual preview from displayed core fields when recognizable.
            date_col = next((c for c in dfraw.columns if str(c).lower() in {"monitoring_date", "date", "start"}), None)
            activity_col = next((c for c in dfraw.columns if "activity" in str(c).lower()), None)
            woreda_col = next((c for c in dfraw.columns if "woreda" in str(c).lower()), None)
            chart_cols = st.columns(3)
            if date_col:
                tmp = pd.to_datetime(dfraw[date_col], errors="coerce").dt.date.value_counts().sort_index().reset_index()
                tmp.columns = ["Date", "Records"]
                with chart_cols[0]: st.plotly_chart(px.bar(tmp, x="Date", y="Records", title="Records by date"), use_container_width=True, config={"displayModeBar":False})
            if activity_col:
                tmp = dfraw[activity_col].astype(str).value_counts().head(10).reset_index(); tmp.columns=["Activity","Records"]
                with chart_cols[1]: st.plotly_chart(px.pie(tmp, values="Records", names="Activity", hole=.5, title="Activity composition"), use_container_width=True, config={"displayModeBar":False})
            if woreda_col:
                tmp = dfraw[woreda_col].astype(str).value_counts().head(10).reset_index(); tmp.columns=["Woreda","Records"]
                with chart_cols[2]: st.plotly_chart(px.bar(tmp.sort_values("Records"), x="Records", y="Woreda", orientation="h", title="Top woredas"), use_container_width=True, config={"displayModeBar":False})

        st.markdown('<div class="section-title">Raw data preview</div>', unsafe_allow_html=True)
        st.dataframe(dfraw, use_container_width=True, height=540, hide_index=True)
        st.download_button("Download current preview as CSV", dfraw.to_csv(index=False).encode("utf-8-sig"), file_name=f"{Path(chosen).stem}_preview.csv", mime="text/csv")

# -----------------------------------------------------------------------------
# MONITORING OVERVIEW
# -----------------------------------------------------------------------------
elif nav == "Monitoring Overview":
    a = st.session_state.analysis
    if not a:
        page_heading("Monitoring Overview", "Overview of monitoring coverage, submissions and key trends across the selected operational cohort.")
        st.warning("Run the controlled monthly analysis from Data Quality & Matching.")
    else:
        rows = a["rows"]
        activity, provider, zone, woreda, mtype = dashboard_filter_row(rows, "overview")
        filtered_rows = [r for r in rows if row_matches_filters(r, activity, provider, zone, woreda, mtype)]
        df = rows_to_df(filtered_rows)
        detail = pd.DataFrame(a.get("detail", []))
        if not detail.empty:
            if activity != "All": detail = detail[detail["Activity"].astype(str).str.contains(activity, case=False, na=False)]
            if zone != "All": detail = detail[detail["Zone"].astype(str).eq(zone)]
            if woreda != "All": detail = detail[detail["Woreda"].astype(str).eq(woreda)]

        page_heading("Monitoring Overview", "Overview of monitoring coverage, submissions and key trends across Somali Region.")
        if df.empty:
            st.info("No monitoring records match the current filters.")
        else:
            provider_std = df["provider"].astype(str).map(lambda x: "TPM" if ("third party" in x.casefold() or "wfma" in x.casefold()) else "WFP" if "wfp" in x.casefold() else x)
            event_cols = [c for c in ["monitoring_date_iso","site","activity","provider"] if c in df.columns]
            events = int(df[event_cols].drop_duplicates().shape[0]) if event_cols else len(df)
            sites = int(df["site"].replace("", pd.NA).nunique())
            woredas_n = int(df["woreda"].replace("", pd.NA).nunique())
            wfp_n = int((provider_std == "WFP").sum()); tpm_n = int((provider_std == "TPM").sum())
            total_p = max(wfp_n+tpm_n,1)
            high = int(detail["Severity"].isin(["Critical","High"]).sum()) if not detail.empty and "Severity" in detail.columns else 0
            cards = st.columns(7)
            with cards[0]: kpi("Sites monitored", f"{sites:,}", "●", "blue")
            with cards[1]: kpi("Submissions", f"{len(df):,}", "▤", "green")
            with cards[2]: kpi("Monitoring events", f"{events:,}", "♟", "purple")
            with cards[3]: kpi("Woredas covered", f"{woredas_n:,}", "⌂", "orange")
            with cards[4]: kpi("WFP share", f"{100*wfp_n/total_p:.0f}%", "◔", "blue")
            with cards[5]: kpi("TPM share", f"{100*tpm_n/total_p:.0f}%", "♣", "teal")
            with cards[6]: kpi("High-priority findings", f"{high:,}", "!", "red", "pending validation")

            main, rail = st.columns([4.4, 1.35])
            with main:
                r1 = st.columns([1.35,1.12,.88])
                with r1[0]:
                    daily = df.copy()
                    daily["_date"] = pd.to_datetime(daily["monitoring_date_iso"], errors="coerce")
                    daily = daily.dropna(subset=["_date"])
                    daily["Provider"] = provider_std.loc[daily.index]
                    trend = daily.groupby([daily["_date"].dt.date,"Provider"]).size().reset_index(name="Submissions")
                    trend.columns=["Date","Provider","Submissions"]
                    fig=px.line(trend,x="Date",y="Submissions",color="Provider",markers=True,title="Monitoring Submissions Trend",color_discrete_map={"WFP":"#1685D3","TPM":"#23A567"})
                    st.plotly_chart(clean_plot(fig, 290), use_container_width=True, config={"displayModeBar":False})
                with r1[1]:
                    act = df["activity"].astype(str).map(lambda x: ACTIVITY_LABELS.get(x,x)).value_counts().reset_index(); act.columns=["Activity","Submissions"]
                    fig=px.bar(act,x="Activity",y="Submissions",title="Submissions by Activity",color="Activity",color_discrete_sequence=["#268BE0","#EF4E5F","#7B5CD6","#2DA45D","#F29432"])
                    st.plotly_chart(clean_plot(fig,290,legend=False),use_container_width=True,config={"displayModeBar":False})
                with r1[2]:
                    pp=pd.DataFrame({"Provider":["WFP","TPM"],"Submissions":[wfp_n,tpm_n]})
                    fig=px.pie(pp,values="Submissions",names="Provider",hole=.58,title="WFP vs TPM Share",color="Provider",color_discrete_map={"WFP":"#268BE0","TPM":"#23A68B"})
                    fig.update_traces(textinfo="percent")
                    st.plotly_chart(clean_plot(fig,290),use_container_width=True,config={"displayModeBar":False})

                r2=st.columns(3)
                with r2[0]:
                    z=(df.dropna(subset=["zone"]).groupby("zone")["site"].nunique().sort_values(ascending=False).head(10).reset_index(name="Sites"))
                    fig=px.bar(z.sort_values("Sites"),x="Sites",y="zone",orientation="h",title="Visited Sites by Zone")
                    fig.update_traces(marker_color="#318CDC")
                    st.plotly_chart(clean_plot(fig,270,legend=False),use_container_width=True,config={"displayModeBar":False})
                with r2[1]:
                    mt=df["monitoring_type"].replace("", "Unspecified").value_counts().head(6).reset_index(); mt.columns=["Monitoring Type","Records"]
                    fig=px.pie(mt,values="Records",names="Monitoring Type",hole=.50,title="Monitoring Type Distribution")
                    st.plotly_chart(clean_plot(fig,270),use_container_width=True,config={"displayModeBar":False})
                with r2[2]:
                    zones_n=int(df["zone"].replace("",pd.NA).nunique())
                    cov_html=f'''<div class="panel" style="height:270px"><div class="panel-title">Coverage Summary</div>
                    <div style="font-size:11px;line-height:2.05;color:#476781">
                    <b style="color:#0A4D82">▰</b>&nbsp; Zones covered <b style="float:right">{zones_n}</b><br>
                    <b style="color:#0A4D82">▱</b>&nbsp; Woredas covered <b style="float:right">{woredas_n}</b><br>
                    <b style="color:#0A4D82">●</b>&nbsp; Sites monitored <b style="float:right">{sites}</b><br>
                    <b style="color:#0A4D82">♟</b>&nbsp; Monitoring events <b style="float:right">{events}</b><br>
                    <b style="color:#0A4D82">◫</b>&nbsp; Reporting month <b style="float:right">{html.escape(reporting_month)}</b>
                    </div></div>'''
                    st.markdown(cov_html,unsafe_allow_html=True)
            with rail:
                act_counts=df["activity"].astype(str).map(lambda x:ACTIVITY_LABELS.get(x,x)).value_counts()
                lead_act=act_counts.index[0] if len(act_counts) else "—"
                lead_share=(100*act_counts.iloc[0]/len(df)) if len(act_counts) else 0
                insight=[
                    ("Monitoring footprint",f"{sites} sites and {events} monitoring events are represented in the selected cohort."),
                    ("Largest activity share",f"{lead_act} accounts for {lead_share:.0f}% of current submissions."),
                    ("Provider composition",f"WFP contributes {100*wfp_n/total_p:.0f}% and TPM {100*tpm_n/total_p:.0f}% of the current evidence base."),
                    ("Geographic coverage",f"Monitoring reached {woredas_n} woredas across {int(df['zone'].replace('',pd.NA).nunique())} zones."),
                    ("High-priority findings",f"{high} current detailed finding records are classified High/Critical and require validation/follow-up."),
                ]
                body=''.join(f'<div class="rail-item"><div style="width:25px;height:25px;border-radius:50%;background:#E6F3FF;color:#1377BF;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:11px">{i}</div><div class="rail-main"><b>{html.escape(t)}</b><div class="rail-meta">{html.escape(x)}</div></div></div>' for i,(t,x) in enumerate(insight,1))
                st.markdown(f'<div class="rail-card"><div class="rail-head">Insights & Highlights</div>{body}</div>',unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# VISITED SITES MAP
# -----------------------------------------------------------------------------
elif nav == "Visited Sites Map":
    a = st.session_state.analysis
    if not a:
        page_heading("Visited Sites Map", "Interactive monitoring footprint with site evidence, administrative boundaries and MoDa photos.")
        st.warning("Run the controlled monthly analysis first. The map uses the accepted cohort, MoDa GPS coordinates and available photo links.")
    else:
        rows = a["rows"]
        activity, provider, zone, woreda, mtype = dashboard_filter_row(rows, "map")
        filtered_rows = [r for r in rows if row_matches_filters(r, activity, provider, zone, woreda, mtype)]
        filtered_photos = apply_photo_filters(
            a.get("photos", []),
            activities=[] if activity=="All" else [activity],
            zones=[] if zone=="All" else [zone],
            woredas=[] if woreda=="All" else [woreda],
            providers=[] if provider=="All" else [provider],
        )
        detail = a.get("detail", [])
        detail_filtered=[]
        for d in detail:
            da=str(d.get("Activity") or "")
            dz=str(d.get("Zone") or ""); dw=str(d.get("Woreda") or "")
            if activity!="All" and activity.casefold() not in da.casefold(): continue
            if zone!="All" and dz!=zone: continue
            if woreda!="All" and dw!=woreda: continue
            detail_filtered.append(d)
        points = build_site_points(filtered_rows, detail_filtered, filtered_photos)

        page_heading("Visited Sites Map", "View monitored sites, visit frequency, findings and photo evidence in geographic context.")
        valid_submission_count = int(points["submissions"].sum()) if not points.empty else 0
        total_filtered = len(filtered_rows)
        events = int(points["events"].sum()) if not points.empty else 0
        woredas_n = int(points["woreda"].replace("", pd.NA).nunique()) if not points.empty else 0
        zones_n = int(points["zone"].replace("", pd.NA).nunique()) if not points.empty else 0
        photo_sites = int((points["photo_count"]>0).sum()) if not points.empty and "photo_count" in points.columns else 0
        geocoded = 100*valid_submission_count/max(total_filtered,1)
        cards=st.columns(7)
        with cards[0]: kpi("Sites monitored", f"{len(points):,}", "●", "blue")
        with cards[1]: kpi("Submissions", f"{total_filtered:,}", "▤", "green")
        with cards[2]: kpi("Monitoring events", f"{events:,}", "♟", "purple")
        with cards[3]: kpi("Woredas covered", f"{woredas_n:,}", "⌂", "orange")
        with cards[4]: kpi("Zones covered", f"{zones_n:,}", "▱", "teal")
        with cards[5]: kpi("Sites with photos", f"{photo_sites:,}", "📷", "red")
        with cards[6]: kpi("Geocoded", f"{geocoded:.0f}%", "◎", "blue")

        ctl=st.columns([1.3,1,3.7])
        with ctl[0]: display_mode=st.selectbox("Marker display", ["Activity","Visit frequency","Provider","Finding severity"], key="map_display_mode")
        with ctl[1]: show_woreda_labels=st.checkbox("Woreda labels", value=False, key="map_woreda_labels")
        with ctl[2]: st.caption("Click any site marker to update the Site Details panel. Use the map buttons to switch between Street, Satellite / Imagery and the Somali Region hollow administrative map.")

        if points.empty:
            st.warning("No valid MoDa GPS coordinates were found for the current filters. Review the raw coordinates or broaden the filters.")
        else:
            assets=load_admin_assets(APP_DIR)
            html_map=leaflet_dashboard_html(points, assets, initial_basemap="Street", display_mode=display_mode, show_woreda_labels=show_woreda_labels)
            components.html(html_map, height=685, scrolling=False)
            st.caption("Street: © OpenStreetMap contributors. Satellite imagery: © Esri and contributors. Administrative boundaries: Somali Region 2024 shapefile supplied for J-MAAP. Photo links remain hosted by MoDa and follow the source system's access controls.")

            st.markdown('<div class="section-title">Monitored Sites (Filtered View)</div>', unsafe_allow_html=True)
            q=st.text_input("Search site, woreda or zone", value="", placeholder="Search site, woreda, zone...", key="map_site_search", label_visibility="collapsed")
            table=points[["site","woreda","zone","activities","providers","events","submissions","latest_date","finding_signals","photo_count"]].copy()
            table.columns=["Site name","Woreda","Zone","Activity","Provider","Monitoring events","Submissions","Latest visit","Findings","Photos"]
            if q:
                mask=table.astype(str).apply(lambda x:x.str.contains(q,case=False,na=False)).any(axis=1)
                table=table[mask]
            st.dataframe(table,use_container_width=True,height=285,hide_index=True)
            st.download_button("Export mapped sites", table.to_csv(index=False).encode("utf-8-sig"), f"{reporting_month}_{sub_office.replace(' ','_')}_Mapped_Sites.csv", "text/csv", use_container_width=False)

# -----------------------------------------------------------------------------
# AAP / CFM
# -----------------------------------------------------------------------------
elif nav == "AAP / CFM":
    page_heading("AAP / CFM", "Track the community feedback journey from awareness and access through usage, response and satisfaction.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        rows = a["rows"]
        act_label = st.selectbox("Activity", ["Activity 1 (Relief response)", "Activity 2 (Nutrition assistance)", "Activity 3 (Refugee operations)", "Activity 6 (resilience)"], key="cfmact")
        res = cfm_journey(rows, activity=act_label)
        fdf = pd.DataFrame(res["stages"])
        if res["base"]:
            fdf["Percent of base"] = fdf["pct_base"] * 100
            c1, c2 = st.columns([1.15, 1])
            with c1:
                st.plotly_chart(px.funnel(fdf, x="count", y="stage", title=f"CFM journey — base n={res['base']}"), use_container_width=True)
            with c2:
                stage_cards = "".join([f'<div class="step-card"><b>{r.stage}</b><span style="float:right;color:#006EA9;font-weight:700">{r["Percent of base"]:.0f}%</span><div class="muted">{int(r["count"])} records</div></div>' for _, r in fdf.iterrows()])
                st.markdown(stage_cards, unsafe_allow_html=True)
            st.caption("Response and Satisfaction are only shown where supported by the activity module; Response uses a strict nested rule to avoid skip-logic inflation.")

        if act_label == "Activity 1 (Relief response)":
            frames = []
            for mod in ["Food", "Cash"]:
                rr = cfm_journey(rows, activity=act_label, modality=mod)
                for x in rr["stages"]:
                    frames.append({"Modality": mod, "Stage": x["stage"], "Percent of base": 100 * x["pct_base"] if x["pct_base"] is not None else None, "Base": rr["base"]})
            mdf = pd.DataFrame(frames)
            st.plotly_chart(px.line(mdf, x="Stage", y="Percent of base", color="Modality", markers=True, title="Relief CFM journey — Food vs Cash"), use_container_width=True)
            st.caption("Interpret modality differences descriptively: the August cash sample is geographically concentrated and should not be treated as a causal modality effect.")

        heat = cfm_woreda_table(rows, act_label)
        if not heat.empty:
            stages = [x for x in ["Awareness", "Access", "Usage", "Response", "Satisfaction"] if x in heat.columns]
            long = heat.melt(id_vars=["Woreda", "N"], value_vars=stages, var_name="Stage", value_name="Rate").dropna()
            if not long.empty:
                pivot = long.pivot(index="Woreda", columns="Stage", values="Rate") * 100
                st.plotly_chart(px.imshow(pivot, text_auto=".0f", aspect="auto", color_continuous_scale="Blues", title="Woreda CFM stage heatmap (%)"), use_container_width=True)
                st.dataframe(heat, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# PROTECTION
# -----------------------------------------------------------------------------
elif nav == "Protection, Safety & Dignity":
    page_heading("Protection, Safety & Dignity", "Review protection, inclusion, integrity and dignity signals while separating prevalence from seriousness.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        pdf = protection_summary(a["indicators"])
        if not pdf.empty:
            view = pdf.copy(); view["Issue rate %"] = view["Issue rate"].map(lambda x: 100 * x if pd.notna(x) else None)
            st.plotly_chart(px.bar(view.sort_values("Issue rate %"), x="Issue rate %", y="Finding", color="Activity", orientation="h", hover_data=["Severity", "Applicable", "Affected sites"], title="Protection / integrity assurance signals"), use_container_width=True)
            st.dataframe(view[["Activity", "Theme", "Finding", "Issues", "Applicable", "Issue rate %", "Severity", "Affected sites"]], use_container_width=True, hide_index=True)
            st.info("Frequency and seriousness are interpreted separately. Rare payment, misconduct or stock-control signals remain high-priority verification items even when prevalence is low.")

# -----------------------------------------------------------------------------
# ACTIVITY ANALYSIS
# -----------------------------------------------------------------------------
elif nav == "Activity Analysis":
    page_heading("Activity Analysis", "Review configured thematic findings by programme activity before moving to management follow-up.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        idf = pd.DataFrame(a["indicators"])
        if idf.empty:
            st.info("No configured indicators available.")
        else:
            activity_col = "Activity" if "Activity" in idf.columns else ("activity" if "activity" in idf.columns else None)
            if activity_col:
                acts = sorted(idf[activity_col].dropna().astype(str).unique().tolist())
                act = st.selectbox("Activity", acts)
                view = idf[idf[activity_col].astype(str) == act].copy()
            else:
                view = idf.copy()
            rate_col = next((c for c in ["Issue_Rate", "Issue rate", "issue_rate"] if c in view.columns), None)
            if rate_col:
                view["Issue rate %"] = pd.to_numeric(view[rate_col], errors="coerce") * 100
                finding_col = next((c for c in ["Finding", "finding", "Indicator", "indicator"] if c in view.columns), view.columns[0])
                st.plotly_chart(px.bar(view.sort_values("Issue rate %"), x="Issue rate %", y=finding_col, orientation="h", title="Configured issue rates"), use_container_width=True)
            st.dataframe(view, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# FINDINGS & ACTIONS
# -----------------------------------------------------------------------------
elif nav == "Findings & Actions":
    a = st.session_state.analysis
    if not a:
        page_heading("Findings & Actions", "Track findings, manage actions and monitor progress towards resolution.")
        st.warning("Run the controlled monthly analysis first.")
    else:
        rows=a["rows"]
        activity, provider, zone, woreda, mtype=dashboard_filter_row(rows,"findings")
        filtered_rows=[r for r in rows if row_matches_filters(r,activity,provider,zone,woreda,mtype)]
        valid_sites={(str(r.get("site") or ""),str(r.get("woreda") or ""),str(r.get("zone") or "")) for r in filtered_rows}
        ddf=pd.DataFrame(a.get("detail",[])); adf=pd.DataFrame(a.get("agg",[]))
        if not ddf.empty:
            ddf=ddf[ddf.apply(lambda r:(str(r.get("Site") or ""),str(r.get("Woreda") or ""),str(r.get("Zone") or "")) in valid_sites,axis=1)]
        if not adf.empty:
            if activity!="All": adf=adf[adf["Activities"].astype(str).str.contains(activity,case=False,na=False)]
            if zone!="All": adf=adf[adf["Zones"].astype(str).str.contains(zone,case=False,na=False)]
            if woreda!="All": adf=adf[adf["Woredas"].astype(str).str.contains(woreda,case=False,na=False)]

        page_heading("Findings & Actions", "Track findings, manage actions and monitor progress towards resolution.")
        detailed_n=len(ddf); aggregate_n=len(adf)
        high_n=int(ddf["Severity"].isin(["Critical","High"]).sum()) if not ddf.empty else 0
        if not adf.empty:
            status=adf["Status"].fillna("Proposed").astype(str)
            open_n=int((~status.str.contains("closed",case=False,na=False)).sum())
            due=pd.to_datetime(adf["Due_Date"],errors="coerce")
            overdue_n=int(((due < pd.Timestamp.today().normalize()) & ~status.str.contains("closed",case=False,na=False)).sum())
            comp=pd.to_datetime(adf["Completion_Date"],errors="coerce")
            try: period=pd.Period(reporting_month,freq="M"); closed_month=int(((comp.dt.to_period("M")==period) & status.str.contains("closed",case=False,na=False)).sum())
            except Exception: closed_month=int(status.str.contains("closed",case=False,na=False).sum())
        else:
            open_n=overdue_n=closed_month=0
        cards=st.columns(6)
        with cards[0]: kpi("Detailed findings",f"{detailed_n:,}","⌕","blue")
        with cards[1]: kpi("Aggregated issues",f"{aggregate_n:,}","!","orange")
        with cards[2]: kpi("High priority findings",f"{high_n:,}","!","red")
        with cards[3]: kpi("Open actions",f"{open_n:,}","▤","green")
        with cards[4]: kpi("Overdue actions",f"{overdue_n:,}","◷","red")
        with cards[5]: kpi("Closed this month",f"{closed_month:,}","✓","purple")

        tabs=st.tabs(["Overview","Findings Analysis","Action Tracker","Issues by Woreda","Trends"])
        with tabs[0]:
            left,rail=st.columns([4.5,1.35])
            with left:
                charts=st.columns([1,1.1,1,1.35])
                with charts[0]:
                    sev=ddf["Severity"].fillna("Unspecified").value_counts().reindex(["Critical","High","Medium-High","Medium","Low"]).dropna().reset_index() if not ddf.empty else pd.DataFrame(columns=["Severity","Findings"])
                    if not sev.empty:
                        sev.columns=["Severity","Findings"]
                        fig=px.bar(sev,x="Severity",y="Findings",title="Findings by Severity",color="Severity",color_discrete_map={"Critical":"#F04452","High":"#FF7A1A","Medium-High":"#FF9E2A","Medium":"#F5C928","Low":"#20A75A"})
                        st.plotly_chart(clean_plot(fig,255,legend=False),use_container_width=True,config={"displayModeBar":False})
                with charts[1]:
                    theme=ddf["Theme"].fillna("Other").value_counts().reset_index() if not ddf.empty else pd.DataFrame(columns=["Theme","Findings"])
                    if not theme.empty:
                        theme.columns=["Theme","Findings"]
                        fig=px.pie(theme,values="Findings",names="Theme",hole=.58,title="Findings by Theme")
                        st.plotly_chart(clean_plot(fig,255),use_container_width=True,config={"displayModeBar":False})
                with charts[2]:
                    if not adf.empty:
                        ac=adf["Status"].fillna("Proposed").value_counts().reset_index(); ac.columns=["Status","Actions"]
                        fig=px.pie(ac,values="Actions",names="Status",hole=.56,title="Action Status")
                        st.plotly_chart(clean_plot(fig,255),use_container_width=True,config={"displayModeBar":False})
                with charts[3]:
                    if not ddf.empty:
                        heat=pd.crosstab(ddf["Woreda"],ddf["Theme"])
                        if not heat.empty:
                            top_w=heat.sum(axis=1).sort_values(ascending=False).head(8).index
                            top_t=heat.sum(axis=0).sort_values(ascending=False).head(5).index
                            heat=heat.loc[top_w,top_t]
                            fig=px.imshow(heat,text_auto=True,aspect="auto",color_continuous_scale="Blues",title="Findings by Woreda and Theme")
                            st.plotly_chart(clean_plot(fig,255,legend=False),use_container_width=True,config={"displayModeBar":False})
                st.markdown('<div style="height:4px"></div>',unsafe_allow_html=True)
                search=st.text_input("Search issue, woreda, activity or owner",value="",placeholder="Search issue, woreda, activity, owner...",key="finding_search",label_visibility="collapsed")
                table_df=ddf.copy()
                if search and not table_df.empty:
                    mask=table_df.astype(str).apply(lambda x:x.str.contains(search,case=False,na=False)).any(axis=1); table_df=table_df[mask]
                st.markdown(findings_action_table(table_df,max_rows=12),unsafe_allow_html=True)
            with rail:
                st.markdown(priority_actions_html(adf,max_items=5),unsafe_allow_html=True)
                st.markdown(recurring_issues_html(ddf,max_items=5),unsafe_allow_html=True)
        with tabs[1]:
            if ddf.empty: st.info("No detailed findings under the current filters.")
            else:
                c1,c2=st.columns(2)
                with c1:
                    t=ddf.groupby("Theme").agg(Findings=("Action_ID","count"),Affected_sites=("Site","nunique")).reset_index().sort_values("Findings",ascending=False)
                    fig=px.bar(t.sort_values("Findings"),x="Findings",y="Theme",orientation="h",title="Findings concentration by theme")
                    st.plotly_chart(clean_plot(fig,350,legend=False),use_container_width=True,config={"displayModeBar":False})
                with c2:
                    w=ddf.groupby("Woreda").size().sort_values(ascending=False).head(15).reset_index(name="Findings")
                    fig=px.bar(w.sort_values("Findings"),x="Findings",y="Woreda",orientation="h",title="Top woredas by detailed finding records")
                    st.plotly_chart(clean_plot(fig,350,legend=False),use_container_width=True,config={"displayModeBar":False})
                st.dataframe(ddf,use_container_width=True,height=420,hide_index=True)
        with tabs[2]:
            if adf.empty: st.info("No aggregated actions under the current filters.")
            else:
                st.dataframe(adf,use_container_width=True,height=520,hide_index=True)
        with tabs[3]:
            if ddf.empty: st.info("No woreda findings under the current filters.")
            else:
                heat=pd.crosstab(ddf["Woreda"],ddf["Theme"])
                fig=px.imshow(heat,text_auto=True,aspect="auto",color_continuous_scale="Blues",title="Woreda × thematic finding matrix")
                st.plotly_chart(clean_plot(fig,520,legend=False),use_container_width=True,config={"displayModeBar":False})
        with tabs[4]:
            if ddf.empty: st.info("No findings available for trend analysis.")
            else:
                tmp=ddf.copy(); tmp["Date"]=pd.to_datetime(tmp["First_Observed"],errors="coerce")
                tmp=tmp.dropna(subset=["Date"]).groupby(tmp["Date"].dt.date).size().reset_index(name="New findings")
                fig=px.line(tmp,x="Date",y="New findings",markers=True,title="New finding records by first observed date")
                st.plotly_chart(clean_plot(fig,380,legend=False),use_container_width=True,config={"displayModeBar":False})

# -----------------------------------------------------------------------------
# PHOTO EVIDENCE
# -----------------------------------------------------------------------------
elif nav == "Photo Evidence":
    a=st.session_state.analysis
    if not a:
        page_heading("Photo Evidence", "Review MoDa image evidence linked to monitoring sites and submissions.")
        st.warning("Run the controlled monthly analysis first.")
    else:
        rows=a["rows"]
        activity, provider, zone, woreda, mtype=dashboard_filter_row(rows,"photos")
        photos=apply_photo_filters(a.get("photos",[]), activities=[] if activity=="All" else [activity], zones=[] if zone=="All" else [zone], woredas=[] if woreda=="All" else [woreda], providers=[] if provider=="All" else [provider])
        page_heading("Photo Evidence", "Review field imagery together with its site, activity, date and submission traceability.")
        sites=sorted({str(x.get("site") or "").strip() for x in photos if str(x.get("site") or "").strip()})
        types=sorted({str(x.get("photo_type") or "Monitoring photo") for x in photos})
        c=st.columns(4)
        with c[0]: kpi("Photo links",f"{len(photos):,}","📷","red")
        with c[1]: kpi("Sites with photos",f"{len(sites):,}","●","blue")
        with c[2]: kpi("Photo types",f"{len(types):,}","▤","purple")
        with c[3]: kpi("Source", "MoDa", "↗", "green")
        if not photos:
            st.info("No photo evidence is available under the current filters.")
        else:
            chosen=st.selectbox("Site",["All sites"]+sites,key="photo_site")
            view=[x for x in photos if chosen=="All sites" or str(x.get("site") or "")==chosen]
            gallery=view[:18]
            cols=st.columns(3)
            for i,pic in enumerate(gallery):
                with cols[i%3]:
                    url=html.escape(str(pic.get("url") or ""),quote=True)
                    meta=f"{pic.get('site','')} · {pic.get('photo_type','Monitoring photo')} · {pic.get('date','')}"
                    st.markdown(f'''<div class="panel"><a href="{url}" target="_blank" rel="noopener"><img src="{url}" style="width:100%;height:185px;object-fit:cover;border-radius:4px;background:#EDF2F5" onerror="this.style.display='none';this.nextElementSibling.style.display='block'"></a><div style="display:none;padding:40px 8px;text-align:center;color:#748A9E">Image requires MoDa authentication. Open the source link.</div><div style="font-size:10px;color:#536E84;margin-top:7px">{html.escape(meta)}</div><a href="{url}" target="_blank" style="font-size:10px;color:#087BC2;text-decoration:none">Open full image ↗</a></div>''',unsafe_allow_html=True)
            pdf=pd.DataFrame(view)
            cols_show=[x for x in ["date","site","woreda","zone","activity","provider","photo_type","uuid","url"] if x in pdf.columns]
            st.dataframe(pdf[cols_show],use_container_width=True,height=330,hide_index=True)

# -----------------------------------------------------------------------------
# REPORTS & EXPORTS
# -----------------------------------------------------------------------------
elif nav == "Reports & Exports":
    page_heading("Reports & Exports", "Generate monthly monitoring products, export data and package the evidence trail for management follow-up.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        left, status_col = st.columns([4.5,1.35])
        with left:
            st.markdown('<div class="section-title">Export Center</div>',unsafe_allow_html=True)
            if st.button("Generate / Refresh Monthly Outputs", type="primary", use_container_width=True, key="prepare_outputs"):
                with st.spinner("Building trackers and reports..."):
                    detail_xlsx = detailed_tracker_xlsx(a["detail"])
                    agg_xlsx = aggregated_tracker_xlsx(a["agg"])
                    mgmt_docx = management_report_docx(reporting_month, a["dq"], a["indicators"], a["agg"], a["detail"], sub_office=sub_office)
                    dq_docx = dq_report_docx(reporting_month, a["dq"], a["indicators"], sub_office=sub_office)
                    std_df=rows_to_df(a["rows"]).drop(columns=["date_obj"],errors="ignore")
                    std_csv = std_df.to_csv(index=False).encode("utf-8-sig")
                    ind_csv = pd.DataFrame(a["indicators"]).to_csv(index=False).encode("utf-8-sig")
                    pts=build_site_points(a["rows"],a.get("detail",[]),a.get("photos",[]))
                    map_csv=pts.to_csv(index=False).encode("utf-8-sig")
                    so_slug = sub_office.replace(" ", "_")
                    bundle = bundle_zip({
                        f"{reporting_month}_{so_slug}_Detailed_Findings_Tracker.xlsx": detail_xlsx,
                        f"{reporting_month}_{so_slug}_Aggregated_Action_Tracker.xlsx": agg_xlsx,
                        f"{reporting_month}_{so_slug}_Management_Monitoring_Report.docx": mgmt_docx,
                        f"{reporting_month}_{so_slug}_Data_Quality_Report.docx": dq_docx,
                        f"{reporting_month}_{so_slug}_Mapped_Site_Register.csv": map_csv,
                        f"{reporting_month}_{so_slug}_Standardized_Submissions.csv": std_csv,
                        f"{reporting_month}_{so_slug}_Indicator_Summary.csv": ind_csv,
                        f"{reporting_month}_{so_slug}_DQ_Summary.json": json.dumps(a["dq"], default=str, indent=2).encode("utf-8"),
                    })
                    st.session_state.generated_outputs={"month":reporting_month,"sub_office":sub_office,"detail":detail_xlsx,"agg":agg_xlsx,"mgmt":mgmt_docx,"dq":dq_docx,"map":map_csv,"std":std_csv,"ind":ind_csv,"bundle":bundle}
            out=st.session_state.generated_outputs
            ready=bool(out and out.get("month")==reporting_month and out.get("sub_office")==sub_office)
            cards=st.columns(4)
            names=[
                ("Detailed Findings Tracker","Site-level findings with evidence and validation fields","detail","xlsx"),
                ("Aggregated Action Tracker","Consolidated issues and follow-up status","agg","xlsx"),
                ("Management Monitoring Report","Monthly narrative report with key findings","mgmt","docx"),
                ("Data Quality Report","Completeness, consistency and DQ exceptions","dq","docx"),
            ]
            for i,(title,desc,key,ext) in enumerate(names):
                with cards[i]:
                    st.markdown(f'<div class="panel" style="min-height:135px"><div class="panel-title">{html.escape(title)}</div><div style="font-size:10px;color:#6D8295;min-height:50px">{html.escape(desc)}</div><div style="font-size:10px;color:#1A8B55;margin-top:8px">.{ext}</div></div>',unsafe_allow_html=True)
                    if ready:
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if ext=="xlsx" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        st.download_button("Download",out[key],f"{reporting_month}_{sub_office.replace(' ','_')}_{title.replace(' ','_')}.{ext}",mime,use_container_width=True,key=f"dl_{key}")
            cards2=st.columns(3)
            with cards2[0]:
                st.markdown('<div class="panel" style="min-height:120px"><div class="panel-title">Mapped Site Register</div><div style="font-size:10px;color:#6D8295">Mapped sites, coordinates, monitoring events, findings and photo counts.</div></div>',unsafe_allow_html=True)
                if ready: st.download_button("Download",out["map"],f"{reporting_month}_{sub_office.replace(' ','_')}_Mapped_Sites.csv","text/csv",use_container_width=True,key="dl_map")
            with cards2[1]:
                st.markdown('<div class="panel" style="min-height:120px"><div class="panel-title">Standardized Dataset</div><div style="font-size:10px;color:#6D8295">Clean monthly monitoring cohort for reproducible downstream analysis.</div></div>',unsafe_allow_html=True)
                if ready: st.download_button("Download",out["std"],f"{reporting_month}_{sub_office.replace(' ','_')}_Standardized_Submissions.csv","text/csv",use_container_width=True,key="dl_std")
            with cards2[2]:
                st.markdown('<div class="panel" style="min-height:120px"><div class="panel-title">Full Monthly Package</div><div style="font-size:10px;color:#6D8295">Complete set of trackers, reports and machine-readable outputs.</div></div>',unsafe_allow_html=True)
                if ready: st.download_button("Download ZIP",out["bundle"],f"{reporting_month}_{sub_office.replace(' ','_')}_JMAAP_Monthly_Package.zip","application/zip",use_container_width=True,key="dl_bundle")
        with status_col:
            dq_n=len(a.get("dq",{}).get("issues",[])) if isinstance(a.get("dq"),dict) else 0
            detail=pd.DataFrame(a.get("detail",[]))
            pending=int(detail["Validation_Status"].astype(str).str.contains("Pending",case=False,na=False).sum()) if not detail.empty and "Validation_Status" in detail.columns else 0
            rows_html=[
                ("✓","MoDa Upload Loaded",f"{len(a.get('rows',[])):,} accepted monthly submissions","#1DA65A"),
                ("✓","DQ Review Available",f"{dq_n} automated DQ exception type(s)","#1DA65A"),
                ("✓","Analysis Ready",f"{len(a.get('detail',[])):,} detailed findings generated","#1DA65A"),
                ("○","Validation Pending",f"{pending:,} finding records pending field validation","#9AAABC"),
            ]
            body=''.join(f'<div class="rail-item"><div style="width:25px;height:25px;border-radius:50%;background:{c};color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800">{ic}</div><div class="rail-main"><b>{html.escape(t)}</b><div class="rail-meta">{html.escape(d)}</div></div></div>' for ic,t,d,c in rows_html)
            st.markdown(f'<div class="rail-card"><div class="rail-head">Data Readiness Status</div>{body}</div>',unsafe_allow_html=True)
            st.markdown('<div class="status-ok">Ready to generate reports<br><span style="font-size:10px">Core analysis requirements are complete.</span></div>',unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# ABOUT
# -----------------------------------------------------------------------------
elif nav == "About":
    page_heading("About J-MAAP", "Jijiga Area Office Monitoring Automation and Analysis Platform.")
    st.markdown('''<div class="panel"><div class="panel-title">Purpose</div><div style="font-size:12px;color:#526E85;line-height:1.6">J-MAAP turns raw monthly MoDa monitoring data into transparent data-quality diagnostics, monitoring coverage views, thematic findings, geographic evidence, photo evidence and management action tracking. Automated findings remain provisional until field validation and internal agreement.</div></div>''',unsafe_allow_html=True)
