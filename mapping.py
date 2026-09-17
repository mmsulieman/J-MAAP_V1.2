from __future__ import annotations

import csv
import html
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from typing import Iterable

import pandas as pd

SOMALI_BOUNDS = (3.35, 38.70, 11.30, 48.05)  # south, west, north, east

ACTIVITY_LABELS = {
    "Activity 1 (Relief response)": "Relief",
    "Activity 2 (Nutrition assistance)": "Nutrition",
    "Activity 3 (Refugee operations)": "Refugee",
    "Activity 6 (resilience)": "Resilience",
}

ACTIVITY_COLORS = {
    "Relief": "#007DBC",
    "Nutrition": "#2E9F49",
    "Refugee": "#F28E2B",
    "Resilience": "#7A4DB3",
    "Other": "#6B7F92",
}

PROVIDER_COLORS = {
    "WFP": "#007DBC",
    "TPM": "#E6A700",
    "Other": "#6B7F92",
}

SEVERITY_COLORS = {
    "Critical": "#8B0000",
    "High": "#D64545",
    "Medium-High": "#F28E2B",
    "Medium": "#F2C94C",
    "Low": "#2E9F49",
    "No flagged finding": "#2E9F49",
}
SEVERITY_RANK = {"Critical": 5, "High": 4, "Medium-High": 3, "Medium": 2, "Low": 1, "No flagged finding": 0}
PHOTO_WORDS = re.compile(r"\b(photo|picture|image|photograph)\b", re.I)
HTTP_URL = re.compile(r"^https?://", re.I)


def _provider_label(v: str) -> str:
    s = str(v or "").strip()
    low = s.casefold()
    if "third party" in low or "wfma" in low or low == "tpm":
        return "TPM"
    if "wfp" in low:
        return "WFP"
    return s or "Other"


def _float(v):
    try:
        if v in (None, ""):
            return None
        return float(v)
    except Exception:
        return None


def _coords_from_row(r: dict):
    lat = _float(r.get("latitude"))
    lon = _float(r.get("longitude"))
    if lat is None or lon is None:
        raw = str(r.get("geographic_coordinates") or "").strip()
        if raw:
            parts = raw.replace(",", " ").split()
            if len(parts) >= 2:
                lat = _float(parts[0]); lon = _float(parts[1])
    if lat is None or lon is None:
        return None, None
    south, west, north, east = SOMALI_BOUNDS
    if not (south <= lat <= north and west <= lon <= east):
        return None, None
    return lat, lon


def _short_activity(v: str) -> str:
    return ACTIVITY_LABELS.get(str(v or "").strip(), str(v or "").strip() or "Other")


def normalize_image_url(url: str) -> str:
    """Normalize MoDa attachment URLs for HTTPS-hosted Streamlit deployments."""
    u = str(url or "").strip()
    if u.startswith("http://api.moda.wfp.org/"):
        u = "https://api.moda.wfp.org/" + u[len("http://api.moda.wfp.org/"):]
    return u


def _photo_type(question: str) -> str:
    q = str(question or "").casefold()
    if "distribution site" in q:
        return "Distribution site"
    if "storeroom" in q or "storage room" in q:
        return "Storeroom"
    if "operational" in q:
        return "Operational picture"
    if "barcode" in q or "logo" in q or "this food" in q:
        return "Commodity / diversion evidence"
    return "Monitoring photo"


def _photo_candidate_header(header: str) -> bool:
    h = str(header or "").strip()
    if not PHOTO_WORDS.search(h):
        return False
    low = h.casefold()
    # Exclude questions that ask about a photo concept rather than request an attachment.
    if "has a photo" in low or "how many picture" in low or "how many photo" in low:
        return False
    return True


def _base_meta_from_raw(raw: dict, core) -> dict:
    def first(aliases):
        for a in aliases:
            v = raw.get(a)
            if v not in (None, ""):
                return str(v).strip()
        return ""

    d = core.excel_serial_to_date(first(core.BASE_FIELDS["monitoring_date"]))
    site = first(core.BASE_FIELDS.get("site_alt", [])) or first(core.BASE_FIELDS.get("site", []))
    date_iso = d.isoformat() if d else ""
    activity = _short_activity(first(core.BASE_FIELDS["activity"]))
    provider = _provider_label(first(core.BASE_FIELDS["provider"]))
    return {
        "uuid": first(core.BASE_FIELDS["uuid"]),
        "date": date_iso,
        "date_obj": d,
        "sub_office": first(core.BASE_FIELDS["sub_office"]),
        "zone": first(core.BASE_FIELDS["zone"]),
        "woreda": first(core.BASE_FIELDS["woreda"]),
        "site": site,
        "activity": activity,
        "provider": provider,
        "event_key": "|".join([date_iso, site, activity, provider]),
        "parent_index": str(raw.get("_index") or "").replace(".0", ""),
    }


def _stream_main_photo_records(path: str, reporting_month: str, sub_office: str):
    """Stream attachment URLs from the main MoDa `data` sheet.

    Duplicate questionnaire labels are preserved because extraction is keyed by Excel
    column rather than question text. Returns photo rows and selected parent metadata
    keyed by the MoDa `_index` field for repeat-sheet linkage.
    """
    import monitoring_automation_v1 as core

    start, end = core.month_bounds(reporting_month)
    target_sub = str(sub_office or "").strip().casefold()
    photos = []
    parents = {}

    with zipfile.ZipFile(path) as z:
        ss = core._shared_strings(z)
        sp = core._sheet_path(z, "data")
        header_by_col = {}
        meta_cols = {}
        photo_cols = {}
        needed_meta = set()
        for aliases in [
            core.BASE_FIELDS["uuid"], core.BASE_FIELDS["monitoring_date"], core.BASE_FIELDS["sub_office"],
            core.BASE_FIELDS["zone"], core.BASE_FIELDS["woreda"], core.BASE_FIELDS["site"],
            core.BASE_FIELDS.get("site_alt", []), core.BASE_FIELDS["activity"], core.BASE_FIELDS["provider"],
        ]:
            needed_meta.update(aliases)
        needed_meta.add("_index")

        for _, e in ET.iterparse(z.open(sp), events=("end",)):
            if e.tag != core.MAIN_NS + "row":
                continue
            rn = int(e.attrib.get("r", "0"))
            if rn == 1:
                for c in e.findall(core.MAIN_NS + "c"):
                    col = core.col_letters(c.attrib.get("r", "A1"))
                    val = core._cell_value(c, ss)
                    header_by_col[col] = val
                    if val in needed_meta:
                        meta_cols[col] = val
                    if _photo_candidate_header(val):
                        photo_cols[col] = val
                e.clear()
                continue

            raw = {}
            photo_vals = []
            for c in e.findall(core.MAIN_NS + "c"):
                col = core.col_letters(c.attrib.get("r", "A1"))
                if col not in meta_cols and col not in photo_cols:
                    continue
                val = core._cell_value(c, ss)
                if col in meta_cols:
                    key = meta_cols[col]
                    # `_index` and aliases are unique enough for metadata; preserve first nonblank.
                    if key not in raw or (not raw.get(key) and val not in (None, "")):
                        raw[key] = val
                if col in photo_cols and HTTP_URL.match(str(val or "").strip()):
                    photo_vals.append((photo_cols[col], str(val).strip()))
            meta = _base_meta_from_raw(raw, core)
            d = meta.get("date_obj")
            if d and start <= d <= end and meta.get("sub_office", "").casefold() == target_sub:
                if meta.get("parent_index"):
                    parents[meta["parent_index"]] = {k: v for k, v in meta.items() if k != "date_obj"}
                for question, raw_url in photo_vals:
                    photos.append({
                        **{k: v for k, v in meta.items() if k != "date_obj"},
                        "url": normalize_image_url(raw_url),
                        "raw_url": raw_url,
                        "photo_type": _photo_type(question),
                        "question": question,
                        "source_sheet": "data",
                        "source_file": os.path.basename(path),
                    })
            e.clear()
    return photos, parents


def _stream_repeat_photos(path: str, sheet_name: str, parents: dict):
    import monitoring_automation_v1 as core
    if not parents:
        return []
    out = []
    try:
        with zipfile.ZipFile(path) as z:
            ss = core._shared_strings(z)
            sp = core._sheet_path(z, sheet_name)
            parent_col = None
            photo_cols = {}
            for _, e in ET.iterparse(z.open(sp), events=("end",)):
                if e.tag != core.MAIN_NS + "row":
                    continue
                rn = int(e.attrib.get("r", "0"))
                if rn == 1:
                    for c in e.findall(core.MAIN_NS + "c"):
                        col = core.col_letters(c.attrib.get("r", "A1"))
                        val = core._cell_value(c, ss)
                        if val == "_parent_index":
                            parent_col = col
                        if _photo_candidate_header(val):
                            photo_cols[col] = val
                    e.clear(); continue
                parent_index = ""
                pvals = []
                for c in e.findall(core.MAIN_NS + "c"):
                    col = core.col_letters(c.attrib.get("r", "A1"))
                    if col != parent_col and col not in photo_cols:
                        continue
                    val = core._cell_value(c, ss)
                    if col == parent_col:
                        parent_index = str(val or "").replace(".0", "")
                    elif col in photo_cols and HTTP_URL.match(str(val or "").strip()):
                        pvals.append((photo_cols[col], str(val).strip()))
                parent = parents.get(parent_index)
                if parent:
                    for question, raw_url in pvals:
                        out.append({
                            **parent,
                            "url": normalize_image_url(raw_url),
                            "raw_url": raw_url,
                            "photo_type": _photo_type(question),
                            "question": question,
                            "source_sheet": sheet_name,
                            "source_file": os.path.basename(path),
                        })
                e.clear()
    except (KeyError, ValueError, FileNotFoundError):
        return []
    return out


def extract_photo_evidence(paths: Iterable[str], reporting_month: str, sub_office: str) -> list[dict]:
    """Extract and deduplicate MoDa attachment links for the selected cohort.

    Supports attachment columns on the main `data` sheet and the operational-picture
    repeat sheet. Photo evidence remains link-based; J-MAAP never copies or republishes
    MoDa images, and normal MoDa access controls continue to apply.
    """
    import monitoring_automation_v1 as core

    photos = []
    for p in paths or []:
        p = str(p)
        if p.lower().endswith((".xlsx", ".xlsm")):
            main, parents = _stream_main_photo_records(p, reporting_month, sub_office)
            photos.extend(main)
            photos.extend(_stream_repeat_photos(p, "Technical_module1_picture_repea", parents))
        elif p.lower().endswith(".csv"):
            start, end = core.month_bounds(reporting_month)
            target_sub = str(sub_office or "").strip().casefold()
            with open(p, encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                photo_headers = [h for h in (reader.fieldnames or []) if _photo_candidate_header(h)]
                for raw in reader:
                    meta = _base_meta_from_raw(raw, core)
                    d = meta.get("date_obj")
                    if not (d and start <= d <= end and meta.get("sub_office", "").casefold() == target_sub):
                        continue
                    for h in photo_headers:
                        raw_url = str(raw.get(h) or "").strip()
                        if HTTP_URL.match(raw_url):
                            photos.append({
                                **{k: v for k, v in meta.items() if k != "date_obj"},
                                "url": normalize_image_url(raw_url), "raw_url": raw_url,
                                "photo_type": _photo_type(h), "question": h,
                                "source_sheet": "data", "source_file": os.path.basename(p),
                            })

    # Re-exported MoDa files can contain the same attachment. De-duplicate by UUID+URL;
    # where UUID is unavailable, URL alone is sufficiently specific for an attachment.
    dedup = []
    seen = set()
    for x in photos:
        key = (x.get("uuid") or "", x.get("url") or "")
        if key in seen:
            continue
        seen.add(key)
        dedup.append(x)
    return sorted(dedup, key=lambda x: (x.get("date", ""), x.get("site", ""), x.get("photo_type", "")), reverse=True)


def apply_photo_filters(photos: Iterable[dict], activities=None, zones=None, woredas=None, providers=None):
    activities = set(activities or [])
    zones = set(zones or [])
    woredas = set(woredas or [])
    providers = set(providers or [])
    out = []
    for r in photos or []:
        if activities and _short_activity(r.get("activity")) not in activities: continue
        if zones and str(r.get("zone") or "").strip() not in zones: continue
        if woredas and str(r.get("woreda") or "").strip() not in woredas: continue
        if providers and _provider_label(r.get("provider")) not in {_provider_label(x) for x in providers}: continue
        out.append(r)
    return out


def build_site_points(rows: Iterable[dict], detailed_findings: Iterable[dict] | None = None, photos: Iterable[dict] | None = None) -> pd.DataFrame:
    """Aggregate valid GPS submissions into one point per site/woreda/zone.

    Coordinates use the median of valid MoDa GPS observations to dampen handheld-GPS jitter.
    Monitoring events use date × site × activity × provider as the operational event grain.
    Photo evidence is linked by site/woreda/zone and retained as traceable submission URLs.
    """
    recs = []
    for r in rows:
        lat, lon = _coords_from_row(r)
        if lat is None:
            continue
        site = str(r.get("site_reported") or "").strip()
        woreda = str(r.get("woreda") or "").strip()
        zone = str(r.get("zone") or "").strip()
        activity = _short_activity(r.get("activity"))
        provider = _provider_label(r.get("provider"))
        date = str(r.get("monitoring_date_iso") or "").strip()
        if not site:
            site = f"GPS site {lat:.4f}, {lon:.4f}"
        recs.append({
            "site": site, "woreda": woreda, "zone": zone,
            "activity": activity, "provider": provider,
            "date": date, "lat": lat, "lon": lon,
            "uuid": str(r.get("uuid") or ""),
            "event_key": "|".join([date, site, activity, provider]),
        })
    cols = ["site","woreda","zone","lat","lon","submissions","events","activities","providers","latest_date","finding_signals","issue_observations","severity","photo_count","photo_types","latest_photo_url","popup_photos"]
    if not recs:
        return pd.DataFrame(columns=cols)
    d = pd.DataFrame(recs)

    find_by_site = {}
    if detailed_findings:
        fdf = pd.DataFrame(list(detailed_findings))
        if not fdf.empty and "Site" in fdf.columns:
            group_cols = [c for c in ["Site", "Woreda", "Zone"] if c in fdf.columns]
            for keys, g in fdf.groupby(group_cols, dropna=False):
                if not isinstance(keys, tuple): keys = (keys,)
                keymap = dict(zip(group_cols, keys))
                fkey = (str(keymap.get("Site", "")), str(keymap.get("Woreda", "")), str(keymap.get("Zone", "")))
                sevs = [str(x) for x in g.get("Severity", pd.Series(dtype=str)).dropna().tolist()]
                sev = max(sevs, key=lambda x: SEVERITY_RANK.get(x, -1), default="No flagged finding")
                find_by_site[fkey] = {
                    "finding_signals": int(g["Finding_Code"].nunique()) if "Finding_Code" in g else int(len(g)),
                    "issue_observations": int(pd.to_numeric(g.get("Issue_Count", 0), errors="coerce").fillna(0).sum()) if "Issue_Count" in g else 0,
                    "severity": sev,
                }

    photo_by_site = defaultdict(list)
    for p in photos or []:
        key = (str(p.get("site") or ""), str(p.get("woreda") or ""), str(p.get("zone") or ""))
        if key[0] and p.get("url"):
            photo_by_site[key].append(dict(p))
    for key, vals in list(photo_by_site.items()):
        unique = []
        seen_urls = set()
        for p in sorted(vals, key=lambda x: x.get("date", ""), reverse=True):
            if p.get("url") in seen_urls: continue
            seen_urls.add(p.get("url")); unique.append(p)
        photo_by_site[key] = unique

    out = []
    for (site, woreda, zone), g in d.groupby(["site","woreda","zone"], dropna=False):
        primary_activity = g["activity"].value_counts().index[0] if len(g) else "Other"
        primary_provider = g["provider"].value_counts().index[0] if len(g) else "Other"
        fd = find_by_site.get((str(site), str(woreda), str(zone)), {"finding_signals":0, "issue_observations":0, "severity":"No flagged finding"})
        ph = photo_by_site.get((str(site), str(woreda), str(zone)), [])
        out.append({
            "site": str(site), "woreda": str(woreda), "zone": str(zone),
            "lat": float(g["lat"].median()), "lon": float(g["lon"].median()),
            "submissions": int(len(g)), "events": int(g["event_key"].nunique()),
            "activities": ", ".join(sorted(set(g["activity"].astype(str)))),
            "providers": ", ".join(sorted(set(g["provider"].astype(str)))),
            "primary_activity": primary_activity, "primary_provider": primary_provider,
            "latest_date": max([x for x in g["date"].astype(str).tolist() if x], default=""),
            "photo_count": len(ph),
            "photo_types": ", ".join(sorted({str(x.get("photo_type") or "Monitoring photo") for x in ph})),
            "latest_photo_url": ph[0].get("url", "") if ph else "",
            "popup_photos": [{"url": x.get("url", ""), "type": x.get("photo_type", "Monitoring photo"), "date": x.get("date", "")} for x in ph[:4]],
            **fd,
        })
    return pd.DataFrame(out).sort_values(["zone","woreda","site"]).reset_index(drop=True)


def apply_map_filters(rows: Iterable[dict], activities=None, zones=None, woredas=None, providers=None):
    activities = set(activities or [])
    zones = set(zones or [])
    woredas = set(woredas or [])
    providers = set(providers or [])
    provider_norm = {_provider_label(x) for x in providers}
    out = []
    for r in rows:
        if activities and _short_activity(r.get("activity")) not in activities: continue
        if zones and str(r.get("zone") or "").strip() not in zones: continue
        if woredas and str(r.get("woreda") or "").strip() not in woredas: continue
        if providers and _provider_label(r.get("provider")) not in provider_norm: continue
        out.append(r)
    return out


def load_admin_assets(app_dir: str | Path) -> dict:
    p = Path(app_dir) / "assets" / "maps"
    return {
        "zones": json.loads((p / "somali_zones_2024.geojson").read_text(encoding="utf-8")),
        "woredas": json.loads((p / "somali_woredas_2024.geojson").read_text(encoding="utf-8")),
        "zone_labels": json.loads((p / "zone_labels.json").read_text(encoding="utf-8")),
        "woreda_labels": json.loads((p / "woreda_labels.json").read_text(encoding="utf-8")),
        "bounds": json.loads((p / "somali_bounds.json").read_text(encoding="utf-8")),
    }


def _marker_color(row: dict, mode: str) -> str:
    if mode == "Provider":
        return PROVIDER_COLORS.get(row.get("primary_provider"), PROVIDER_COLORS["Other"])
    if mode == "Finding severity":
        return SEVERITY_COLORS.get(row.get("severity"), "#6B7F92")
    return ACTIVITY_COLORS.get(row.get("primary_activity"), ACTIVITY_COLORS["Other"])


def _popup_html(row: dict, show_photos: bool) -> str:
    esc = lambda x: html.escape(str(x or ""), quote=True)
    photo_block = ""
    photos = row.get("popup_photos") or []
    if show_photos and photos:
        cards = []
        for p in photos:
            url = normalize_image_url(p.get("url", ""))
            if not HTTP_URL.match(url):
                continue
            qurl = html.escape(url, quote=True)
            ptype = esc(p.get("type", "Monitoring photo"))
            pdate = esc(p.get("date", ""))
            cards.append(
                f"<div class='photo-card'><a href='{qurl}' target='_blank' rel='noopener noreferrer'>"
                f"<img src='{qurl}' loading='lazy' referrerpolicy='no-referrer' onerror=\"this.style.display='none';this.nextElementSibling.style.display='block';\">"
                f"<span class='photo-fallback'>Open image</span></a><div class='photo-meta'>{ptype}{' · ' + pdate if pdate else ''}</div></div>"
            )
        if cards:
            photo_block = f"<div class='photo-title'>📷 Monitoring photo evidence ({int(row.get('photo_count') or 0)})</div><div class='photo-grid'>{''.join(cards)}</div>"
            if int(row.get("photo_count") or 0) > len(cards):
                photo_block += "<div class='photo-note'>More photos are available in the gallery below the map.</div>"
    elif int(row.get("photo_count") or 0) > 0:
        photo_block = f"<div class='photo-title'>📷 {int(row.get('photo_count') or 0)} photo(s) available in the gallery below the map.</div>"

    return (
        f"<div class='popup-wrap'><div class='popup-site'>{esc(row.get('site'))}</div>"
        f"<b>Zone / Woreda:</b> {esc(row.get('zone'))} / {esc(row.get('woreda'))}<br>"
        f"<b>Activities:</b> {esc(row.get('activities'))}<br>"
        f"<b>Providers:</b> {esc(row.get('providers'))}<br>"
        f"<b>Monitoring events:</b> {int(row.get('events') or 0)}<br>"
        f"<b>Submissions:</b> {int(row.get('submissions') or 0)}<br>"
        f"<b>Latest monitoring:</b> {esc(row.get('latest_date'))}<br>"
        f"<b>Finding signals:</b> {int(row.get('finding_signals') or 0)}<br>"
        f"<b>Highest severity:</b> {esc(row.get('severity') or 'No flagged finding')}"
        f"{photo_block}</div>"
    )


def leaflet_html(points: pd.DataFrame, assets: dict, initial_basemap: str = "Street", display_mode: str = "Activity", show_woreda_labels: bool = False, show_photos: bool = True) -> str:
    """Return a Leaflet map body for Streamlit, including optional MoDa photo thumbnails."""
    pts = []
    for r in points.to_dict("records"):
        r = dict(r)
        r["color"] = _marker_color(r, display_mode)
        r["radius"] = max(6, min(15, 5 + (r.get("events", 1) ** 0.5) * 2.2)) if display_mode == "Visit frequency" else 7
        r["popup_html"] = _popup_html(r, show_photos=show_photos)
        # Keep the browser payload lean: popup HTML contains the thumbnails already.
        r.pop("popup_photos", None)
        pts.append(r)
    payload = json.dumps(pts, ensure_ascii=False).replace("</", "<\\/")
    zones = json.dumps(assets["zones"], ensure_ascii=False).replace("</", "<\\/")
    woredas = json.dumps(assets["woredas"], ensure_ascii=False).replace("</", "<\\/")
    zl = json.dumps(assets["zone_labels"], ensure_ascii=False).replace("</", "<\\/")
    wl = json.dumps(assets["woreda_labels"], ensure_ascii=False).replace("</", "<\\/")
    b = assets["bounds"]
    initial = initial_basemap
    show_w = "true" if show_woreda_labels else "false"

    legend_items = ""
    palette = ACTIVITY_COLORS if display_mode in ("Activity", "Visit frequency") else PROVIDER_COLORS if display_mode == "Provider" else SEVERITY_COLORS
    for k, v in palette.items():
        legend_items += f'<div><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{v};margin-right:6px"></span>{html.escape(k)}</div>'

    return f"""<!doctype html>
<html><head><meta charset='utf-8'>
<link rel='stylesheet' href='https://unpkg.com/leaflet@1.9.4/dist/leaflet.css'/>
<script src='https://unpkg.com/leaflet@1.9.4/dist/leaflet.js'></script>
<style>
html,body,#map{{height:100%;margin:0;font-family:Arial,sans-serif}} #map{{background:#fff}}
.zone-label{{font-weight:700;color:#143D63;background:rgba(255,255,255,.82);border:0;box-shadow:none;font-size:12px;padding:1px 3px}}
.woreda-label{{font-weight:500;color:#333;background:rgba(255,255,255,.72);border:0;box-shadow:none;font-size:9px;padding:0 2px}}
.legend{{background:rgba(255,255,255,.94);padding:9px 11px;border-radius:6px;line-height:1.55;box-shadow:0 1px 5px rgba(0,0,0,.25);font-size:12px}}
.leaflet-popup-content{{min-width:280px;max-width:410px;margin:12px 14px}}
.popup-site{{font-weight:700;font-size:15px;color:#123D67;margin-bottom:5px}}
.photo-title{{font-weight:700;color:#123D67;margin:10px 0 6px;border-top:1px solid #E1E8EE;padding-top:7px}}
.photo-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:6px}}
.photo-card{{border:1px solid #DCE6EF;border-radius:7px;padding:4px;background:#F8FBFD}}
.photo-card img{{display:block;width:100%;height:95px;object-fit:cover;border-radius:5px;background:#EEF3F6}}
.photo-fallback{{display:none;padding:25px 8px;text-align:center;color:#006AA6;font-weight:700}}
.photo-meta{{font-size:10px;color:#60758A;margin-top:3px;line-height:1.2}}
.photo-note{{font-size:10px;color:#60758A;margin-top:5px}}
</style></head><body><div id='map'></div><script>
const points={payload}; const zones={zones}; const woredas={woredas}; const zoneLabels={zl}; const woredaLabels={wl};
const map=L.map('map',{{zoomControl:true, attributionControl:true}});
const street=L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png',{{maxZoom:19,attribution:'&copy; OpenStreetMap contributors'}});
const satellite=L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{{z}}/{{y}}/{{x}}',{{maxZoom:19,attribution:'Tiles &copy; Esri and contributors'}});
const blank=L.layerGroup();
const zoneLayer=L.geoJSON(zones,{{style:()=>({{color:'#16324F',weight:2.7,fill:false,opacity:.95}}),onEachFeature:(f,l)=>l.bindTooltip(f.properties.Zone_Name||'',{{sticky:true}})}}).addTo(map);
const woredaLayer=L.geoJSON(woredas,{{style:()=>({{color:'#A9B5C0',weight:.8,fill:false,opacity:.85}}),onEachFeature:(f,l)=>l.bindTooltip((f.properties.DISTRICTS||'')+' · '+(f.properties.Zone_Name||''),{{sticky:true}})}}).addTo(map);
const zoneText=L.layerGroup(zoneLabels.map(x=>L.marker([x.lat,x.lon],{{icon:L.divIcon({{className:'',html:`<div class="zone-label">${{x.name}}</div>`,iconSize:null}}),interactive:false}}))).addTo(map);
const woredaText=L.layerGroup(woredaLabels.map(x=>L.marker([x.lat,x.lon],{{icon:L.divIcon({{className:'',html:`<div class="woreda-label">${{x.name}}</div>`,iconSize:null}}),interactive:false}})));
if({show_w}) woredaText.addTo(map);
const siteLayer=L.layerGroup().addTo(map);
points.forEach(p=>{{
 const marker=L.circleMarker([p.lat,p.lon],{{radius:p.radius,color:'#ffffff',weight:1.3,fillColor:p.color,fillOpacity:.9}});
 marker.bindPopup(p.popup_html,{{maxWidth:430}}); marker.bindTooltip(p.site + ((p.photo_count||0)>0 ? ' 📷' : ''),{{sticky:true}}); marker.addTo(siteLayer);
}});
const baseMaps={{'Street':street,'Satellite imagery':satellite,'Administrative hollow':blank}};
const overlays={{'Zone boundaries':zoneLayer,'Woreda boundaries':woredaLayer,'Visited sites':siteLayer,'Zone labels':zoneText,'Woreda labels':woredaText}};
L.control.layers(baseMaps,overlays,{{collapsed:false}}).addTo(map);
if('{initial}'==='Satellite imagery') satellite.addTo(map); else if('{initial}'==='Administrative hollow') blank.addTo(map); else street.addTo(map);
if(points.length){{ const bb=L.latLngBounds(points.map(p=>[p.lat,p.lon])); map.fitBounds(bb.pad(.12),{{maxZoom:10}}); }}
else map.fitBounds([[{b['south']},{b['west']}],[{b['north']},{b['east']}]]);
const legend=L.control({{position:'bottomright'}}); legend.onAdd=()=>{{const d=L.DomUtil.create('div','legend');d.innerHTML='<b>{html.escape(display_mode)}</b>{legend_items.replace("'", "\\'")}';return d;}};legend.addTo(map);
</script></body></html>"""


def leaflet_dashboard_html(points: pd.DataFrame, assets: dict, initial_basemap: str = "Street", display_mode: str = "Activity", show_woreda_labels: bool = False) -> str:
    """Render the polished J-MAAP map + interactive site-details rail in one Leaflet component.

    Marker clicks update the right-hand evidence panel entirely in-browser, avoiding a
    roundtrip through Streamlit while keeping the map responsive on Community Cloud.
    """
    pts = []
    for r in points.to_dict("records"):
        x = dict(r)
        x["color"] = _marker_color(x, display_mode)
        x["radius"] = max(6, min(14, 5 + (x.get("events", 1) ** 0.5) * 2.1)) if display_mode == "Visit frequency" else 6.5
        # Normalize image URLs before they reach the browser.
        photos = []
        for p in (x.get("popup_photos") or []):
            u = normalize_image_url(p.get("url", ""))
            if HTTP_URL.match(u):
                photos.append({"url": u, "type": p.get("type", "Monitoring photo"), "date": p.get("date", "")})
        x["popup_photos"] = photos
        pts.append(x)

    payload = json.dumps(pts, ensure_ascii=False, default=str).replace("</", "<\\/")
    zones = json.dumps(assets["zones"], ensure_ascii=False).replace("</", "<\\/")
    woredas = json.dumps(assets["woredas"], ensure_ascii=False).replace("</", "<\\/")
    zl = json.dumps(assets["zone_labels"], ensure_ascii=False).replace("</", "<\\/")
    wl = json.dumps(assets["woreda_labels"], ensure_ascii=False).replace("</", "<\\/")
    b = assets["bounds"]
    initial = initial_basemap
    show_w = "true" if show_woreda_labels else "false"

    legend_items = ""
    palette = ACTIVITY_COLORS if display_mode in ("Activity", "Visit frequency") else PROVIDER_COLORS if display_mode == "Provider" else SEVERITY_COLORS
    for k, v in palette.items():
        legend_items += f'<div class="legend-row"><span style="background:{v}"></span>{html.escape(k)}</div>'

    return f"""<!doctype html>
<html><head><meta charset='utf-8'>
<meta name='viewport' content='width=device-width,initial-scale=1'>
<link rel='stylesheet' href='https://unpkg.com/leaflet@1.9.4/dist/leaflet.css'/>
<script src='https://unpkg.com/leaflet@1.9.4/dist/leaflet.js'></script>
<style>
*{{box-sizing:border-box}} html,body{{height:100%;margin:0;font-family:'Open Sans','Segoe UI',Arial,sans-serif;color:#183F60;background:#fff}}
.dashboard-map{{height:100%;display:flex;flex-direction:column;border:1px solid #DDE7F0;border-radius:7px;overflow:hidden;background:#fff}}
.map-toolbar{{height:43px;display:flex;align-items:center;border-bottom:1px solid #DDE7F0;background:#fff;padding:0 8px;gap:0}}
.map-btn{{height:31px;padding:0 18px;border:1px solid #D7E3ED;background:#fff;color:#244E70;font-weight:600;font-size:12px;cursor:pointer}}
.map-btn:first-child{{border-radius:5px 0 0 5px}} .map-btn:last-of-type{{border-radius:0 5px 5px 0}}
.map-btn.active{{background:#1685D3;color:#fff;border-color:#1685D3}}
.map-body{{display:grid;grid-template-columns:minmax(0,1fr) 285px;min-height:0;flex:1}}
#map{{height:100%;min-height:580px;background:#EDF2F5}}
.site-panel{{height:100%;border-left:1px solid #DDE7F0;background:#fff;overflow-y:auto;padding:0}}
.panel-head{{padding:10px 12px;border-bottom:1px solid #DDE7F0;display:flex;justify-content:space-between;align-items:center;color:#0A4D82;font-weight:800;font-size:14px}}
.site-content{{padding:12px 13px}}
.site-name{{color:#0A4D82;font-size:18px;font-weight:800;margin-bottom:2px}}
.site-loc{{color:#788EA3;font-size:10px;margin-bottom:8px}}
.monitor-tag{{display:inline-block;background:#DFF4E5;color:#34824E;border-radius:4px;padding:3px 9px;font-size:10px;margin-bottom:10px}}
.kv{{display:grid;grid-template-columns:100px 1fr;gap:4px 7px;font-size:10px;line-height:1.65}}
.kv .k{{color:#72879A}} .kv .v{{color:#1F4E71;font-weight:600}}
.photo-title{{color:#0A4D82;font-weight:800;font-size:11px;margin:12px 0 7px;border-top:1px solid #E8EEF3;padding-top:9px}}
.hero-photo{{width:100%;height:155px;object-fit:cover;border-radius:4px;background:#EDF2F5;display:block}}
.photo-link{{display:block;text-decoration:none}}
.thumb-row{{display:flex;gap:5px;margin-top:7px;overflow-x:auto;padding-bottom:2px}}
.thumb{{width:52px;height:40px;object-fit:cover;border:1px solid #DDE7F0;border-radius:3px;cursor:pointer;background:#EDF2F5;flex:0 0 auto}}
.photo-note{{font-size:9px;color:#788EA3;margin-top:5px}}
.view-link{{display:block;color:#087BC2;font-size:10px;font-weight:700;text-decoration:none;margin-top:8px}}
.no-photo{{background:#F5F8FA;color:#7890A5;border-radius:5px;padding:28px 8px;text-align:center;font-size:10px}}
.zone-label{{font-weight:800;color:#122F49;background:rgba(255,255,255,.72);border:0;box-shadow:none;font-size:13px;padding:1px 3px;text-shadow:0 1px #fff}}
.woreda-label{{font-weight:600;color:#333;background:rgba(255,255,255,.63);border:0;box-shadow:none;font-size:8px;padding:0 2px}}
.legend{{background:rgba(255,255,255,.94);padding:9px 10px;border-radius:5px;line-height:1.45;box-shadow:0 1px 5px rgba(0,0,0,.20);font-size:10px;min-width:120px}}
.legend-title{{font-weight:800;color:#123E64;margin-bottom:5px}}
.legend-row{{display:flex;align-items:center;gap:5px;margin:2px 0}} .legend-row span{{display:inline-block;width:9px;height:9px;border-radius:50%}}
.leaflet-control-layers{{display:none}}
.camera-dot{{font-size:10px}}
@media(max-width:900px){{.map-body{{grid-template-columns:1fr}} .site-panel{{display:none}}}}
</style></head>
<body>
<div class='dashboard-map'>
 <div class='map-toolbar'>
   <button id='btnStreet' class='map-btn' onclick="switchBase('Street')">▰ &nbsp; Street Map</button>
   <button id='btnSatellite' class='map-btn' onclick="switchBase('Satellite imagery')">▣ &nbsp; Satellite / Imagery</button>
   <button id='btnHollow' class='map-btn' onclick="switchBase('Administrative hollow')">◇ &nbsp; Administrative Map (Hollow)</button>
 </div>
 <div class='map-body'>
   <div id='map'></div>
   <div class='site-panel'>
     <div class='panel-head'><span>Site Details</span><span style='color:#7890A5;font-weight:400'>×</span></div>
     <div id='siteContent' class='site-content'><div style='color:#7890A5;font-size:11px;padding:18px 5px'>Click a monitored site on the map to view its monitoring evidence.</div></div>
   </div>
 </div>
</div>
<script>
const points={payload}; const zones={zones}; const woredas={woredas}; const zoneLabels={zl}; const woredaLabels={wl};
const map=L.map('map',{{zoomControl:true,attributionControl:true}});
const street=L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png',{{maxZoom:19,attribution:'&copy; OpenStreetMap contributors'}});
const satellite=L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{{z}}/{{y}}/{{x}}',{{maxZoom:19,attribution:'Tiles &copy; Esri and contributors'}});
const blank=L.layerGroup();
let currentBase=null;
function switchBase(name){{
  [street,satellite,blank].forEach(x=>{{if(map.hasLayer(x)) map.removeLayer(x);}});
  if(name==='Satellite imagery'){{satellite.addTo(map);currentBase=satellite;}}
  else if(name==='Administrative hollow'){{blank.addTo(map);currentBase=blank;}}
  else{{street.addTo(map);currentBase=street;}}
  document.querySelectorAll('.map-btn').forEach(x=>x.classList.remove('active'));
  if(name==='Satellite imagery') document.getElementById('btnSatellite').classList.add('active');
  else if(name==='Administrative hollow') document.getElementById('btnHollow').classList.add('active');
  else document.getElementById('btnStreet').classList.add('active');
}}
const zoneLayer=L.geoJSON(zones,{{style:()=>({{color:'#172E42',weight:2.5,fill:false,opacity:.95}})}}).addTo(map);
const woredaLayer=L.geoJSON(woredas,{{style:()=>({{color:'#7E8D99',weight:.75,fill:false,opacity:.8}})}}).addTo(map);
const zoneText=L.layerGroup(zoneLabels.map(x=>L.marker([x.lat,x.lon],{{icon:L.divIcon({{className:'',html:`<div class="zone-label">${{x.name}}</div>`,iconSize:null}}),interactive:false}}))).addTo(map);
const woredaText=L.layerGroup(woredaLabels.map(x=>L.marker([x.lat,x.lon],{{icon:L.divIcon({{className:'',html:`<div class="woreda-label">${{x.name}}</div>`,iconSize:null}}),interactive:false}})));
if({show_w}) woredaText.addTo(map);
const siteLayer=L.layerGroup().addTo(map);
function esc(s){{ return String(s??'').replace(/[&<>'\"]/g,c=>({{'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','\"':'&quot;'}}[c])); }}
function showSite(p){{
  const photos=(p.popup_photos||[]).filter(x=>x.url);
  let photo='';
  if(photos.length){{
    const thumbs=photos.map((x,i)=>`<img class="thumb" src="${{esc(x.url)}}" title="${{esc(x.type||'Monitoring photo')}}" onclick="document.getElementById('heroPhoto').src=this.src">`).join('');
    photo=`<div class="photo-title">Latest Monitoring Photo</div><a class="photo-link" href="${{esc(photos[0].url)}}" target="_blank" rel="noopener"><img id="heroPhoto" class="hero-photo" src="${{esc(photos[0].url)}}" onerror="this.style.display='none';document.getElementById('photoFallback').style.display='block'"></a><div id="photoFallback" class="no-photo" style="display:none">Image requires MoDa authentication.<br>Use the image link below.</div><div class="thumb-row">${{thumbs}}</div><div class="photo-note">${{Math.min(photos.length,p.photo_count||photos.length)}} of ${{p.photo_count||photos.length}} photo(s) shown</div><a class="view-link" href="${{esc(photos[0].url)}}" target="_blank" rel="noopener">View full image ↗</a>`;
  }} else if((p.photo_count||0)>0){{ photo=`<div class="photo-title">Monitoring Photo Evidence</div><div class="no-photo">${{p.photo_count}} photo link(s) available but no embeddable thumbnail was returned.</div>`; }}
  document.getElementById('siteContent').innerHTML=`
    <div class="site-name">${{esc(p.site)}}</div>
    <div class="site-loc">${{esc(p.woreda)}} Woreda &nbsp;|&nbsp; ${{esc(p.zone)}} Zone</div>
    <div class="monitor-tag">Monitored</div>
    <div class="kv">
      <div class="k">Activity</div><div class="v">${{esc(p.activities)}}</div>
      <div class="k">Provider</div><div class="v">${{esc(p.providers)}}</div>
      <div class="k">Monitoring events</div><div class="v">${{p.events||0}}</div>
      <div class="k">Submissions</div><div class="v">${{p.submissions||0}}</div>
      <div class="k">Latest visit</div><div class="v">${{esc(p.latest_date)}}</div>
      <div class="k">Findings</div><div class="v" style="color:${{(p.finding_signals||0)>0?'#D92D43':'#20854C'}}">${{p.finding_signals||0}}${{(p.severity&&p.severity!=='No flagged finding')?' ('+esc(p.severity)+')':''}}</div>
      <div class="k">Photos</div><div class="v">${{p.photo_count||0}} 📷</div>
    </div>${{photo}}`;
}}
points.forEach((p,i)=>{{
 const marker=L.circleMarker([p.lat,p.lon],{{radius:p.radius,color:'#fff',weight:1.5,fillColor:p.color,fillOpacity:.94}});
 marker.bindTooltip(p.site + ((p.photo_count||0)>0 ? ' 📷' : ''),{{sticky:true}});
 marker.on('click',()=>showSite(p)); marker.addTo(siteLayer);
 if(i===0) setTimeout(()=>showSite(p),150);
}});
switchBase('{initial}');
if(points.length){{ const bb=L.latLngBounds(points.map(p=>[p.lat,p.lon])); map.fitBounds(bb.pad(.10),{{maxZoom:9}}); }}
else map.fitBounds([[{b['south']},{b['west']}],[{b['north']},{b['east']}] ]);
const legend=L.control({{position:'bottomleft'}}); legend.onAdd=()=>{{ const d=L.DomUtil.create('div','legend'); d.innerHTML='<div class="legend-title">Sites by {html.escape(display_mode)}</div>{legend_items.replace("'", "\\'")}<div style="margin-top:7px;border-top:1px solid #E3E9EE;padding-top:5px">○ Site without photo<br>◉ Site with photo</div>'; return d; }}; legend.addTo(map);
</script></body></html>"""
