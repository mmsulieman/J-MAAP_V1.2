from __future__ import annotations

from collections import Counter
from functools import lru_cache
from pathlib import Path
import itertools
import pandas as pd
import sys

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
import monitoring_automation_v1 as core

from io_utils import classify_source, detect_site_column, normalize_site, list_xlsx_sheets, preview_file, read_table, detect_date_column, xlsx_header_names


def profile_file(path: str, preview_rows: int = 500) -> dict:
    sheets = list_xlsx_sheets(path) if Path(path).suffix.lower() in {".xlsx", ".xlsm", ".xls"} else []
    selected_sheet = "data" if "data" in sheets else (sheets[0] if sheets else None)
    # Header-only path prevents giant MoDa forms from being loaded merely for source classification.
    if Path(path).suffix.lower() in {".xlsx", ".xlsm"} and selected_sheet:
        columns = xlsx_header_names(path, selected_sheet)
        source_type = classify_source(Path(path).name, columns, sheets)
    else:
        head = preview_file(path, selected_sheet, nrows=5)
        columns = list(head.columns)
        source_type = classify_source(Path(path).name, columns, sheets)
    site_col = detect_site_column(columns)
    date_col = detect_date_column(columns)
    uuid_unique = None; xform = []; row_count = None
    if source_type != "MoDa":
        try:
            df = read_table(path, selected_sheet)
            row_count = len(df)
            if "_uuid" in df.columns:
                uuid_unique = int(df["_uuid"].replace("", pd.NA).nunique())
            if "_xform_id" in df.columns:
                xform = sorted([x for x in df["_xform_id"].astype(str).unique().tolist() if x])
        except Exception:
            row_count = None
    # MoDa raw row/UUID counts are intentionally deferred to the matching/analysis pass so
    # the same very wide workbook is not scanned twice during initial upload.
    return {
        "path": path, "file": Path(path).name, "source_type": source_type,
        "sheet": selected_sheet or "CSV", "sheet_count": len(sheets) if sheets else 1,
        "preview_rows": row_count, "columns": len(columns), "column_names": columns,
        "site_column": site_col, "date_column": date_col,
        "uuid_unique_preview": uuid_unique, "xform_ids_preview": xform,
    }

@lru_cache(maxsize=64)
def _full_key_sets(path: str, source_type: str) -> dict:
    # Avoid loading the thousands-column MoDa workbook into pandas simply to reconcile keys.
    if source_type == "MoDa" and Path(path).suffix.lower() in {".xlsx", ".xlsm"}:
        wanted = {"_uuid", "FDP name", "Specify FDP"}
        try:
            raw, _ = core.read_moda_xlsx(path, wanted)
        except Exception:
            raw = []
        uuids = {str(r.get("_uuid", "")).strip() for r in raw if str(r.get("_uuid", "")).strip()}
        sites = set()
        for r in raw:
            v = r.get("Specify FDP") or r.get("FDP name") or ""
            n = normalize_site(v)
            if n: sites.add(n)
        return {"uuid": uuids, "site": sites, "columns": set(xlsx_header_names(path, "data"))}
    try:
        df = read_table(path, sheet_name=None)
    except Exception:
        df = pd.DataFrame()
    out = {"uuid": set(), "site": set(), "columns": set(df.columns)}
    if df.empty:
        return out
    if "_uuid" in df.columns:
        out["uuid"] = {str(x).strip() for x in df["_uuid"] if str(x).strip()}
    site_col = detect_site_column(df.columns)
    if site_col:
        out["site"] = {normalize_site(x) for x in df[site_col] if normalize_site(x)}
    return out

def pairwise_matching(profiles: list[dict]) -> pd.DataFrame:
    keys = {p["file"]: _full_key_sets(p["path"], p["source_type"]) for p in profiles}
    rows = []
    for a, b in itertools.combinations(profiles, 2):
        ka, kb = keys[a["file"]], keys[b["file"]]
        uuid_overlap = len(ka["uuid"] & kb["uuid"]) if ka["uuid"] and kb["uuid"] else None
        site_overlap = len(ka["site"] & kb["site"]) if ka["site"] and kb["site"] else None
        cols_overlap = len(ka["columns"] & kb["columns"])
        rows.append({
            "File A": a["file"], "Type A": a["source_type"], "File B": b["file"], "Type B": b["source_type"],
            "Shared columns": cols_overlap, "UUID overlap": uuid_overlap, "Normalized site overlap": site_overlap,
            "A UUIDs": len(ka["uuid"]) or None, "B UUIDs": len(kb["uuid"]) or None,
            "A sites": len(ka["site"]) or None, "B sites": len(kb["site"]) or None,
        })
    return pd.DataFrame(rows)


def moda_file_overlap(profiles: list[dict]) -> tuple[pd.DataFrame, pd.DataFrame]:
    moda = [p for p in profiles if p["source_type"] == "MoDa"]
    keys = {p["file"]: _full_key_sets(p["path"], "MoDa") for p in moda}
    names = [p["file"] for p in moda]
    uuid_mat = pd.DataFrame(index=names, columns=names, dtype=object)
    site_mat = pd.DataFrame(index=names, columns=names, dtype=object)
    for a in names:
        for b in names:
            ua, ub = keys[a]["uuid"], keys[b]["uuid"]
            sa, sb = keys[a]["site"], keys[b]["site"]
            uuid_mat.loc[a, b] = len(ua & ub) if ua and ub else 0
            site_mat.loc[a, b] = len(sa & sb) if sa and sb else 0
    return uuid_mat, site_mat


def unmatched_sites(profiles: list[dict], reference_type: str = "MoDa") -> pd.DataFrame:
    refs = [p for p in profiles if p["source_type"] == reference_type]
    if not refs:
        return pd.DataFrame()
    ref_sites = set()
    for p in refs:
        ref_sites |= _full_key_sets(p["path"], p["source_type"])["site"]
    rows = []
    for p in profiles:
        if p["source_type"] == reference_type:
            continue
        sites = _full_key_sets(p["path"], p["source_type"])["site"]
        if not sites:
            continue
        matched = sites & ref_sites
        unmatched = sorted(sites - ref_sites)
        rows.append({
            "Source": p["file"], "Type": p["source_type"], "Sites": len(sites),
            "Matched to MoDa": len(matched), "Unmatched": len(unmatched),
            "Match %": (len(matched) / len(sites)) if sites else None,
            "Unmatched examples": ", ".join(unmatched[:12]),
        })
    return pd.DataFrame(rows)


def scoped_source_site_matching(profiles: list[dict], analysis_rows: list[dict]) -> pd.DataFrame:
    ref_sites = {normalize_site(r.get("site_reported", "")) for r in analysis_rows if normalize_site(r.get("site_reported", ""))}
    rows=[]
    for p in profiles:
        if p["source_type"] == "MoDa":
            continue
        sites=_full_key_sets(p["path"], p["source_type"])["site"]
        if not sites:
            continue
        matched=sites & ref_sites; unmatched=sorted(sites-ref_sites)
        rows.append({"Source":p["file"],"Type":p["source_type"],"Source sites":len(sites),"Matched to selected MoDa cohort":len(matched),"Unmatched":len(unmatched),"Match %":len(matched)/len(sites) if sites else None,"Unmatched examples":", ".join(unmatched[:12])})
    return pd.DataFrame(rows)
