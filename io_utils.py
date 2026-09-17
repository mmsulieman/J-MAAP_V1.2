from __future__ import annotations

import io
import os
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable

import pandas as pd

MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def safe_name(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name or "upload")
    return name[:180]


def save_uploads(uploaded_files: Iterable, workdir: str) -> list[str]:
    paths = []
    Path(workdir).mkdir(parents=True, exist_ok=True)
    for f in uploaded_files or []:
        p = Path(workdir) / safe_name(f.name)
        p.write_bytes(f.getbuffer())
        paths.append(str(p))
    return paths


def list_xlsx_sheets(path: str) -> list[str]:
    try:
        xls = pd.ExcelFile(path)
        return list(xls.sheet_names)
    except Exception:
        return []


def preview_file(path: str, sheet_name: str | int | None = None, nrows: int = 500) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path, dtype=str, keep_default_na=False, nrows=nrows, low_memory=False)
    if ext in {".xlsx", ".xlsm", ".xls"}:
        if sheet_name is None:
            sheets = list_xlsx_sheets(path)
            sheet_name = "data" if "data" in sheets else (sheets[0] if sheets else 0)
        return pd.read_excel(path, sheet_name=sheet_name, dtype=str, keep_default_na=False, nrows=nrows)
    return pd.DataFrame()


def read_table(path: str, sheet_name: str | int | None = None, max_rows: int | None = None) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path, dtype=str, keep_default_na=False, nrows=max_rows, low_memory=False)
    if ext in {".xlsx", ".xlsm", ".xls"}:
        sheets = list_xlsx_sheets(path)
        if sheet_name is None:
            sheet_name = "data" if "data" in sheets else (sheets[0] if sheets else 0)
        return pd.read_excel(path, sheet_name=sheet_name, dtype=str, keep_default_na=False, nrows=max_rows)
    return pd.DataFrame()


def classify_source(name: str, columns: Iterable[str], sheet_names: Iterable[str] | None = None) -> str:
    cols = {str(c).strip() for c in columns}
    lowname = (name or "").lower()
    sheets = {str(x).lower() for x in (sheet_names or [])}
    if "_uuid" in cols and ("WFP Sub Office:" in cols or "Activities to be assessed for this visit" in cols):
        return "MoDa"
    if "handover" in lowname or "Delivery_Recipient_Code" in cols or ("FRN_Number" in cols and "FRN_Code" not in cols):
        return "Logistics / Handover"
    if "frn" in lowname or {"FRN_Code", "Point_Of_Interest_ID"} & cols or {"FRN ID", "FRN_ID"} & cols:
        return "FRN / Programme"
    if "rbmf" in lowname or "rbmf" in sheets or "monitoring coverage" in " ".join(c.lower() for c in cols):
        return "RBMF / Monitoring Plan"
    if "action" in lowname and "tracker" in lowname:
        return "Previous Action Tracker"
    return "Other"


def detect_site_column(columns: Iterable[str]) -> str | None:
    columns = [str(c) for c in columns]
    exact = [
        "FDP name", "Specify FDP", "Point_Of_Interest", "Point_Of_Interest_Name",
        "Delivery_Recipient", "Delivery Recipient", "Destination", "Site", "site",
        "FDP", "FDP Name", "POI_Name", "Location Name", "location_name"
    ]
    for c in exact:
        if c in columns:
            return c
    priorities = ["fdp", "site", "point_of_interest", "recipient", "destination", "location"]
    lowered = {c: re.sub(r"[^a-z0-9]+", "_", c.lower()).strip("_") for c in columns}
    for p in priorities:
        for c, lc in lowered.items():
            if p in lc and not any(x in lc for x in ["id", "code", "date"]):
                return c
    return None


def normalize_site(value: str) -> str:
    s = str(value or "").strip().lower()
    s = s.replace("&", " and ")
    s = re.sub(r"[’'`´]", "", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    tokens = [t for t in s.split() if t not in {"fdp", "site", "centre", "center", "town"}]
    return " ".join(tokens)


def detect_date_column(columns: Iterable[str]) -> str | None:
    candidates = ["Enter the date:", "Monitoring Date", "Date", "date", "Dispatch_Date", "Handover_Date", "Requested_Date"]
    cols = [str(c) for c in columns]
    for c in candidates:
        if c in cols:
            return c
    for c in cols:
        if "date" in c.lower():
            return c
    return None


def xlsx_header_names(path: str, sheet_name: str = "data") -> list[str]:
    """Read only the first row of an XLSX worksheet without loading the data body."""
    import sys, posixpath
    root_dir = Path(__file__).resolve().parents[1]
    if str(root_dir) not in sys.path:
        sys.path.insert(0, str(root_dir))
    import monitoring_automation_v1 as core
    try:
        with zipfile.ZipFile(path) as z:
            ss = core._shared_strings(z)
            sp = core._sheet_path(z, sheet_name)
            for _, e in ET.iterparse(z.open(sp), events=("end",)):
                if e.tag != MAIN_NS + "row":
                    continue
                rn = int(e.attrib.get("r", "0"))
                if rn == 1:
                    vals = [core._cell_value(c, ss) for c in e.findall(MAIN_NS + "c")]
                    e.clear()
                    return vals
                e.clear()
    except Exception:
        return []
    return []


def preview_moda_core(path: str, nrows: int = 500, sub_office: str | None = None) -> pd.DataFrame:
    """Fast human-readable preview of core raw MoDa fields, preserving raw values."""
    import sys
    root_dir = Path(__file__).resolve().parents[1]
    if str(root_dir) not in sys.path:
        sys.path.insert(0, str(root_dir))
    import monitoring_automation_v1 as core
    wanted = {
        "_uuid", "_xform_id", "_submission_time", "Enter the date:", "WFP Sub Office:",
        "Geographic Coordinates", "General_information/_HHCoord_latitude", "General_information/_HHCoord_longitude",
        "Activities to be assessed for this visit", "Zone:", "Wereda", "FDP name", "Specify FDP",
        "Information collected by :", "Monitoring month:", "Please select which type or DM survey you're starting",
        "What types of WFP assistance has your household received?", "HHAsstMode/Food", "HHAsstMode/Cash",
        "1.4 What is the Sex of the household head?",
        "If you wanted to ask a question, get more information, make a complaint/appeal, report misconduct or provide feedback, do you know what to do/who to contact?",
        "Have you ever utilized the available feedback mechanisms?",
        "Have you ever been assisted/received feedback based on your complaints?",
        "Are you satisfied with the existing complaints and feedback mechanisms in place?",
        "Please take a picture of the distribution site",
        "2.2 Take photo of the storeroom",
    }
    raw, _ = core.read_moda_xlsx(path, wanted)
    if sub_office:
        target = str(sub_office).strip().casefold()
        raw = [r for r in raw if str(r.get("WFP Sub Office:", "")).strip().casefold() == target]
    return pd.DataFrame(raw[:nrows])
