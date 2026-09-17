from __future__ import annotations

from collections import Counter, defaultdict
import pandas as pd


def rows_to_df(rows: list[dict]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    if "date_obj" in df.columns:
        df["monitoring_date"] = pd.to_datetime(df["date_obj"], errors="coerce")
    return df


def overview_metrics(rows: list[dict]) -> dict:
    df = rows_to_df(rows)
    if df.empty:
        return {"submissions": 0, "sites": 0, "woredas": 0, "activities": 0, "providers": 0}
    return {
        "submissions": len(df),
        "sites": int(df.get("site_reported", pd.Series(dtype=str)).replace("", pd.NA).nunique()),
        "woredas": int(df.get("woreda", pd.Series(dtype=str)).replace("", pd.NA).nunique()),
        "activities": int(df.get("activity", pd.Series(dtype=str)).replace("", pd.NA).nunique()),
        "providers": int(df.get("provider", pd.Series(dtype=str)).replace("", pd.NA).nunique()),
    }


def cfm_journey(rows: list[dict], activity: str | None = None, modality: str | None = None, hh_gender: str | None = None) -> dict:
    """Strict nested CFM journey. Returns base and retained count/% at each available stage.

    Relief/Refugee: Awareness -> channel identified (Access proxy) -> Usage -> Response -> Satisfaction.
    Nutrition: Awareness -> channel identified -> Usage (Response/Satisfaction unavailable in module).
    Resilience: Awareness -> channel identified -> Usage -> resolution information (Response proxy).
    """
    subset = rows
    if activity:
        subset = [r for r in subset if r.get("activity") == activity]
    if modality:
        m = modality.lower()
        truthy = {"true", "1", "yes"}
        food_yes = lambda r: str(r.get("hh_food", "")).strip().lower() in truthy
        cash_yes = lambda r: str(r.get("hh_cash", "")).strip().lower() in truthy
        if m == "food":
            subset = [r for r in subset if food_yes(r) and not cash_yes(r)]
        elif m == "cash":
            subset = [r for r in subset if cash_yes(r) and not food_yes(r)]
        elif m == "mixed":
            subset = [r for r in subset if food_yes(r) and cash_yes(r)]
    if hh_gender:
        subset = [r for r in subset if str(r.get("hh_head_sex", "")).lower() == hh_gender.lower()]

    if activity in {"Activity 1 (Relief response)", "Activity 3 (Refugee operations)"}:
        base = [r for r in subset if r.get("cfm_awareness_general")]
        awareness = [r for r in base if r.get("cfm_awareness_general", "").lower() == "yes"]
        access = [r for r in awareness if bool(r.get("cfm_channels_general"))]
        usage = [r for r in access if r.get("cfm_usage_general", "").lower() == "yes"]
        response = [r for r in usage if r.get("cfm_response_general", "").lower() == "yes"]
        satisfaction = [r for r in response if r.get("cfm_satisfaction_general", "").lower() == "yes"]
        stages = [("Awareness", awareness), ("Access", access), ("Usage", usage), ("Response", response), ("Satisfaction", satisfaction)]
    elif activity == "Activity 2 (Nutrition assistance)":
        base = [r for r in subset if r.get("tsfp_cfm_awareness")]
        awareness = [r for r in base if r.get("tsfp_cfm_awareness", "").lower() == "yes"]
        access = [r for r in awareness if bool(r.get("tsfp_cfm_channel"))]
        usage = [r for r in access if r.get("tsfp_cfm_usage", "").lower() == "yes"]
        stages = [("Awareness", awareness), ("Access", access), ("Usage", usage), ("Response", []), ("Satisfaction", [])]
    elif activity == "Activity 6 (resilience)":
        base = [r for r in subset if r.get("res_cfm_awareness")]
        awareness = [r for r in base if r.get("res_cfm_awareness", "").lower() == "yes"]
        access = [r for r in awareness if bool(r.get("res_cfm_channel"))]
        usage = [r for r in access if r.get("res_cfm_usage", "").lower() == "yes"]
        response = [r for r in usage if bool(r.get("res_cfm_resolution_time"))]
        stages = [("Awareness", awareness), ("Access", access), ("Usage", usage), ("Response", response), ("Satisfaction", [])]
    else:
        return {"base": 0, "stages": []}
    n = len(base)
    return {"base": n, "stages": [{"stage": name, "count": len(vals), "pct_base": len(vals)/n if n else None} for name, vals in stages]}


def cfm_woreda_table(rows: list[dict], activity: str, min_n: int = 5) -> pd.DataFrame:
    woredas = sorted({r.get("woreda", "") for r in rows if r.get("activity") == activity and r.get("woreda")})
    output = []
    for w in woredas:
        res = cfm_journey([r for r in rows if r.get("woreda") == w], activity=activity)
        row = {"Woreda": w, "N": res["base"]}
        for s in res["stages"]:
            row[s["stage"]] = s["pct_base"] if res["base"] >= min_n and s["pct_base"] is not None else None
        output.append(row)
    return pd.DataFrame(output)


def protection_summary(indicator_summary: list[dict]) -> pd.DataFrame:
    rows = []
    for r in indicator_summary:
        theme = str(r.get("theme", ""))
        if any(x in theme.lower() for x in ["protection", "beneficiary verification", "entitlement", "stock control"]):
            rows.append({
                "Activity": r.get("activity_short"), "Theme": theme, "Indicator": r.get("code"),
                "Finding": r.get("finding"), "Issues": r.get("issues"), "Applicable": r.get("applicable"),
                "Issue rate": r.get("issue_rate"), "Severity": r.get("severity"), "Affected sites": r.get("affected_sites"),
            })
    return pd.DataFrame(rows)
