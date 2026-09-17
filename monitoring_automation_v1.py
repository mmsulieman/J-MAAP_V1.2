#!/usr/bin/env python3
"""Jijiga AO Monthly Monitoring Automation Engine v1

Purpose
-------
Read one or more MoDa XLSX/CSV exports, standardize a controlled set of monitoring
indicators, generate site-level findings, aggregated management actions, and a data
quality audit package.

Design principles
-----------------
1. Derive reporting month from monitoring date, not manually selected Monitoring Month.
2. Match fields by exact/approved question text aliases, never Excel column position.
3. Keep respondent/site-observation denominators explicit.
4. Generate proposed findings only; field validation and management agreement remain human gates.
5. Do not silently calculate an indicator when its mapped source field is missing.

This core engine deliberately writes canonical CSV/JSON outputs. Formatted XLSX trackers
are built from these canonical files by the companion artifact-tool workbook builder.
"""
from __future__ import annotations

import argparse, csv, datetime as dt, json, os, posixpath, re, sys, zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

BASE_FIELDS = {
    "activity": ["Activities to be assessed for this visit"],
    "monitoring_date": ["Enter the date:"],
    "sub_office": ["WFP Sub Office:"],
    "zone": ["Zone:"],
    "woreda": ["Wereda"],
    "site": ["FDP name"],
    "site_alt": ["Specify FDP"],
    "provider": ["Information collected by :"],
    "monitoring_month_manual": ["Monitoring month:"],
    "monitoring_type": ["Please select which type or DM survey you're starting"],
    "tsfp_checklist": ["Please select TSFP checklists"],
    "uuid": ["_uuid"],
    "xform_id": ["_xform_id"],
    "submission_time": ["_submission_time"],
    "cfm_usage_general": ["Have you ever utilized the available feedback mechanisms?"],
    "cfm_response_general": ["Have you ever been assisted/received feedback based on your complaints?"],
    "cfm_satisfaction_general": ["Are you satisfied with the existing complaints and feedback mechanisms in place?"],
}

INDICATORS = [
    dict(code="A1_CFM_AWARE", activity="Activity 1 (Relief response)", activity_short="A1 - Relief", theme="AAP/CFM",
         question="If you wanted to ask a question, get more information, make a complaint/appeal, report misconduct or provide feedback, do you know what to do/who to contact?", issue_values=["No"], severity="High",
         finding="Beneficiary does not know how/whom to contact for information, feedback or complaint",
         action="Reinforce CFM channel communication before and during distributions; verify visibility and understanding at follow-up.", owner="Programme / AAP / CP"),
    dict(code="A1_SAFETY", activity="Activity 1 (Relief response)", activity_short="A1 - Relief", theme="Protection & Safety",
         question="Are adequate measures taken to ensure your safety during the distribution?", issue_values=["No"], severity="High",
         finding="Adequate safety measures not reported during distribution",
         action="Review site organization, crowd control, access for vulnerable people and other safety arrangements; address identified gaps.", owner="Programme / Protection-AAP / CP"),
    dict(code="A1_ENT_VERBAL", activity="Activity 1 (Relief response)", activity_short="A1 - Relief", theme="Entitlement Communication",
         question="Have entitlements been clearly communicated verbally to beneficiaries prior to distributions?", issue_values=["No"], severity="Medium-High",
         finding="Entitlements not clearly communicated verbally before distribution",
         action="Reinforce pre-distribution entitlement briefing and monitor compliance.", owner="Programme / CP"),
    dict(code="A1_ENT_DISPLAY", activity="Activity 1 (Relief response)", activity_short="A1 - Relief", theme="Entitlement Communication",
         question="Are written notices or pictorial displays with the correct entitlements posted at the distribution site?", issue_values=["No"], severity="Medium-High",
         finding="Correct entitlement notice/pictorial display not posted",
         action="Ensure current entitlement/ration information is visibly displayed at the FDP and verified during monitoring.", owner="Programme / CP"),
    dict(code="A1_BEN_LIST", activity="Activity 1 (Relief response)", activity_short="A1 - Relief", theme="Beneficiary Verification",
         question="Do you think there are any problems or concerns with how beneficiaries have been selected for inclusion for the distribution (not related to targeting)? As in, are there signs of potential fraud at this distribution regarding beneficiary lists, electronic records, entitlement cards, etc?", issue_values=["Yes"], severity="High",
         finding="Potential concerns with beneficiary selection/list controls observed",
         action="Conduct targeted verification of beneficiary lists/cards/electronic records and distribution documentation.", owner="Programme / RAM-M&E / CP"),
    dict(code="A1_NO_ENTITLE", activity="Activity 1 (Relief response)", activity_short="A1 - Relief", theme="Entitlement",
         question="Did you receive your in-kind entitlement?", issue_values=["No"], severity="High",
         finding="Respondent reported not receiving in-kind entitlement",
         action="Verify distribution records and entitlement status for affected site/households and document corrective action.", owner="Programme / CP / RAM-M&E"),
    dict(code="A1_PAYMENT", activity="Activity 1 (Relief response)", activity_short="A1 - Relief", theme="Protection / Misconduct",
         question="Did you pay to receive your entitlement? (either in kind or cash)", issue_values=["Yes"], severity="Critical",
         finding="Respondent reported paying to receive assistance",
         action="Immediately verify the allegation through appropriate programme/AAP assurance channels and take corrective action if substantiated.", owner="Programme / AAP / RAM-M&E"),
    dict(code="A1_CFM_VISIBLE", activity="Activity 1 (Relief response)", activity_short="A1 - Relief", theme="AAP/CFM",
         question="Is the Community Feedback Mechanism (CFM) information clearly visible throughout the distribution site?", issue_values=["No"], severity="Medium-High",
         finding="CFM information not clearly visible at distribution site",
         action="Display standardized CFM information prominently and verify during next monitoring visit.", owner="AAP / Programme / CP"),
    dict(code="A2_CFM_BANNER", activity="Activity 2 (Nutrition assistance)", activity_short="A2 - Nutrition", theme="AAP/CFM",
         question="5. Is the Complaints, Feedback Mechanism (CFM) banner has been displayed at TSFP Center?", issue_values=["No"], severity="High",
         finding="CFM banner not displayed at TSFP centre",
         action="Ensure standardized CFM banner/materials are displayed at functioning TSFP sites and verify at follow-up.", owner="Nutrition / AAP / CP / Health"),
    dict(code="A2_MUAC", activity="Activity 2 (Nutrition assistance)", activity_short="A2 - Nutrition", theme="Programme Quality",
         question="1. Is MUAC properly measured as per the standard MUAC measurement procedure?", issue_values=["No"], severity="High",
         finding="MUAC not measured according to standard procedure",
         action="Provide targeted coaching/refresher and verify correct MUAC measurement during follow-up.", owner="Nutrition / Health / CP"),
    dict(code="A2_NONREG", activity="Activity 2 (Nutrition assistance)", activity_short="A2 - Nutrition", theme="Beneficiary Verification",
         question="3. Did you observe any non-registered beneficiaries receiving TSF assistance?", issue_values=["Yes"], severity="High",
         finding="Non-registered beneficiary observed receiving TSF assistance",
         action="Review admission/registration and beneficiary verification records; address control gaps.", owner="Nutrition / Health / CP / RAM-M&E"),
    dict(code="A2_CFM_AWARE", activity="Activity 2 (Nutrition assistance)", activity_short="A2 - Nutrition", theme="AAP/CFM",
         question="1. Do you know how to present your complaint or feedback regarding the TSFP in case of need?", issue_values=["No"], severity="High",
         finding="Respondent does not know how to present TSFP complaint/feedback",
         action="Strengthen beneficiary orientation on complaint/feedback channels at TSFP sites.", owner="Nutrition / AAP / CP"),
    dict(code="A2_DELIVERY_LATE", activity="Activity 2 (Nutrition assistance)", activity_short="A2 - Nutrition", theme="Supply / Timeliness",
         question="1. Did the food delivery occur on time to TSF Center?", issue_values=["No"], severity="Medium-High",
         finding="Food delivery to TSF centre reported late",
         action="Review delivery timing and supply documentation; address recurring causes of delay.", owner="Nutrition / Supply Chain / CP"),
    dict(code="A3_CFM_AWARE", activity="Activity 3 (Refugee operations)", activity_short="A3 - Refugee", theme="AAP/CFM",
         question="If you wanted to ask a question, get more information, make a complaint/appeal, report misconduct or provide feedback, do you know what to do/who to contact?", issue_values=["No"], severity="High",
         finding="Refugee respondent does not know how/whom to contact for information or complaint",
         action="Reinforce CFM communication and channel visibility across refugee sites.", owner="Refugee Programme / AAP / CP"),
    dict(code="A3_STOCK_RECON", activity="Activity 3 (Refugee operations)", activity_short="A3 - Refugee", theme="Stock Control",
         question="3.2.27. Does the physical stock count reconcile with the recorded?", issue_values=["No"], severity="High",
         finding="Physical stock did not reconcile with recorded stock",
         action="Conduct focused physical/document stock reconciliation and document resolution.", owner="Refugee Programme / Supply Chain / CP"),
    dict(code="A6_END_INFO", activity="Activity 6 (resilience)", activity_short="A6 - Resilience", theme="Programme Communication",
         question="K1. Were you informed when this intervention will end?", issue_values=["No"], severity="Medium-High",
         finding="Participant was not informed when the intervention will end",
         action="Communicate intervention duration, milestones and exit/transition expectations to participants.", owner="Resilience / CP"),
    dict(code="A6_RECORDS", activity="Activity 6 (resilience)", activity_short="A6 - Resilience", theme="Capacity / Record Keeping",
         question="D8. Do you keep records of your farming activities", issue_values=["No"], severity="Medium",
         finding="Participant does not keep farming activity records",
         action="Strengthen practical farm record-keeping coaching and follow-up.", owner="Resilience / CP / Technical counterparts"),
    dict(code="A6_POSTHARVEST", activity="Activity 6 (resilience)", activity_short="A6 - Resilience", theme="Capacity / Post-harvest",
         question="F6. Did you receive any training on proper harvesting, drying, shelling and storage management in the last one year?", issue_values=["No"], severity="Medium",
         finding="Participant did not receive post-harvest management training in the last year",
         action="Review training coverage and provide targeted post-harvest management support where required.", owner="Resilience / CP / Technical counterparts"),
    dict(code="A6_MARKETTRAIN", activity="Activity 6 (resilience)", activity_short="A6 - Resilience", theme="Market Access",
         question="G5. Have you been trained on market access and aggregation system since you became part of this project?", issue_values=["No"], severity="Medium",
         finding="Participant not trained on market access and aggregation systems",
         action="Expand market-access and aggregation training based on site needs.", owner="Resilience / CP / Technical counterparts"),
    dict(code="A6_MARKETINFO", activity="Activity 6 (resilience)", activity_short="A6 - Resilience", theme="Market Access",
         question="G15. Do you get market information for your produce (crop and livestock)?", issue_values=["No"], severity="Medium",
         finding="Participant does not receive market information for produce",
         action="Strengthen access to timely market information and aggregation linkages.", owner="Resilience / CP / Technical counterparts"),
    dict(code="A6_NUTRITION", activity="Activity 6 (resilience)", activity_short="A6 - Resilience", theme="Nutrition-sensitive Agriculture",
         question="C5. Were you sensitised on the benefits of producing and consuming different nutritious foods e.g., vegetables, fruits, etc?", issue_values=["No"], severity="Medium",
         finding="Participant not sensitized on diverse nutritious food production/consumption",
         action="Integrate/strengthen nutrition-sensitive agriculture messaging and practical sensitization.", owner="Resilience / Nutrition / CP"),
    dict(code="A6_CFM_AWARE", activity="Activity 6 (resilience)", activity_short="A6 - Resilience", theme="AAP/CFM",
         question="K3. Do you know where to report if you have a complaint?", issue_values=["No"], severity="High",
         finding="Participant does not know where to report a complaint",
         action="Integrate CFM orientation into resilience community engagement and group meetings.", owner="Resilience / AAP / CP"),
]

AGG_GROUPS = [
    dict(title="Strengthen CFM awareness and visibility across activities", codes=["A1_CFM_AWARE","A1_CFM_VISIBLE","A2_CFM_BANNER","A2_CFM_AWARE","A3_CFM_AWARE","A6_CFM_AWARE"], priority="High", owner="Programme / AAP / CPs"),
    dict(title="Improve safety arrangements at Relief distribution sites", codes=["A1_SAFETY"], priority="High", owner="Programme / Protection-AAP / CP"),
    dict(title="Reinforce Relief entitlement communication and FDP display", codes=["A1_ENT_VERBAL","A1_ENT_DISPLAY"], priority="Medium-High", owner="Programme / CP"),
    dict(title="Verify reported entitlement non-receipt and payment allegations", codes=["A1_NO_ENTITLE","A1_PAYMENT"], priority="Critical", owner="Programme / RAM-M&E / CP"),
    dict(title="Verify Relief beneficiary-list and selection controls", codes=["A1_BEN_LIST"], priority="High", owner="Programme / RAM-M&E / CP"),
    dict(title="Strengthen TSFP beneficiary verification/admission controls", codes=["A2_NONREG"], priority="High", owner="Nutrition / Health counterparts / CP"),
    dict(title="Improve MUAC measurement procedure quality", codes=["A2_MUAC"], priority="High", owner="Nutrition / Health counterparts"),
    dict(title="Follow up TSFP food delivery timeliness exceptions", codes=["A2_DELIVERY_LATE"], priority="Medium-High", owner="Nutrition / Supply Chain / CP"),
    dict(title="Verify Refugee stock reconciliation exceptions", codes=["A3_STOCK_RECON"], priority="High", owner="Refugee Programme / Supply Chain / CP"),
    dict(title="Clarify Activity 6 intervention duration and beneficiary information", codes=["A6_END_INFO"], priority="Medium-High", owner="Resilience / CP"),
    dict(title="Strengthen Activity 6 record keeping, post-harvest, market and nutrition training", codes=["A6_RECORDS","A6_POSTHARVEST","A6_MARKETTRAIN","A6_MARKETINFO","A6_NUTRITION"], priority="Medium", owner="Resilience / CP / technical counterparts"),
]

def excel_serial_to_date(v: str) -> dt.date | None:
    v=(v or '').strip()
    if not v: return None
    try:
        n=float(v)
        return (dt.datetime(1899,12,30)+dt.timedelta(days=n)).date()
    except Exception:
        pass
    for fmt in ("%Y-%m-%d","%d/%m/%Y","%m/%d/%Y","%Y-%m-%dT%H:%M:%S"):
        try: return dt.datetime.strptime(v[:19],fmt).date()
        except Exception: pass
    return None

def col_letters(ref: str) -> str:
    m=re.match(r"([A-Z]+)",ref or "A1")
    return m.group(1) if m else "A"

def _sheet_path(z: zipfile.ZipFile, sheet_name="data") -> str:
    wb=ET.fromstring(z.read("xl/workbook.xml"))
    rel=ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    relmap={r.attrib["Id"]: r.attrib["Target"] for r in rel}
    sheets=wb.find(MAIN_NS+"sheets")
    for sh in sheets:
        if sh.attrib.get("name")==sheet_name:
            target=relmap[sh.attrib[REL_NS+"id"]].lstrip("/")
            return target if target.startswith("xl/") else posixpath.normpath("xl/"+target)
    raise ValueError(f"Worksheet {sheet_name!r} not found")

def _shared_strings(z: zipfile.ZipFile):
    if "xl/sharedStrings.xml" not in z.namelist(): return []
    root=ET.fromstring(z.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter(MAIN_NS+"t")) for si in root.findall(MAIN_NS+"si")]

def _cell_value(c, ss):
    t=c.attrib.get("t")
    if t=="inlineStr":
        node=c.find(MAIN_NS+"is")
        return "".join(x.text or "" for x in node.iter(MAIN_NS+"t")) if node is not None else ""
    v=c.find(MAIN_NS+"v")
    if v is None: return ""
    raw=v.text or ""
    if t=="s":
        try: return ss[int(raw)]
        except Exception: return raw
    return raw

def all_required_headers():
    headers=set()
    for aliases in BASE_FIELDS.values(): headers.update(aliases)
    for ind in INDICATORS: headers.add(ind["question"])
    return headers

def read_moda_xlsx(path: str, wanted_headers: set[str]):
    """Stream only selected columns from MoDa data sheet.

    Performance note: after row 1, cell values are decoded only for selected columns.
    This is important for wide XLSForms with several thousand possible question columns.
    """
    out=[]; meta={"file":os.path.basename(path),"matched_headers":[],"missing_headers":[]}
    with zipfile.ZipFile(path) as z:
        ss=_shared_strings(z); sp=_sheet_path(z,"data")
        selected_cols={}; header_values={}
        for ev,e in ET.iterparse(z.open(sp),events=("end",)):
            if e.tag!=MAIN_NS+"row": continue
            rn=int(e.attrib.get("r","0"))
            if rn==1:
                vals={}
                for c in e.findall(MAIN_NS+"c"):
                    col=col_letters(c.attrib.get("r","A1"))
                    vals[col]=_cell_value(c,ss)
                header_values=vals
                selected_cols={col:txt for col,txt in vals.items() if txt in wanted_headers}
                meta["matched_headers"]=sorted(set(selected_cols.values()))
                meta["missing_headers"]=sorted(wanted_headers-set(meta["matched_headers"]))
                meta["header_count"]=len(vals)
            else:
                rec={}
                for c in e.findall(MAIN_NS+"c"):
                    col=col_letters(c.attrib.get("r","A1"))
                    if col not in selected_cols:
                        continue
                    v=_cell_value(c,ss)
                    key=selected_cols[col]
                    if key not in rec or (not rec.get(key) and v not in (None, "")):
                        rec[key]=v
                if rec: out.append(rec)
            e.clear()
    return out, meta

def canonicalize(raw: dict) -> dict:
    rec={}
    for key,aliases in BASE_FIELDS.items():
        rec[key]=""
        for a in aliases:
            if raw.get(a,"") not in (None,""):
                rec[key]=str(raw[a]).strip(); break
    for ind in INDICATORS:
        rec[ind["code"]]=str(raw.get(ind["question"],"") or "").strip()
    rec["site_reported"]=(rec.get("site_alt") or rec.get("site") or "").strip()
    d=excel_serial_to_date(rec.get("monitoring_date","")); rec["date_obj"]=d
    rec["monitoring_date_iso"]=d.isoformat() if d else ""
    return rec

def month_bounds(month: str):
    y,m=[int(x) for x in month.split("-")]
    start=dt.date(y,m,1)
    end=dt.date(y+1,1,1)-dt.timedelta(days=1) if m==12 else dt.date(y,m+1,1)-dt.timedelta(days=1)
    return start,end

def month_label(month: str):
    y,m=[int(x) for x in month.split("-")]
    return dt.date(y,m,1).strftime("%b-%Y")

def load_inputs(paths, reporting_month, sub_office="Jijiga"):
    required=all_required_headers(); all_rows=[]; schema=[]
    for p in paths:
        p=str(p)
        if p.lower().endswith(".xlsx"):
            raw,meta=read_moda_xlsx(p,required); schema.append(meta)
        elif p.lower().endswith(".csv"):
            with open(p,encoding="utf-8-sig",newline="") as f: raw=list(csv.DictReader(f))
            hs=set(raw[0].keys()) if raw else set(); schema.append({"file":os.path.basename(p),"matched_headers":sorted(required&hs),"missing_headers":sorted(required-hs),"header_count":len(hs)})
        else: continue
        all_rows.extend(canonicalize(r) for r in raw)
    start,end=month_bounds(reporting_month)
    filt=[r for r in all_rows if str(r.get("sub_office", "")).strip().casefold()==str(sub_office).strip().casefold() and r.get("date_obj") and start<=r["date_obj"]<=end]
    # UUID dedup protects against cumulative/re-exported files.
    dedup=[]; seen=set(); duplicate_uuid=0
    for r in filt:
        uid=r.get("uuid") or ""
        if uid:
            if uid in seen: duplicate_uuid+=1; continue
            seen.add(uid)
        dedup.append(r)
    return dedup,schema,{"pre_dedup_rows":len(filt),"duplicate_uuid_rows":duplicate_uuid,"unique_uuid_count":len(seen)}

def indicator_summary(rows):
    out=[]
    for ind in INDICATORS:
        vals=[r[ind["code"]] for r in rows if r.get("activity")==ind["activity"] and r.get(ind["code"])]
        issues=sum(v.strip().lower() in {x.lower() for x in ind["issue_values"]} for v in vals)
        sites={r["site_reported"] for r in rows if r.get("activity")==ind["activity"] and r.get(ind["code"]) and r[ind["code"]].strip().lower() in {x.lower() for x in ind["issue_values"]}}
        out.append({**ind,"applicable":len(vals),"issues":issues,"issue_rate":issues/len(vals) if vals else None,"affected_sites":len(sites)})
    return out

def detailed_findings(rows, month):
    result=[]; idx=1
    for ind in INDICATORS:
        issue_set={x.lower() for x in ind["issue_values"]}
        groups=defaultdict(lambda:{"applicable":0,"issues":0,"dates":[]})
        for r in rows:
            if r.get("activity")!=ind["activity"]: continue
            v=r.get(ind["code"],"").strip()
            if not v: continue
            key=(r.get("zone",""),r.get("woreda",""),r.get("site_reported",""))
            g=groups[key]; g["applicable"]+=1
            if v.lower() in issue_set: g["issues"]+=1
            if r.get("date_obj"): g["dates"].append(r["date_obj"])
        for (zone,woreda,site),g in sorted(groups.items()):
            if g["issues"]<=0: continue
            dates=g["dates"] or [None]
            result.append({
                "Action_ID":f"ACT-{month.replace('-','')}-{idx:04d}", "Finding_Code":ind["code"], "Month":month_label(month),
                "Activity":ind["activity_short"],"Zone":zone,"Woreda":woreda,"Site":site,"Theme":ind["theme"],"Finding":ind["finding"],"Severity":ind["severity"],
                "First_Observed":min(dates).isoformat() if dates[0] else "", "Last_Observed":max(dates).isoformat() if dates[0] else "",
                "Issue_Count":g["issues"],"Applicable_Count":g["applicable"],"Issue_Rate":g["issues"]/g["applicable"] if g["applicable"] else None,
                "Proposed_Action":ind["action"],"Suggested_Responsible_Unit":ind["owner"],"Responsible_Person":"","Due_Date":"","Status":"Proposed",
                "Validation_Status":"Pending field validation","Management_Agreement":"Not reviewed","Completion_Date":"","Verification_Status":"Not started","Verified_By":"","Closure_Evidence":"","Management_Comments":"",
                "Source":f"MoDa {month_label(month)}","Action_Basis":"Automated signal derived from configured indicator; field validation and management agreement required"
            }); idx+=1
    return result

def aggregate_actions(detail, ind_summary, month):
    by_code={x["code"]:x for x in ind_summary}; out=[]
    for i,g in enumerate(AGG_GROUPS,1):
        codes=[c for c in g["codes"] if by_code.get(c,{}).get("issues",0)>0]
        drows=[r for r in detail if r["Finding_Code"] in codes]
        if not drows: continue
        acts=sorted({r["Activity"] for r in drows}); themes=sorted({r["Theme"] for r in drows}); zones=sorted({r["Zone"] for r in drows if r["Zone"]}); woredas=sorted({r["Woreda"] for r in drows if r["Woreda"]}); sites=sorted({r["Site"] for r in drows if r["Site"]})
        ev=[]
        for c in codes:
            s=by_code[c]; ev.append(f"{c}: {s['issues']}/{s['applicable']} ({s['issue_rate']:.0%}), {s['affected_sites']} sites")
        total_issue=sum(r["Issue_Count"] for r in drows); total_app=sum(r["Applicable_Count"] for r in drows)
        actions=[]
        for c in codes:
            a=by_code[c]["action"]
            if a not in actions: actions.append(a)
        out.append({
            "Aggregated_Action_ID":f"AG-{month.replace('-','')}-{i:03d}","Month":month_label(month),"Aggregated_Action_Title":g["title"],
            "Activities":", ".join(acts),"Themes":", ".join(themes),"Priority":g["priority"],"Zones":", ".join(zones),"Woredas":", ".join(woredas),"Affected_Sites":len(sites),
            "Evidence_Summary":"; ".join(ev),"Total_Issue_Count":total_issue,"Total_Applicable_Count":total_app,"Overall_Issue_Rate":total_issue/total_app if total_app else None,
            "Detailed_Action_Records":len(drows),"Proposed_Management_Action":"; ".join(actions),"Suggested_Responsible_Unit":g["owner"],"Responsible_Person":"","Due_Date":"","Status":"Proposed",
            "Validation_Status":"Pending field validation","Management_Agreement":"Not reviewed","Completion_Date":"","Verification_Status":"Not started","Verified_By":"","Closure_Evidence":"","Management_Comments":"",
            "Source_Finding_Codes":", ".join(codes),"Detailed_Action_IDs":", ".join(r["Action_ID"] for r in drows),"Action_Basis":"Aggregated from automated site-level signals; field validation and management agreement required"
        })
    return out

def data_quality(rows, schema, intake, month):
    issues=[]
    # Manual month mismatch
    y,m=[int(x) for x in month.split("-")]; expected=dt.date(y,m,1).strftime("%B")
    manual_present=[r for r in rows if r.get("monitoring_month_manual")]
    mismatch=[r for r in manual_present if r.get("monitoring_month_manual","").strip().lower()!=expected.lower()]
    if mismatch: issues.append(dict(category="Date/Period",severity="Medium",issue="Manual Monitoring Month differs from monitoring date",count=len(mismatch),recommendation="Use monitoring date as reporting-month source of truth; retain manual month only as a DQ check."))
    # Missing geography/site
    for field,label in [("zone","Zone"),("woreda","Woreda"),("site_reported","Site")]:
        n=sum(not r.get(field) for r in rows)
        if n: issues.append(dict(category="Completeness",severity="High" if field=="site_reported" else "Medium",issue=f"Missing {label}",count=n,recommendation=f"Resolve missing {label.lower()} before management reporting and site-level action assignment."))
    # General CFM skip consistency
    use_no_resp_yes=sum(r.get("cfm_usage_general","").lower()=="no" and r.get("cfm_response_general","").lower()=="yes" for r in rows)
    use_blank_resp=sum(not r.get("cfm_usage_general") and bool(r.get("cfm_response_general")) for r in rows)
    if use_no_resp_yes: issues.append(dict(category="Skip logic",severity="High",issue="CFM response recorded where CFM usage = No",count=use_no_resp_yes,recommendation="Correct/enforce skip logic; do not publish response-rate KPI without nested journey validation."))
    if use_blank_resp: issues.append(dict(category="Skip logic",severity="Medium",issue="CFM response populated while usage is blank",count=use_blank_resp,recommendation="Review form branching and legacy records; treat response-stage KPI cautiously."))
    # Duplicate UUIDs handled at intake
    if intake.get("duplicate_uuid_rows"):
        issues.append(dict(category="Duplicates",severity="High",issue="Duplicate UUID records across input exports",count=intake["duplicate_uuid_rows"],recommendation="Deduplicate by _uuid before analysis; retain duplicate log."))
    # Small denominators
    summ=indicator_summary(rows)
    small=[s for s in summ if s["applicable"] and s["applicable"]<10]
    for s in small:
        issues.append(dict(category="Denominator",severity="Medium",issue=f"Small denominator for {s['code']} ({s['applicable']} applicable)",count=s["applicable"],recommendation="Treat percentage as a targeted assurance signal; show n/N and avoid broad generalization."))
    return {
        "reporting_month":month,"rows":len(rows),"uuid_unique":intake.get("unique_uuid_count",0),"duplicate_uuid_rows":intake.get("duplicate_uuid_rows",0),
        "activities":Counter(r.get("activity","") for r in rows),"providers":Counter(r.get("provider","") for r in rows),"xform_ids":Counter(r.get("xform_id","") for r in rows),
        "date_min":min((r["date_obj"] for r in rows if r.get("date_obj")),default=None),"date_max":max((r["date_obj"] for r in rows if r.get("date_obj")),default=None),
        "manual_month_present":len(manual_present),"manual_month_mismatch":len(mismatch),"schema":schema,"issues":issues,
    }

def write_csv(path, rows, fields=None):
    Path(path).parent.mkdir(parents=True,exist_ok=True)
    if fields is None: fields=list(rows[0].keys()) if rows else []
    with open(path,"w",encoding="utf-8-sig",newline="") as f:
        w=csv.DictWriter(f,fieldnames=fields,extrasaction="ignore"); w.writeheader(); w.writerows(rows)

def json_safe(x):
    if isinstance(x,Counter): return dict(x)
    if isinstance(x,(dt.date,dt.datetime)): return x.isoformat()
    raise TypeError(type(x).__name__)

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("inputs",nargs="+"); ap.add_argument("--month",required=True,help="YYYY-MM"); ap.add_argument("--sub-office",default="Jijiga"); ap.add_argument("--outdir",required=True)
    args=ap.parse_args(); out=Path(args.outdir); out.mkdir(parents=True,exist_ok=True)
    rows,schema,intake=load_inputs(args.inputs,args.month,args.sub_office)
    summ=indicator_summary(rows); detail=detailed_findings(rows,args.month); agg=aggregate_actions(detail,summ,args.month); dq=data_quality(rows,schema,intake,args.month)
    write_csv(out/"standardized_submissions.csv",[{k:(v.isoformat() if isinstance(v,dt.date) else v) for k,v in r.items() if k!="date_obj"} for r in rows])
    write_csv(out/"indicator_summary.csv",summ)
    write_csv(out/"detailed_findings.csv",detail)
    write_csv(out/"aggregated_actions.csv",agg)
    write_csv(out/"dq_issues.csv",dq["issues"],fields=["category","severity","issue","count","recommendation"])
    with open(out/"dq_summary.json","w",encoding="utf-8") as f: json.dump(dq,f,indent=2,default=json_safe)
    with open(out/"run_summary.json","w",encoding="utf-8") as f: json.dump({"rows":len(rows),"detailed_findings":len(detail),"aggregated_actions":len(agg),"indicator_count":len(summ),"reporting_month":args.month},f,indent=2)
    print(json.dumps({"rows":len(rows),"detailed_findings":len(detail),"aggregated_actions":len(agg),"indicator_count":len(summ),"dq_issues":len(dq['issues'])},indent=2))

if __name__=="__main__": main()
