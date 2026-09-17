"""Optional August 2026 regression test.

Usage:
python tests/regression_test.py /path/early.xlsx /path/late.xlsx
"""
from __future__ import annotations
import sys
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from pipeline import run_analysis

EXPECTED={"rows":1817,"indicators":22,"detail":414,"agg":11,"manual_month_mismatch":109,"photos":247,"photo_sites":146}

def main():
    if len(sys.argv)!=3:
        print(__doc__); return 2
    a=run_analysis(sys.argv[1:3],"2026-08","Jijiga")
    got={"rows":len(a["rows"]),"indicators":len(a["indicators"]),"detail":len(a["detail"]),"agg":len(a["agg"]),"manual_month_mismatch":a["dq"]["manual_month_mismatch"],"photos":len(a.get("photos",[])),"photo_sites":a["dq"].get("photo_evidence",{}).get("sites_with_photos",0)}
    print("Expected:",EXPECTED); print("Got:",got)
    if got!=EXPECTED:
        print("REGRESSION FAILED — review form mapping/logic before publishing.")
        return 1
    print("REGRESSION PASSED")
    return 0

if __name__=="__main__": raise SystemExit(main())
