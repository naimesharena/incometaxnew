"""Storage layer: return drafts (save/load/list) and government prefill
application.

Drafts: JSON snapshots of the UI input keyed by (form, PAN, AY).
Prefill: the e-filing portal supplies prefill data in the same ITR JSON
envelope used for filing (verified from PreFillJson.bas which parses
jsonObject("ITR")... and calls ImportTDSonSalary_pfl, ImportTDSOthThanSals_pfl,
ImportScheduleTDS3Dtls_pfl, etc.). We extract the key sections back into
form inputs so the user can review before computing.
"""
import json, os, re
from typing import Dict, Any

DRAFT_DIR = os.environ.get("ITR_DRAFT_DIR",
                           os.path.join(os.path.dirname(__file__), "..", "..", "data", "drafts"))


def _draft_path(form: str, pan: str, ay: str = "2026") -> str:
    pan = re.sub(r"[^A-Z0-9]", "", (pan or "UNKNOWN").upper()) or "UNKNOWN"
    os.makedirs(DRAFT_DIR, exist_ok=True)
    return os.path.join(DRAFT_DIR, f"{form}_{pan}_{ay}.json")


def save_draft(form: str, pan: str, payload: Dict[str, Any]) -> str:
    path = _draft_path(form, pan)
    with open(path, "w") as fh:
        json.dump({"form": form, "pan": pan, "assessment_year": "2026",
                   "input": payload}, fh, indent=1)
    return path


def load_draft(form: str, pan: str) -> Dict[str, Any]:
    path = _draft_path(form, pan)
    if not os.path.exists(path):
        return {}
    return json.load(open(path))


def list_drafts() -> list:
    if not os.path.isdir(DRAFT_DIR):
        return []
    out = []
    for fn in sorted(os.listdir(DRAFT_DIR)):
        try:
            d = json.load(open(os.path.join(DRAFT_DIR, fn)))
            out.append({"file": fn, "form": d.get("form"), "pan": d.get("pan"),
                        "assessment_year": d.get("assessment_year")})
        except Exception:
            continue
    return out


# ---------------------------------------------------------------------------
# Government prefill application
# ---------------------------------------------------------------------------
def apply_prefill(prefill: Dict[str, Any]) -> Dict[str, Any]:
    """Map a government prefill JSON (ITR envelope) into UI inputs.

    Supported source sections (per PreFillJson.bas import functions):
      TDSonSalaries, TDSonOthThanSals, ScheduleTCS, TaxPayments/ScheduleIT,
      salary totals, other-source interest/dividend lines.
    Returns {inputs, prefilled_sections, notes}.
    """
    itr = (prefill or {}).get("ITR", {})
    inner = itr.get("ITR1") or itr.get("ITR2") or itr.get("ITR3") or itr.get("ITR4") or {}
    inputs: Dict[str, Any] = {}
    done = []
    notes = []

    # --- personal ---
    pi = inner.get("PersonalInfo", {})
    name = pi.get("AssesseeName", {})
    if name:
        inputs["first_name"] = name.get("FirstName", "")
        inputs["middle_name"] = name.get("MiddleName", "")
        inputs["last_name"] = name.get("SurNameOrOrgName", "")
    if pi.get("PAN"):
        inputs["pan"] = pi["PAN"]
    if pi.get("DOB"):
        inputs["dob"] = pi["DOB"]
    if pi.get("PAN") or name:
        done.append("PersonalInfo")

    # --- TDS on salary ---
    tds_sal = inner.get("TDSonSalaries", {})
    entries = tds_sal.get("TDSonSalary") or []
    if entries:
        first = entries[0]
        detl = first.get("EmployerOrDeductorOrCollectDetl", {})
        inputs["tds_amount"] = first.get("TotalTDSSal", 0)
        inputs["tds_deductor"] = detl.get("EmployerOrDeductorOrCollecterName", "")
        inputs["tds_tan"] = detl.get("TAN", "")
        if len(entries) > 1:
            notes.append(f"{len(entries) - 1} additional salary-TDS entries present - "
                         "add them under Taxes Paid")
        done.append("TDSonSalaries")

    # --- TDS other than salary ---
    tds_oth = inner.get("TDSonOthThanSals", {})
    oth_entries = tds_oth.get("TDSonOthThanSal") or []
    if oth_entries:
        inputs["tds_other_amount"] = sum(e.get("TotalTDSOnOthThanSal", 0) for e in oth_entries)
        notes.append("Other-than-salary TDS captured as a single lump - review entries")
        done.append("TDSonOthThanSals")

    # --- TCS ---
    tcs = inner.get("ScheduleTCS", {})
    tcs_entries = tcs.get("TCS") or []
    if tcs_entries:
        inputs["tcs_amount"] = sum(e.get("AmtTCSClaimedThisYear", 0) for e in tcs_entries)
        done.append("ScheduleTCS")

    # --- advance tax / self assessment ---
    tp = inner.get("TaxPayments") or inner.get("ScheduleIT") or {}
    payments = tp.get("TaxPayment") or []
    if payments:
        inputs["advance_tax"] = sum(p.get("Amt", 0) for p in payments)
        done.append("TaxPayments")

    # --- salary ---
    ss = inner.get("ScheduleS", {})
    if ss.get("TotalGrossSalary"):
        inputs["gross_salary"] = ss.get("TotalGrossSalary", 0)
        done.append("ScheduleS")

    # --- other sources ---
    os_ = inner.get("ScheduleOS", {})
    if os_.get("TotOthSrcNoRaceHorse"):
        inputs["other_income"] = os_.get("TotOthSrcNoRaceHorse", 0)
        inputs["other_interest"] = os_.get("TotOthSrcNoRaceHorse", 0)
        done.append("ScheduleOS")

    return {"inputs": inputs, "prefilled_sections": done, "notes": notes}
