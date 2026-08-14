"""ITR Filing Software - web API & UI (AY 2026-27).

Endpoints:
  GET  /                     single-page filing UI
  GET  /api/constants        dropdown lists & form field specs
  POST /api/compute/{form}   build return + live computation + validation
  GET  /api/download/{form}  schema-valid CBDT JSON (query params unused;
                             POST body via /api/json/{form})
  POST /api/json/{form}      returns the CBDT JSON payload
"""
import os, sys, json

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, ".."))

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Dict, Any

from web.builder import BUILDERS, format_result
from itr_filing.storage import (save_draft, load_draft, list_drafts,
                                apply_prefill)
from itr_filing import constants as C
from itr_filing.json_builder import build_itr1_json, validate_against_schema
from itr_filing.itr2_json import build_itr2_json
from itr_filing.itr3_json import build_itr3_json
from itr_filing.itr4_json import build_itr4_json
from itr_filing.validation import validate_itr1

REPO = os.path.join(HERE, "..", "..")
SCHEMAS = {
    "ITR-1": os.path.join(REPO, "ITR-1", "ITR-1_2026_Main_V1.1.json"),
    "ITR-2": os.path.join(REPO, "ITR-2", "ITR-2_2026_Main_V1.1.json"),
    "ITR-3": os.path.join(REPO, "ITR-3", "ITR-3_2026_Main_V1.1.json"),
    "ITR-4": os.path.join(REPO, "ITR-4", "ITR-4_2026_Main_V1.1.json"),
}
JSON_BUILDERS = {"ITR-1": build_itr1_json, "ITR-2": build_itr2_json,
                 "ITR-3": build_itr3_json, "ITR-4": build_itr4_json}

app = FastAPI(title="ITR Filing Software AY 2026-27")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"],
                   allow_headers=["*"])


class Payload(BaseModel):
    input: Dict[str, Any]


def _num(key, label, section, step=1):
    return {"key": key, "label": label, "type": "number", "section": section}


def _sel(key, label, options, section):
    return {"key": key, "label": label, "type": "select",
            "options": options, "section": section}


def _txt(key, label, section, placeholder=""):
    return {"key": key, "label": label, "type": "text", "section": section,
            "placeholder": placeholder}


def _chk(key, label, section):
    return {"key": key, "label": label, "type": "checkbox", "section": section}


COMMON_PERSONAL = [
    _txt("first_name", "First Name", "Personal"),
    _txt("middle_name", "Middle Name", "Personal"),
    _txt("last_name", "Last Name / Surname", "Personal"),
    _txt("father_name", "Father's Name", "Personal"),
    _txt("pan", "PAN", "Personal", "ABCPS1234K"),
    _txt("aadhaar", "Aadhaar (12 digits)", "Personal"),
    {"key": "dob", "label": "Date of Birth", "type": "date", "section": "Personal"},
    _txt("mobile", "Mobile", "Personal"),
    _txt("email", "Email", "Personal"),
    _txt("addr1", "Flat / Door / Block No.", "Address"),
    _txt("addr2", "Name of Premises / Building", "Address"),
    _txt("locality", "Area / Locality", "Address"),
    _txt("city", "City / Town / District", "Address"),
    _sel("state", "State", [f"{k}-{v}" for k, v in C.STATE_CODES.items()], "Address"),
    _txt("pincode", "PIN Code", "Address"),
]

REGIME_FIELD = _sel("regime", "Tax Regime (115BAC)", [
    {"value": "new", "label": "New Regime (default u/s 115BAC)"},
    {"value": "old", "label": "Old Regime (opted out)"}], "Filing")

FILING_SEC_FIELD = _sel("filing_sec", "Filed u/s", [
    {"value": "11", "label": "139(1) - On or before due date"},
    {"value": "12", "label": "139(4) - Belated"},
    {"value": "17", "label": "139(5) - Revised"},
    {"value": "21", "label": "139(8A) - Updated return"}], "Filing")

BANK_FIELDS = [
    _txt("ifsc", "IFSC Code (for refund)", "Refund / Bank", "SBIN0001234"),
    _txt("account_no", "Bank Account No.", "Refund / Bank"),
    _txt("bank_name", "Bank Name", "Refund / Bank"),
]

TAXES_PAID = [
    _num("tds_amount", "TDS Amount", "Taxes Paid"),
    _txt("tds_deductor", "TDS Deductor Name", "Taxes Paid"),
    _txt("tds_tan", "Deductor TAN", "Taxes Paid", "MUMC12345A"),
    _num("advance_tax", "Advance Tax Paid", "Taxes Paid"),
    _num("self_tax", "Self Assessment Tax", "Taxes Paid"),
    _num("relief_89", "Relief u/s 89 (arrears)", "Taxes Paid"),
    {"key": "verification_date", "label": "Verification / Filing Date",
     "type": "date", "section": "Taxes Paid"},
]

HP_FIELDS = [
    _sel("hp_type", "House Property Type", ["(Select)", "Self Occupied",
                                            "Let Out", "Deemed Let Out"], "House Property"),
    _num("hp_rent", "Gross Rent / Lettable Value", "House Property"),
    _num("hp_municipal_tax", "Municipal Tax Paid", "House Property"),
    _num("hp_interest", "Interest on Borrowed Capital (24b)", "House Property"),
]

FORM_SPECS = {
    "ITR-1": {
        "title": "ITR-1 (SAHAJ) - Resident individuals: Salary/Pension, one HP, Other Sources",
        "fields": COMMON_PERSONAL + [
            REGIME_FIELD, FILING_SEC_FIELD,
            _sel("employment", "Nature of Employment", C.NATURE_OF_EMPLOYMENT, "Salary"),
            _num("gross_salary", "Gross Salary (17(1)+17(2)+17(3))", "Salary"),
            _num("basic_salary", "Basic Salary (for HRA)", "Salary"),
            _num("da", "Dearness Allowance", "Salary"),
            _num("hra", "HRA Received", "Salary"),
            _num("rent_paid", "Rent Paid", "Salary"),
            _chk("metro", "Metro City (50% HRA rule)", "Salary"),
            _num("exempt_allowances", "Exempt Allowances u/s 10", "Salary"),
            _num("professional_tax", "Professional Tax (16(iii))", "Salary"),
        ] + HP_FIELDS + [
            _num("savings_interest", "Savings Bank Interest", "Other Sources"),
            _num("other_interest", "Other Interest Income", "Other Sources"),
            _num("dividend", "Dividend Income", "Other Sources"),
            _num("family_pension", "Family Pension", "Other Sources"),
            _num("ltcg_sale", "LTCG 112A - Sale Consideration", "Capital Gains (112A)"),
            _num("ltcg_cost", "LTCG 112A - Cost of Acquisition", "Capital Gains (112A)"),
            _num("s80c", "80C", "Deductions (VI-A)"),
            _num("s80ccc", "80CCC", "Deductions (VI-A)"),
            _num("s80ccd1", "80CCD(1)", "Deductions (VI-A)"),
            _num("s80ccd1b", "80CCD(1B)", "Deductions (VI-A)"),
            _sel("s80d_selection", "80D Category", [
                {"value": "1", "label": "1 - Self & Family (non-senior) - 25k"},
                {"value": "2", "label": "2 - Self & Family incl. senior - 50k"},
                {"value": "3", "label": "3 - Parents - 25k"},
                {"value": "4", "label": "4 - Parents (senior) - 50k"},
                {"value": "5", "label": "5 - Self+Family incl. parents - 50k"},
                {"value": "6", "label": "6 - Self+Family incl. senior parents - 75k"},
                {"value": "7", "label": "7 - Self(senior)+family incl. senior parents - 75k/1L"}],
                "Deductions (VI-A)"),
            _num("s80d_premium", "80D Premium Paid", "Deductions (VI-A)"),
            _sel("s80dd_type", "80DD Disability", [
                {"value": "0", "label": "- none -"},
                {"value": "1", "label": "Dependent with disability (75,000)"},
                {"value": "2", "label": "Dependent with severe disability (1,25,000)"}],
                "Deductions (VI-A)"),
            _sel("s80u_type", "80U Self Disability", [
                {"value": "0", "label": "- none -"},
                {"value": "1", "label": "Self with disability (75,000)"},
                {"value": "2", "label": "Self with severe disability (1,25,000)"}],
                "Deductions (VI-A)"),
            _sel("s80ddb_selection", "80DDB Category", [
                {"value": "0", "label": "- none -"},
                {"value": "1", "label": "Self or dependent (40,000)"},
                {"value": "2", "label": "Senior citizen (1,00,000)"}],
                "Deductions (VI-A)"),
            _num("s80ddb_amount", "80DDB Amount Paid", "Deductions (VI-A)"),
            _num("s80e", "80E (education loan interest)", "Deductions (VI-A)"),
            _num("s80tta", "80TTA", "Deductions (VI-A)"),
            _num("s80ttb", "80TTB", "Deductions (VI-A)"),
        ] + TAXES_PAID + BANK_FIELDS,
    },
    "ITR-2": {
        "title": "ITR-2 - Individuals/HUFs: salary, HP, capital gains, other sources (no business)",
        "fields": COMMON_PERSONAL + [
            REGIME_FIELD, FILING_SEC_FIELD,
            _num("gross_salary", "Gross Salary", "Income"),
            _num("exempt_allowances", "Exempt Allowances u/s 10", "Income"),
            _num("other_income", "Income from Other Sources", "Income"),
            _num("via_total", "Chapter VI-A Deductions (total)", "Income"),
        ] + HP_FIELDS + [
            _num("stcg_amount", "Short-term CG amount (STT-paid equity)", "Capital Gains"),
            _num("stcg_cost", "STCG Cost of Acquisition", "Capital Gains"),
            {"key": "stcg_acq_date", "label": "STCG Acquisition Date",
             "type": "date", "section": "Capital Gains"},
        ] + TAXES_PAID + BANK_FIELDS,
    },
    "ITR-3": {
        "title": "ITR-3 - Individuals/HUFs with business/profession income",
        "fields": COMMON_PERSONAL + [
            REGIME_FIELD, FILING_SEC_FIELD,
            _num("gross_salary", "Gross Salary", "Income"),
            _num("exempt_allowances", "Exempt Allowances u/s 10", "Income"),
            _num("other_income", "Income from Other Sources", "Income"),
            _num("net_profit_pl", "Net Profit as per P&L", "Business (P&L)"),
            _num("inadmissible", "Inadmissible Expenses (add back)", "Business (P&L)"),
            _num("dep_as_per_pl", "Depreciation as per P&L (add back)", "Business (P&L)"),
            _num("dep_rate", "IT Depreciation Rate (e.g. 0.15)", "Depreciation (DPM)"),
            _num("dep_opening_wdv", "Opening WDV of Block", "Depreciation (DPM)"),
            _num("dep_additions", "Additions used >= 180 days", "Depreciation (DPM)"),
            _num("dep_additions_half", "Additions used < 180 days", "Depreciation (DPM)"),
            _num("via_total", "Chapter VI-A Deductions (total)", "Income"),
        ] + TAXES_PAID + BANK_FIELDS,
    },
    "ITR-4": {
        "title": "ITR-4 (SUGAM) - presumptive taxation 44AD/44ADA/44AE (TI <= 50 lakh)",
        "fields": COMMON_PERSONAL + [
            REGIME_FIELD, FILING_SEC_FIELD,
            _sel("entity", "Status", [
                {"value": "I", "label": "Individual"},
                {"value": "H", "label": "HUF"},
                {"value": "F", "label": "Firm (other than LLP)"}], "Filing"),
            _num("gross_salary", "Salary Income", "Income"),
            _num("other_income", "Income from Other Sources", "Income"),
            _num("turnover_digital", "44AD - Digital/A-c payee Turnover", "Presumptive 44AD"),
            _num("turnover_other", "44AD - Other Mode Turnover", "Presumptive 44AD"),
            _num("income_digital", "44AD - Declared Income (>= 6%)", "Presumptive 44AD"),
            _num("income_other", "44AD - Declared Income Other (>= 8%)", "Presumptive 44AD"),
            _num("gross_receipts_44ada", "44ADA - Gross Receipts", "Presumptive 44ADA"),
            _num("income_44ada", "44ADA - Declared Income (>= 50%)", "Presumptive 44ADA"),
            _num("via_total", "Chapter VI-A Deductions (total)", "Income"),
        ] + TAXES_PAID + BANK_FIELDS,
    },
}


@app.get("/api/constants")
def api_constants():
    return {"forms": FORM_SPECS, "assessment_year": "2026-27",
            "states": C.STATE_CODES}


def _build(form: str, payload: Payload):
    if form not in BUILDERS:
        raise HTTPException(404, f"Unknown form {form}")
    try:
        return BUILDERS[form](payload.input)
    except Exception as e:
        raise HTTPException(400, f"Input error: {e}")


@app.post("/api/compute/{form}")
def api_compute(form: str, payload: Payload):
    ret = _build(form, payload)
    errors = []
    if form == "ITR-1":
        v = validate_itr1(ret)
        errors = v["errors"]
    elif form == "ITR-4":
        errors = ret.validation_errors()
    result = format_result(form, ret)
    return {"result": result, "errors": errors}


@app.post("/api/json/{form}")
def api_json(form: str, payload: Payload):
    ret = _build(form, payload)
    builder = JSON_BUILDERS.get(form)
    if not builder:
        raise HTTPException(404, "no json builder")
    out = builder(ret)
    errs = validate_against_schema(out, SCHEMAS[form])
    return JSONResponse({"schema_valid": errs == [], "schema_errors": errs[:10],
                         "payload": out})


@app.get("/api/drafts")
def api_drafts_list():
    return {"drafts": list_drafts()}


@app.post("/api/draft/save/{form}")
def api_draft_save(form: str, payload: Payload):
    pan = str(payload.input.get("pan", "")).upper() or "UNKNOWN"
    path = save_draft(form, pan, payload.input)
    return {"saved": True, "file": os.path.basename(path), "form": form, "pan": pan}


@app.post("/api/draft/load/{form}")
def api_draft_load(form: str, payload: Payload):
    pan = str(payload.input.get("pan", "")).upper() or "UNKNOWN"
    d = load_draft(form, pan)
    if not d:
        raise HTTPException(404, f"No draft for {form} / {pan}")
    return d


@app.post("/api/prefill")
def api_prefill(payload: Payload):
    """Apply a government prefill JSON (ITR envelope) -> form inputs."""
    return apply_prefill(payload.input)


@app.get("/")
def index():
    return FileResponse(os.path.join(HERE, "static", "index.html"))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
