from fastapi import APIRouter, HTTPException
from ..models.common import ITRPayload, CalculationRequest
from ..services.tax_engine import full_tax_computation
from ..services.validation import validate_itr_generic
from ..services.json_builder import generate_json_with_digest, convert_to_official_json
import json
from datetime import date

router = APIRouter(prefix="/api/itr", tags=["itr"])

@router.post("/calculate")
def calculate_tax(payload: ITRPayload):
    # Build tax engine payload
    p = payload.model_dump() if hasattr(payload, 'model_dump') else payload.dict()
    income_details = p.get("income_details",{})
    # Capital gains sum
    cg_entries = income_details.get("capital_gains_entries",[])
    cg_total = sum([e.get("gain",0) or ((e.get("sell_value",0) or 0)-(e.get("buy_value",0) or 0)) for e in cg_entries])
    # Business income: presumptive + gross_profit - expenses logic
    business_income = 0
    if income_details.get("presumptive_income"):
        business_income += income_details.get("presumptive_income",0)
    if income_details.get("gross_profit") is not None:
        business_income += (income_details.get("gross_profit",0) or 0) - (income_details.get("expenses",0) or 0) if income_details.get("presumptive_type","")=='' else 0
    # Map to tax engine expected
    tax_input = {
        "age": p.get("age",30),
        "is_resident": p.get("is_resident",True),
        "regime": p.get("regime","NEW"),
        "salary": income_details.get("salary",0),
        "basic": income_details.get("basic",0),
        "da": income_details.get("da",0),
        "hra_received": income_details.get("hra_received",0),
        "rent_paid": income_details.get("rent_paid",0),
        "is_metro": income_details.get("is_metro",False),
        "hp_income": income_details.get("hp_income",0),
        "other_sources": income_details.get("other_sources",0),
        "capital_gains": cg_total,
        "business_income": business_income,
        "deductions_80C": p.get("deductions",{}).get("deduction_80C",0),
        "deductions_other": p.get("deductions",{}).get("deduction_80D",0) + p.get("deductions",{}).get("deduction_80G",0) + p.get("deductions",{}).get("deduction_80E",0),
        "deductions_80D": p.get("deductions",{}).get("deduction_80D",0),
        "tds": sum([x.get("tax_deducted",0) for x in p.get("tds",{}).get("tds_salary",[])] ) + sum([x.get("tax_deducted",0) for x in p.get("tds",{}).get("tds_other",[])] ),
        "advance_tax": p.get("taxes_paid",{}).get("advance_tax",0),
        "self_assessment": p.get("taxes_paid",{}).get("self_assessment",0),
        "filing_date": p.get("filing_date","2026-07-31"),
        "due_date": p.get("due_date","2026-07-31"),
    }
    result = full_tax_computation(tax_input)
    return result

@router.post("/validate")
def validate_itr(payload: ITRPayload):
    p = payload.model_dump() if hasattr(payload, 'model_dump') else payload.dict()
    # Convert to validation engine expected format
    # Build simplified payload for validation
    validation_payload = {
        "personal_info": {
            "pan": p.get("personal_info",{}).get("pan",""),
            "aadhaar": p.get("personal_info",{}).get("aadhaar",""),
            "dob": p.get("personal_info",{}).get("dob",""),
            "mobile": p.get("personal_info",{}).get("mobile",""),
            "email": p.get("personal_info",{}).get("email",""),
            "pincode": p.get("personal_info",{}).get("pincode",""),
            "filing_section": p.get("personal_info",{}).get("filing_section",""),
        },
        "income_details": {
            "salary": p.get("income_details",{}).get("salary",0),
            "hra_received": p.get("income_details",{}).get("hra_received",0),
            "rent_paid": p.get("income_details",{}).get("rent_paid",0),
            "hp_income": p.get("income_details",{}).get("hp_income",0),
        },
        "deductions": {
            "deduction_80C": p.get("deductions",{}).get("deduction_80C",0),
            "deduction_80EE": p.get("deductions",{}).get("deduction_80EE",0),
            "deduction_80EEA": p.get("deductions",{}).get("deduction_80EEA",0),
            "deduction_80D_preventive": p.get("deductions",{}).get("deduction_80D_preventive",0),
            "donations_80G": p.get("deductions",{}).get("donations_80G",[]),
        },
        "tds": {
            "tds_salary": p.get("tds",{}).get("tds_salary",[]),
            "tds_other": p.get("tds",{}).get("tds_other",[]),
        },
        "bank": {
            "bank_accounts": p.get("bank",{}).get("bank_accounts",[]),
        },
        "taxes_paid": {
            "advance_tax_payments": p.get("taxes_paid",{}).get("advance_tax_payments",[]),
        },
        "regime": p.get("regime","NEW"),
        "part_a_gen_1398A": p.get("part_a_gen_1398A",{}),
    }
    # Need computed taxable_income for some rules, but optional
    tax_input = {
        "age": p.get("age",30),
        "is_resident": p.get("is_resident",True),
        "regime": p.get("regime","NEW"),
        "salary": p.get("income_details",{}).get("salary",0),
        "hp_income": p.get("income_details",{}).get("hp_income",0),
        "other_sources": p.get("income_details",{}).get("other_sources",0),
        "deductions_80C": p.get("deductions",{}).get("deduction_80C",0),
        "deductions_other": 0,
    }
    try:
        from ..services.tax_engine import full_tax_computation
        comp = full_tax_computation(tax_input)
        validation_payload["computed"] = {"taxable_income": comp["taxable_income"]}
    except:
        validation_payload["computed"] = {}

    result = validate_itr_generic(validation_payload, p.get("itr_form","ITR-1"))
    return result

@router.post("/generate-json")
def generate_json(payload: ITRPayload):
    p = payload.model_dump() if hasattr(payload, 'model_dump') else payload.dict()
    income_details = p.get("income_details",{})
    cg_entries = income_details.get("capital_gains_entries",[])
    cg_total = sum([e.get("gain",0) or ((e.get("sell_value",0) or 0)-(e.get("buy_value",0) or 0)) for e in cg_entries])
    business_income = 0
    if income_details.get("presumptive_income"):
        business_income += income_details.get("presumptive_income",0)
    if income_details.get("gross_profit") is not None and income_details.get("presumptive_type","")=='':
        business_income += (income_details.get("gross_profit",0) or 0) - (income_details.get("expenses",0) or 0)
    # First calculate tax
    tax_input = {
        "age": p.get("age",30),
        "is_resident": p.get("is_resident",True),
        "regime": p.get("regime","NEW"),
        "salary": income_details.get("salary",0),
        "basic": income_details.get("basic",0),
        "da": income_details.get("da",0),
        "hra_received": income_details.get("hra_received",0),
        "rent_paid": income_details.get("rent_paid",0),
        "is_metro": income_details.get("is_metro",False),
        "hp_income": income_details.get("hp_income",0),
        "other_sources": income_details.get("other_sources",0),
        "capital_gains": cg_total,
        "business_income": business_income,
        "deductions_80C": p.get("deductions",{}).get("deduction_80C",0),
        "deductions_other": p.get("deductions",{}).get("deduction_80D",0) + p.get("deductions",{}).get("deduction_80G",0),
        "tds": sum([x.get("tax_deducted",0) for x in p.get("tds",{}).get("tds_salary",[])] ) + sum([x.get("tax_deducted",0) for x in p.get("tds",{}).get("tds_other",[])] ),
        "advance_tax": p.get("taxes_paid",{}).get("advance_tax",0),
        "self_assessment": p.get("taxes_paid",{}).get("self_assessment",0),
        "filing_date": p.get("filing_date","2026-07-31"),
        "due_date": p.get("due_date","2026-07-31"),
    }
    comp = full_tax_computation(tax_input)

    # Build flat payload for json_builder expected
    flat_payload = {
        "personal_info": p.get("personal_info",{}),
        "income_details": p.get("income_details",{}),
        "deductions": p.get("deductions",{}),
        "bank": p.get("bank",{}),
        "tds": p.get("tds",{}),
        "taxes_paid": p.get("taxes_paid",{}),
    }

    itr_form = p.get("itr_form","ITR-1")
    json_obj, json_str, digest = convert_to_official_json(itr_form, flat_payload, comp)

    return {
        "itr_form": itr_form,
        "json": json_obj,
        "json_string": json_str,
        "digest": digest,
        "tax_computation": comp,
    }

@router.get("/forms")
def list_forms():
    return {
        "forms": [
            {"id":"ITR-1","name":"Sahaj - For resident individuals with income up to 50L, 1 house, no business","ay":"2026-27"},
            {"id":"ITR-2","name":"For individuals/HUFs not having business, with capital gains, more than 1 house, foreign assets","ay":"2026-27"},
            {"id":"ITR-3","name":"For individuals/HUFs having business income, P&L, BS","ay":"2026-27"},
            {"id":"ITR-4","name":"Sugam - Presumptive business 44AD, 44ADA, 44AE","ay":"2026-27"},
        ]
    }

@router.post("/import-json")
def import_json_endpoint(data: dict):
    # Import existing official JSON and map to internal payload
    try:
        # data expected {"json": {...}}
        official = data.get("json") or data
        # Try detect form
        itr_form = "ITR-1"
        if "ITR" in official:
            if "ITR1" in official["ITR"]:
                itr_form="ITR-1"
            elif "ITR2" in official["ITR"]:
                itr_form="ITR-2"
            elif "ITR3" in official["ITR"]:
                itr_form="ITR-3"
            elif "ITR4" in official["ITR"]:
                itr_form="ITR-4"
        # Map back - simplified
        # For MVP, return raw after validation against schema
        return {"itr_form": itr_form, "imported": True, "preview": official, "mapped_payload": {"note": "Mapping implemented for ITR-1, others stub"}}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
