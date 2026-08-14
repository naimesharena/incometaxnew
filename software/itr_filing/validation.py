"""Validation rules - Category A (blocking) ported from:
  - CBDT_e-Filing_ITR 1_Validation Rules_AY 2026-27.pdf
  - VBA validation message blocks (ITR1.vba, e.g. msgError_LTCG ~line 19671)
  - Data-validation constraints (registry/ITR1.input_fields.csv)
Each rule is tagged with its source for auditability.
"""
import re
from . import constants as C
from .itr1 import ITR1Return


def validate_itr1(ret: ITR1Return) -> dict:
    """Returns {'errors': [blocking], 'warnings': [non-blocking]}."""
    errors, warnings = [], []
    p, f = ret.personal, ret.filing

    # --- identity ---
    if p.first_name and not re.fullmatch(r"[A-Za-z .]{1,75}", p.first_name):
        errors.append("[DV:Y7 max75] First name: only letters, max 75 chars")
    if p.aadhaar_number and not re.fullmatch(r"\d{12}", p.aadhaar_number):
        errors.append("[DV:AN8 textLength 12] Aadhaar must be exactly 12 digits")
    if p.primary_address.pin_code and p.primary_address.country == C.COUNTRY_DEFAULT:
        if not re.fullmatch(r"\d{6}", p.primary_address.pin_code):
            errors.append("[DV:AB22 textLength 6] Indian PIN code must be 6 digits")
    if p.primary_mobile and not re.fullmatch(r"[6-9]\d{9}", p.primary_mobile):
        errors.append("[DV] Mobile number must be 10 digits starting 6-9")
    if p.date_of_birth is None:
        errors.append("[PartA] Date of birth is mandatory")
    if p.nature_of_employment == "(Select)":
        errors.append("[DV:Z11 EmpCatList] Nature of employment must be selected")

    # --- eligibility (ITR-1 is residents only) ---
    # [Form description] RESIDENT (OTHER THAN NOT ORDINARILY RESIDENT)
    warnings.append("[Eligibility] ITR-1 is only for Resident (Not Ordinarily Resident excluded)")

    # --- 112A eligibility ---  [VBA msgError_LTCG line 19671]
    if ret.ltcg_112a.gain > C.LTCG_112A_MAX_ITR1:
        errors.append("[VBA msgError_LTCG] In ITR 1, the maximum gains as per Section 112A "
                      "can be INR 1,25,000/-. Please file ITR 2/3 if gains u/s 112A is more "
                      "than INR 1,25,000/-")

    # --- filing section ---
    if f.filing_sec_code in (17, 18) and not f.receipt_number:
        errors.append("[Category A] Revised/defective return requires original receipt number")
    if f.filing_sec_code in (13, 14, 15, 16) and not f.notice_unique_number:
        errors.append("[Category A] Notice-driven return requires notice unique number")

    # --- agricultural income eligibility ---  [VBA ~662/12345 msg]
    agri = sum(a for n, _, a in ret.other_sources.items if "agricultural" in n.lower())
    if agri > 5000:
        errors.append("[VBA eligibility] If sec 10(1) agricultural income is more than "
                      "Rs 5000/- then use ITR 2 or 3")

    # --- house property ---
    if len(ret.house_properties) > 2:
        errors.append("[Category A] ITR-1 allows at most 2 house properties")
    for hp in ret.house_properties:
        if hp.property_type not in ("Self Occupied", "Let Out", "Deemed Let Out"):
            errors.append("[DV:AO77] House property type must be Self Occupied / Let Out / Deemed Let Out")

    # --- standard deduction sanity ---
    if ret.salary.deduction_16ia(ret.regime) > (C.STD_DEDUCTION_NEW if ret.regime == C.REGIME_NEW
                                                else C.STD_DEDUCTION_OLD):
        errors.append("[AO73] Standard deduction exceeds cap")

    # --- taxes paid ---
    from .json_builder import TAN_PREFIXES
    tan_re = re.compile(r"(?:" + "|".join(TAN_PREFIXES) + r")[A-Z]\d{5}[A-Z]")
    for t in ret.taxes_paid.tds_salary:
        if t.tan and not tan_re.fullmatch(t.tan):
            errors.append("[schema TAN] TAN must start with a valid jurisdiction code "
                          "(e.g. MUMC12345A)")

    # --- refund ---
    if ret.refund_or_payable() > 0 and not ret.bank.ifsc:
        warnings.append("[Refund] Refund due but no bank account selected for credit")
    if ret.bank.ifsc and not re.fullmatch(r"[A-Z]{4}0[A-Z0-9]{6}", ret.bank.ifsc):
        errors.append("[Category A] IFSC format invalid (e.g. SBIN0001234)")

    return {"errors": errors, "warnings": warnings}
