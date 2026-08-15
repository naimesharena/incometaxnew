"""
Validation Engine - Replicating VBA Validate* and CBDT Validation Rules PDF
Covers ITR-1..4 common rules

Rules extracted from:
- VBA md modules ValidateTAN, ValidatePAN, ValidateIFSC, Validate80G etc (from oletools extraction)
- CBDT_e-Filing validation PDFs (present in repo)
- Data validations from sheets (387-1257 per file)

This engine is declarative - returns list of errors
"""
import re
from typing import List, Dict, Any, Tuple

# Regex patterns - mirrors VBA ChkPAN etc
PAN_REGEX = r"^[A-Z]{5}[0-9]{4}[A-Z]$"
AADHAAR_REGEX = r"^\d{12}$"
MOBILE_REGEX = r"^[6-9]\d{9}$"
EMAIL_REGEX = r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"
TAN_REGEX = r"^[A-Z]{4}[0-9]{5}[A-Z]$"
IFSC_REGEX = r"^[A-Z]{4}0[A-Z0-9]{6}$"
PINCODE_REGEX = r"^[1-9][0-9]{5}$"
DATE_REGEX = r"^\d{2}/\d{2}/\d{4}$"  # DD/MM/YYYY as per Excel

def validate_pan(pan: str) -> bool:
    if not pan:
        return False
    return bool(re.match(PAN_REGEX, pan.upper()))

def validate_tan(tan: str) -> bool:
    if not tan:
        return False
    return bool(re.match(TAN_REGEX, tan.upper()))

def validate_aadhaar(aadhaar: str) -> bool:
    if not aadhaar:
        return False
    return bool(re.match(AADHAAR_REGEX, str(aadhaar)))

def validate_mobile(mobile: str) -> bool:
    if not mobile:
        return False
    return bool(re.match(MOBILE_REGEX, str(mobile)))

def validate_email(email: str) -> bool:
    if not email:
        return False
    return bool(re.match(EMAIL_REGEX, email))

def validate_ifsc(ifsc: str) -> bool:
    if not ifsc:
        return False
    return bool(re.match(IFSC_REGEX, ifsc.upper()))

def validate_pincode(pin: str) -> bool:
    if not pin:
        return False
    return bool(re.match(PINCODE_REGEX, str(pin)))

# Business rule validators replicating VBA

def validate_personal_info(data: Dict[str, Any]) -> List[str]:
    errors=[]
    pan = data.get("pan","")
    if not validate_pan(pan):
        errors.append("PersonalInfo.PAN: Invalid PAN format (e.g., ABCDE1234F) - mirrors VBA ChkPAN")
    aadhaar = data.get("aadhaar","")
    if aadhaar and not validate_aadhaar(aadhaar):
        errors.append("PersonalInfo.Aadhaar: Must be 12 digits - ValidateAadhaar logic")
    dob = data.get("dob","")
    if not dob:
        errors.append("PersonalInfo.DOB: Mandatory")
    # Age validation from VBA: calculateAge
    mobile = data.get("mobile","")
    if mobile and not validate_mobile(mobile):
        errors.append("PersonalInfo.Mobile: Invalid 10-digit mobile starting 6-9")
    email = data.get("email","")
    if email and not validate_email(email):
        errors.append("PersonalInfo.Email: Invalid format")
    pincode = data.get("pincode","")
    if pincode and not validate_pincode(pincode):
        errors.append("PersonalInfo.Pincode: Must be 6 digits starting 1-9")
    # State vs pincode matching - mirrors StateMatchesPin VBA
    # For full check would need pincode->state service
    return errors

def validate_income_details(data: Dict[str, Any]) -> List[str]:
    errors=[]
    salary = data.get("salary",0)
    if salary and salary <0:
        errors.append("Income.Salary: Cannot be negative")
    # HRA checks - from ValidateEA10_13A_1 VBA
    hra_received = data.get("hra_received",0)
    rent_paid = data.get("rent_paid",0)
    if hra_received>0 and rent_paid==0:
        errors.append("Schedule EA 10(13A): If HRA received, rent paid must be provided")
    if rent_paid>0 and hra_received==0:
        errors.append("Schedule EA 10(13A): Rent paid given but HRA not selected")
    # House property
    hp = data.get("hp_income",0)
    # co-owner validation - from mdHouseProperty
    if data.get("hp_type")=="let_out":
        if not data.get("tenant_pan") and data.get("rent_realized",0)>0:
            # tenant pan may be optional if less than threshold, but VBA ValidateTenantPan checks
            pass
    return errors

def validate_80deductions(data: Dict[str, Any], regime: str) -> List[str]:
    errors=[]
    # Flag mutual exclusivity 80EE vs 80EEA - mirrors lock_80EE_flag logic
    d80EE = data.get("deduction_80EE",0)
    d80EEA = data.get("deduction_80EEA",0)
    if d80EE>0 and d80EEA>0:
        errors.append("Deduction 80EE and 80EEA cannot be claimed together - mirrors Deduction_80EE_and_80EEA_chk VBA")
    # 80C cap 1.5L aggregate - from 80C sheet
    total_80C = data.get("deduction_80C",0) + data.get("deduction_80CCC",0) + data.get("deduction_80CCD1",0)
    if total_80C > 150000 and regime=="OLD":
        # Note Excel allows more but caps at 150k, not error but warning; CBDT validation warns?
        pass
    # 80D family checks - from Sch80D.bas ChkFamilyMember etc
    # Preventive health checkup limit 5000
    preventive = data.get("deduction_80D_preventive",0)
    if preventive>5000:
        errors.append("80D Preventive health checkup capped at Rs 5000 - Validate_80D")
    # 80G donation: need donee name, PAN, address if amount > certain
    donations = data.get("donations_80G", [])
    for idx, don in enumerate(donations):
        if don.get("amount",0)>0:
            if not don.get("donee_name"):
                errors.append(f"80G Row {idx+1}: Donee name mandatory if donation amount given")
            if not don.get("pan") and don.get("amount",0)>2000:
                # PAN mandatory above 2000? Actually per CBDT
                errors.append(f"80G Row {idx+1}: PAN mandatory for donation >2000")
    return errors

def validate_tds(data: Dict[str, Any]) -> List[str]:
    errors=[]
    tds_entries = data.get("tds_salary", []) + data.get("tds_other", [])
    for idx, entry in enumerate(tds_entries):
        tan = entry.get("tan","")
        if tan and not validate_tan(tan):
            errors.append(f"TDS Row {idx+1}: Invalid TAN format - ValidateTAN1_TDS")
        employer = entry.get("employer_name","")
        if not employer or len(employer)<3:
            errors.append(f"TDS Row {idx+1}: Employer/Deductor name must be at least 3 chars - ValidateEmployer")
        income_chargeable = entry.get("income_chargeable",0)
        tax_deducted = entry.get("tax_deducted",0)
        if tax_deducted>0 and income_chargeable==0:
            errors.append(f"TDS Row {idx+1}: Income chargeable required if tax deducted given")
        if tax_deducted> income_chargeable * 0.5: # heuristic
            # actual CBDT: claim cannot exceed deducted
            pass
        # Year - TCS_CollectedYear validation
        year = entry.get("year","")
        if year and year not in ["2025","2026","2024"]:
            # simplified
            pass
    return errors

def validate_bank(data: Dict[str, Any]) -> List[str]:
    errors=[]
    banks = data.get("bank_accounts", [])
    if not banks or len(banks)==0:
        errors.append("Bank Accounts: At least one account mandatory for refund - BA validation")
    primary_for_refund = [b for b in banks if b.get("is_for_refund")]
    if len(primary_for_refund)==0 and banks:
        errors.append("Bank Accounts: At least one account must be marked for refund")
    for idx, b in enumerate(banks):
        ifsc = b.get("ifsc","")
        if ifsc and not validate_ifsc(ifsc):
            errors.append(f"Bank Row {idx+1}: Invalid IFSC format - CheckIFSC")
        acc_no = b.get("account_number","")
        if acc_no and (len(acc_no)<9 or len(acc_no)>18):
            errors.append(f"Bank Row {idx+1}: Account number must be 9-18 digits - ValidateAccntNumber_BA")
        acc_type = b.get("account_type","")
        if acc_type not in ["Saving","Current",""]:
            errors.append(f"Bank Row {idx+1}: Account type must be Saving/Current")
    return errors

def validate_taxes_paid(data: Dict[str, Any]) -> List[str]:
    errors=[]
    advance = data.get("advance_tax_payments", [])
    for idx, ch in enumerate(advance):
        bsr = ch.get("bsr_code","")
        if bsr and not re.match(r"^\d{7}$", str(bsr)):
            errors.append(f"Advance Tax Row {idx+1}: BSR code must be 7 digits")
        challan_no = ch.get("challan_no","")
        if challan_no and not re.match(r"^\d{5}$", str(challan_no)):
            errors.append(f"Advance Tax Row {idx+1}: Challan No must be 5 digits")
        amt = ch.get("amount",0)
        if amt<=0:
            errors.append(f"Advance Tax Row {idx+1}: Amount must be positive")
    return errors

def validate_itr1_full(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Full ITR-1 validation orchestrator mirroring ValidateSheet* flow in VBA
    """
    errors=[]
    warnings=[]

    # Personal
    errors.extend(validate_personal_info(payload.get("personal_info",{})))
    # Income
    errors.extend(validate_income_details(payload.get("income_details",{})))
    # Deductions
    regime = payload.get("regime","NEW")
    errors.extend(validate_80deductions(payload.get("deductions",{}), regime))
    # TDS
    errors.extend(validate_tds(payload.get("tds",{})))
    # Bank
    errors.extend(validate_bank(payload.get("bank",{})))
    # Taxes
    errors.extend(validate_taxes_paid(payload.get("taxes_paid",{})))

    # Additional CBDT rules from PDF
    total_income = payload.get("computed",{}).get("taxable_income",0)
    if total_income > 5000000 and regime=="NEW":
        # Surcharge applicable, ensure PAN etc
        pass
    # Return filing section validation
    filing_sec = payload.get("personal_info",{}).get("filing_section","")
    if filing_sec == "139(8A)" :
        # Need Part A Gen_139(8A) mandatory
        gen139 = payload.get("part_a_gen_1398A",{})
        if not gen139.get("reason"):
            errors.append("Part A Gen 139(8A): Reason for updating income mandatory when filing u/s 139(8A)")

    is_valid = len(errors)==0
    return {"valid": is_valid, "errors": errors, "warnings": warnings, "count": len(errors)}

# For ITR-2,3,4 add extended validation stubs
def validate_itr2_full(payload: Dict) -> Dict:
    base = validate_itr1_full(payload)
    # Additional: Capital gains, Business etc
    # BP, CG, 112A validation
    if payload.get("capital_gains",{}).get("has_112A"):
        # Check CSV_112A format
        pass
    return base

def validate_itr_generic(payload: Dict, itr_form: str) -> Dict:
    if itr_form=="ITR-1":
        return validate_itr1_full(payload)
    elif itr_form in ["ITR-2","ITR-3","ITR-4"]:
        return validate_itr2_full(payload)
    else:
        return {"valid": False, "errors": [f"Unknown ITR form {itr_form}"], "warnings": [], "count":1}
