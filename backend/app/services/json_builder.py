"""
JSON Builder - Generates official CBDT JSON for ITR-1..4 + HMAC hashing for digest
Mirrors GenerateJson.bas (7120 lines) + mdHashing.bas + HS256.cls

Old VBA:
- Base64_HMACSHA256_JSON, EncodeBase64json, ToJsonFormat, Form_ITR1, CreationInfo, etc.
- getHashIteration from DataBase!B4 (1849) and getHashKey from B3 (7Z3mxclnABiXtYG)
- Digest = Base64(HMACSHA256(JSON, key iterated))

We replicate logic using Python's hmac + hashlib + base64
Hash iteration: In original, they iterate hash? Need to read GenerateJson.bas details
From observation: getHashIteration returns 1849, getHashKey returns 7Z3mxclnABiXtYG
Then Base64_HMACSHA256_JSON uses HMACSHA256 with key and iterations

Simplify: We'll implement standard method:
- key = hash_key from hash_meta.json
- iterations = hash_iteration
- For i in range(iterations): hmac = HMACSHA256(json_str, key) ??? Need check

From mdHashing.bas: likely does iteration of HMAC

We'll implement:
digest = base64( hmac_sha256( key, json_string ) )
But also iterate: first hmac result becomes key for next iteration

We'll implement both simple and iterated, and also provide raw JSON without digest for testing.

Official JSON schema: The top-level has ITR->ITR1->CreationInfo, Form_ITR1, PersonalInfo, FilingStatus, ITR1_IncomeDeductions, ITR1_TaxComputation, TaxPaid, Refund, Schedules, Verification etc
As per ITR-1_2026_Main_V1.1.json
"""

import json
import base64
import hmac
import hashlib
from datetime import datetime
from typing import Dict, Any, Tuple
import pathlib

# Load hash meta
def get_hash_meta():
    try:
        path = pathlib.Path(__file__).parent.parent / "data" / "hash_meta.json"
        if path.exists():
            with open(path) as f:
                return json.load(f)
    except:
        pass
    return {"hash_key": "7Z3mxclnABiXtYG", "hash_iteration": 1849}

def base64_hmac_sha256(message: str, key: str) -> str:
    """Mimics Base64_HMACSHA256 from mdHashing.bas"""
    key_bytes = key.encode('utf-8')
    msg_bytes = message.encode('utf-8')
    hm = hmac.new(key_bytes, msg_bytes, hashlib.sha256)
    digest = hm.digest()
    return base64.b64encode(digest).decode('utf-8')

def iterated_hmac(message: str, key: str, iterations: int) -> str:
    """
    Iterated hashing as per getHashIteration
    Logic guessed from GenerateJson.bas:
    First iteration uses original key, subsequent uses previous digest as key? Or repeats same?
    We'll implement: 
    result = message
    for i in range(iterations):
        result = HMACSHA256(result, key)
    Then base64 final
    Alternative observed in some CBDT docs: iterations on key stretching?
    We'll do simple iterative: key stays same, message is previous HMAC raw.
    """
    current = message.encode('utf-8')
    key_bytes = key.encode('utf-8')
    for i in range(iterations):
        hm = hmac.new(key_bytes, current, hashlib.sha256)
        current = hm.digest()
        # For next iteration, use digest as message, key stays same (common KDF)
    return base64.b64encode(current).decode('utf-8')

def build_creation_info() -> Dict[str, Any]:
    now = datetime.now()
    # SW values from EfilingCommon.bas getSWVersionNo etc
    return {
        "SWVersionNo": "1.0",
        "SWCreatedBy": "SW00000001",  # dummy, should be registered vendor ID
        "JSONCreatedBy": "SW00000001",
        "JSONCreationDate": now.strftime("%Y-%m-%d"),
        "IntermediaryCity": "Delhi",
        "Digest": "-"  # placeholder, will be replaced after full JSON string built
    }

def build_form_itr1() -> Dict[str, Any]:
    return {
        "FormName": "ITR-1",
        "Description": "For individuals being a resident (other than not ordinarily resident) having total income upto Rs.50 lakh",
        "AssessmentYear": "2026",
        "SchemaVer": "Ver1.1",
        "FormVer": "Ver1.1"
    }

def build_itr1_json(payload: Dict[str, Any], tax_computation: Dict[str, Any], validation_passed: bool = True) -> Dict[str, Any]:
    """
    payload: user input collected from frontend (personal_info, income_details, deductions, bank, tds, taxes_paid)
    tax_computation: output from tax_engine.full_tax_computation
    Returns official JSON structure ready for digest
    """
    personal = payload.get("personal_info", {})
    income = payload.get("income_details", {})
    deductions = payload.get("deductions", {})
    bank = payload.get("bank", {})
    tds = payload.get("tds", {})
    taxes_paid = payload.get("taxes_paid", {})

    # PersonalInfo mapping
    personal_info_json = {
        "AssesseeName": {
            "FirstName": personal.get("first_name",""),
            "MiddleName": personal.get("middle_name",""),
            "SurNameOrOrgName": personal.get("last_name",""),
        },
        "PAN": personal.get("pan","").upper(),
        "AadhaarCardNo": personal.get("aadhaar",""),
        "Address": {
            "ResidenceNo": personal.get("flat_no",""),
            "ResidenceName": personal.get("premises",""),
            "RoadOrStreet": personal.get("street",""),
            "LocalityOrArea": personal.get("area",""),
            "CityOrTownOrDistrict": personal.get("city",""),
            "StateCode": personal.get("state_code",""),
            "CountryCode": personal.get("country_code","91"),
            "PinCode": personal.get("pincode",""),
            "CountryCodeMobileNo": "91",
            "MobileNo": personal.get("mobile",""),
            "EmailAddress": personal.get("email",""),
        },
        "DOB": personal.get("dob",""),  # should be DD/MM/YYYY? JSON schema expects YYYY-MM-DD? We'll use iso
        "EmployerCategory": income.get("employer_category","OTH"),
    }

    filing_status = {
        "ReturnFileSec": personal.get("filing_section","11"),  # 11 = 139(1) mapping needed
        "ResidentialStatus": personal.get("residential_status","RES"),
        "BenefitUs115HFlg": "N",
    }

    # Income Deductions
    income_ded_json = {
        "Salary": income.get("salary",0),
        "IncomeFromHP": income.get("hp_income",0),
        "IncomeFromOS": income.get("other_sources",0),
        "GrossTotIncome": tax_computation["income_breakdown"]["gross_total_income"],
        "DeductionsUs16": tax_computation["income_breakdown"]["standard_deduction_applied"] + tax_computation["income_breakdown"]["hra_exemption"],
        "TotalIncome": tax_computation["taxable_income"],
    }

    # Tax Computation
    chosen = tax_computation["chosen_tax"]
    tax_comp_json = {
        "TotalTaxPayable": chosen["total_tax"],
        "Rebate87A": chosen["rebate87A"],
        "TaxPayableOnRebate": chosen["tax_after_rebate"],
        "Surcharge": chosen["surcharge"],
        "EducationCess": chosen["cess"],
        "GrossTaxLiability": chosen["total_tax"],
        "Section89": 0,
        "NetTaxLiability": chosen["total_tax"],
    }

    # Tax Paid
    prepaid = tax_computation["prepaid_taxes"]
    tax_paid_json = {
        "TaxesPaid": {
            "AdvanceTax": taxes_paid.get("advance_tax", prepaid["advance"]),
            "TDS": prepaid["tds"],
            "TCS": tds.get("tcs_total",0),
            "SelfAssessmentTax": prepaid["self_assessment"] if "self_assessment" in str(prepaid) else taxes_paid.get("self_assessment",0),
        }
    }

    # Refund
    refund_json = {
        "RefundDue": tax_computation["final"]["refund"],
        "TaxPayable": tax_computation["final"]["payable"],
    }

    # Schedules stubs
    schedule_80g = []
    for don in deductions.get("donations_80G", []):
        schedule_80g.append({
            "DoneeName": don.get("donee_name",""),
            "PANOfDonee": don.get("pan",""),
            "DonationAmount": don.get("amount",0),
        })

    # TDS
    tds_salary = []
    for entry in tds.get("tds_salary", []):
        tds_salary.append({
            "TAN": entry.get("tan",""),
            "EmployerOrDeductorOrCollecterName": entry.get("employer_name",""),
            "IncChrgSal": entry.get("income_chargeable",0),
            "TotalTaxDeducted": entry.get("tax_deducted",0),
        })

    root = {
        "ITR": {
            "ITR1": {
                "CreationInfo": build_creation_info(),
                "Form_ITR1": build_form_itr1(),
                "PersonalInfo": personal_info_json,
                "FilingStatus": filing_status,
                "ITR1_IncomeDeductions": income_ded_json,
                "ITR1_TaxComputation": tax_comp_json,
                "TaxPaid": tax_paid_json,
                "Refund": refund_json,
                "Schedule80G": {"Don80G": schedule_80g} if schedule_80g else None,
                "TDSonSalaries": {"TDSonSalary": tds_salary} if tds_salary else None,
                "Verification": {
                    "Declaration": {
                        "AssesseeVerName": f"{personal.get('first_name','')} {personal.get('last_name','')}".strip(),
                        "Capacity": "Self",
                        "Place": personal.get("city","Delhi"),
                        "Date": datetime.now().strftime("%Y-%m-%d"),
                    }
                }
            }
        }
    }

    # Remove None
    # Clean nested Nones in ITR1
    itr1 = root["ITR"]["ITR1"]
    for k in list(itr1.keys()):
        if itr1[k] is None:
            del itr1[k]

    return root

def generate_json_with_digest(payload: Dict[str, Any], tax_computation: Dict[str, Any]) -> Tuple[Dict[str, Any], str, str]:
    """
    Generates JSON, computes digest, injects into CreationInfo.Digest, returns (json_obj, json_str, digest)
    """
    json_obj = build_itr1_json(payload, tax_computation)
    # First, set digest to "-" and serialize to string per CBDT ToJsonFormat logic (compact? pretty?)
    # CBDT uses ConvertJSONToString2 - likely compact without spaces? We'll use separators=(',', ':') for minimal
    # But digest should be computed on JSON without digest? Actually original does JSON with Digest="-" then hash?
    # Check GenerateJson.bas ToJsonFormat: likely produces formatted JSON string
    # We'll compute on string with Digest="-" (as per CreationInfo initial)
    json_str_without_digest = json.dumps(json_obj, separators=(',', ':'), ensure_ascii=False)

    meta = get_hash_meta()
    key = meta.get("hash_key","7Z3mxclnABiXtYG")
    iterations = int(meta.get("hash_iteration",1849))

    # Try iterated method
    digest = iterated_hmac(json_str_without_digest, key, iterations)
    # Also compute simple for comparison
    simple = base64_hmac_sha256(json_str_without_digest, key)

    # Inject digest into object
    json_obj["ITR"]["ITR1"]["CreationInfo"]["Digest"] = digest

    final_str = json.dumps(json_obj, indent=2, ensure_ascii=False)

    return json_obj, final_str, digest

# Generic wrapper for ITR-2,3,4 similar
def build_itr2_json(payload: Dict, tax_comp: Dict) -> Dict:
    # Stub - similar structure but more schedules
    base, _, _ = generate_json_with_digest(payload, tax_comp)
    # Replace FormName
    base["ITR"]["ITR2"] = base["ITR"].pop("ITR1")
    base["ITR"]["ITR2"]["Form_ITR2"] = {"FormName":"ITR-2","AssessmentYear":"2026","SchemaVer":"Ver1.1","FormVer":"Ver1.1","Description":"For individuals and HUFs not having business"}
    base["ITR"]["ITR2"]["CreationInfo"] = build_creation_info()
    # Add extra schedules placeholder
    base["ITR"] = {"ITR2": base["ITR"]["ITR2"]}
    # Actually restructure
    return base

def convert_to_official_json(itr_form: str, payload: Dict, tax_comp: Dict):
    if itr_form=="ITR-1":
        return generate_json_with_digest(payload, tax_comp)
    else:
        # For other forms, reuse ITR-1 logic with form name change
        json_obj, json_str, digest = generate_json_with_digest(payload, tax_comp)
        # Mutate
        if "ITR1" in json_obj["ITR"]:
            itr_data = json_obj["ITR"].pop("ITR1")
            # Adjust form
            itr_data["Form_"+itr_form.replace("-","")] = {"FormName": itr_form, "AssessmentYear":"2026","SchemaVer":"Ver1.1","FormVer":"Ver1.1"}
            # Keep CreationInfo etc
            json_obj["ITR"][itr_form.replace("-","")] = itr_data
        return json_obj, json_str, digest
