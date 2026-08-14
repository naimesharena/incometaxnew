"""CBDT JSON builder - output contract per ITR-1_2026_Main_V1.1.json schema.

Structure mirrors GenerateJson.bas (ITR1.vba): ITR -> ITR1 -> {CreationInfo,
Form_ITR1, PersonalInfo, FilingStatus, ITR1_IncomeDeductions,
ITR1_TaxComputation, Schedule80x..., TDSonSalaries, TDSonOthThanSals,
ScheduleTCS, TaxPayments, LTCG112A, Refund, Verification, TaxReturnPreparer}.

Every field name/format here is checked against the official draft-04 schema
(see validate_against_schema + tests/test_itr1_pipeline.py).
"""
from datetime import date, datetime
import base64, hashlib, json

from . import constants as C
from .itr1 import ITR1Return

# schema TAN jurisdiction prefixes [schema EmployerOrDeductorOrCollectDetl.TAN]
TAN_PREFIXES = ("HYD VPN BBN BPL JBP CHE CMB MRI DEL CAL MRT AHM BRD RKT SRT "
                "BLR AGR KNP CHN TVD ALD LKN MUM NGP AMR JLD PTL RTK KLP NSK "
                "PNE PTN RCH JDH JPR SHL").split()


def _iso(d):
    return d.strftime("%Y-%m-%d") if d else ""


def _empcat(code: str) -> str:
    return code if code in C.NATURE_OF_EMPLOYMENT else "OTH"


def _state_code(state) -> str:
    """Accepts '11', '11-Gujarat' or '(Select)' -> 2-digit schema code."""
    s = str(state or "").strip()
    if s in C.STATE_CODES:
        return s
    for k, v in C.STATE_CODES.items():
        if s.startswith(k) or v.lower() in s.lower():
            return k
    return ""


def build_itr1_json(ret: ITR1Return) -> dict:
    s = ret.compute_summary()
    p, f = ret.personal, ret.filing

    # ---------------- FilingStatus --------------------------------------
    filing_status = {
        "ReturnFileSec": f.filing_sec_code,
        "OptOutNewTaxRegime": "Y" if f.regime == C.REGIME_OLD else "N",
        "SeventhProvisio139": "Y" if _yes(f.seventh_provisio_139) else "N",
        "IncrExpAggAmt2LkTrvFrgnCntryFlg": "N",
        "IncrExpAggAmt1LkElctrctyPrYrFlg": "N",
        "clauseiv7provisio139i": "Y" if _yes(f.clause_iv_seventh_provisio) else "N",
        "AsseseeRepFlg": "Y" if _yes(f.representative.applicable) else "N",
        "ItrFilingDueDate": _iso(f.due_date),
    }
    if f.amount_seventh_provisio_ii > 0:
        filing_status["AmtSeventhProvisio139ii"] = int(f.amount_seventh_provisio_ii)
    if f.amount_seventh_provisio_iii > 0:
        filing_status["AmtSeventhProvisio139iii"] = int(f.amount_seventh_provisio_iii)
    if f.receipt_number:
        filing_status["ReceiptNo"] = f.receipt_number
        if f.original_return_filed_date:
            filing_status["OrigRetFiledDate"] = _iso(f.original_return_filed_date)
    if f.notice_unique_number:
        filing_status["NoticeNo"] = f.notice_unique_number
        if f.notice_date:
            filing_status["NoticeDateUnderSec"] = _iso(f.notice_date)
    if _yes(f.representative.applicable):
        filing_status["AssesseeRep"] = {
            "RepName": f.representative.name,
            "RepEmailID": f.representative.email,
            "CountryCodeRepMobileNo": 91,
            "RepMobileNo": int(re.sub(r"\D", "", f.representative.contact_no) or 0),
        }

    # ---------------- PersonalInfo --------------------------------------
    pa = p.primary_address
    personal_info = {
        "AssesseeName": {
            "FirstName": p.first_name,
            "MiddleName": p.middle_name,
            "SurNameOrOrgName": p.last_name,
        },
        "PAN": p.pan,
        "AadhaarCardNo": p.aadhaar_number,
        "DOB": _iso(p.date_of_birth),
        "EmployerCategory": _empcat(p.nature_of_employment),
        "Address": {
            "ResidenceNo": pa.flat_door_block or "-",
            "ResidenceName": pa.name_of_premises,
            "RoadOrStreet": pa.road_street_post_office,
            "LocalityOrArea": pa.area_locality or "-",
            "CityOrTownOrDistrict": getattr(pa, "city", "") or pa.area_locality or "-",
            "StateCode": _state_code(pa.state),
            "CountryCode": C.COUNTRY_INDIA,
            "PinCode": int(pa.pin_code) if pa.pin_code.isdigit() else 0,
            "CountryCodeMobile": int(C.COUNTRY_INDIA),
            "MobileNo": int(p.primary_mobile) if p.primary_mobile.isdigit() else 0,
            "EmailAddress": p.primary_email,
        },
        "SecondaryAdd": "Y" if _yes(p.secondary_address_same_as_primary) else "N",
    }

    # ---------------- Income & deductions --------------------------------
    sal = ret.salary
    exempt_items = []
    for nature, amt in sal.exempt_allowances:
        code = _sec10_code(nature)
        if code:
            exempt_items.append({"SalNatureDesc": code, "SalOthAmount": int(amt)})
    os_items = []
    for nature, desc, amt in ret.other_sources.items:
        os_items.append({"OthSrcNatureDesc": _os_code(nature),
                         "OthSrcOthNatOfInc": desc or None,
                         "OthSrcOthAmount": int(amt)})
    div = ret.other_sources.dividend_by_quarter
    if any(div):
        os_items.append({"OthSrcNatureDesc": "DIV",
                         "OthSrcOthAmount": int(sum(div)),
                         "DividendInc": {"FromDate": "2025-04-01", "ToDate": "2026-03-31"}})

    comp = s["ChapterVIAComponents"]

    def _via(name):
        return int(comp.get(name, 0))

    usr_via = {
        "Section80C": int(ret.chapter_via.s80c or 0),
        "Section80CCC": int(ret.chapter_via.s80ccc or 0),
        "Section80CCDEmployeeOrSE": int(ret.chapter_via.s80ccd1 or 0),
        "Section80CCD1B": int(ret.chapter_via.s80ccd1b or 0),
        "Section80CCDEmployer": int(ret.chapter_via.s80ccd2 or 0),
        "Section80D": int(ret.chapter_via.s80d_schedule_amount or 0),
        "Section80DD": int(ret.chapter_via.s80dd_schedule_amount or 0),
        "Section80DDB": int(ret.chapter_via.s80ddb_schedule_amount or 0),
        "Section80E": int(ret.chapter_via.s80e or 0),
        "Section80EE": int(ret.chapter_via.s80ee or 0),
        "Section80EEA": int(ret.chapter_via.s80eea or 0),
        "Section80EEB": int(ret.chapter_via.s80eeb or 0),
        "Section80G": int(ret.chapter_via.s80g_schedule_amount or 0),
        "Section80GG": int(ret.chapter_via.s80gg or 0),
        "Section80GGA": int(ret.chapter_via.s80gga_eligible_donation or 0),
        "Section80GGC": int(ret.chapter_via.s80ggc_eligible_donation or 0),
        "Section80U": int(ret.chapter_via.s80u_schedule_amount or 0),
        "Section80TTA": int(ret.chapter_via.s80tta or 0),
        "Section80TTB": int(ret.chapter_via.s80ttb or 0),
        "AnyOthSec80CCH": int(ret.chapter_via.s80cch_agnipath or 0),
        "TotalChapVIADeductions": int(ret.chapter_via.user_total or 0),
    }
    calc_via = {
        "Section80C": _via("80C"), "Section80CCC": _via("80CCC"),
        "Section80CCDEmployeeOrSE": _via("80CCD1"), "Section80CCD1B": _via("80CCD1B"),
        "Section80CCDEmployer": _via("80CCD2"), "Section80D": _via("80D"),
        "Section80DD": _via("80DD"), "Section80DDB": _via("80DDB"),
        "Section80E": _via("80E"), "Section80EE": _via("80EE"),
        "Section80EEA": _via("80EEA"), "Section80EEB": _via("80EEB"),
        "Section80G": _via("80G"), "Section80GG": _via("80GG"),
        "Section80GGA": _via("80GGA"), "Section80GGC": _via("80GGC"),
        "Section80U": _via("80U"), "Section80TTA": _via("80TTA"),
        "Section80TTB": _via("80TTB"), "AnyOthSec80CCH": _via("80CCH"),
        "TotalChapVIADeductions": s["TotalChapterVIA"],
    }

    income = {
        "GrossSalary": sal.gross_salary,
        "Salary": int(sal.salary_17_1),
        "PerquisitesValue": int(sal.perquisites_17_2),
        "ProfitsInSalary": int(sal.profits_in_lieu_17_3),
        "AllwncExemptUs10": {
            "AllwncExemptUs10Dtls": exempt_items,
            "TotalAllwncExemptUs10": int(sal.exempt_u10_total(ret.regime)),
        },
        "NetSalary": sal.net_salary,
        "DeductionUs16ia": sal.deduction_16ia(ret.regime),
        "EntertainmentAlw16ii": int(sal.entertainment_allowance_16_ii),
        "ProfessionalTaxUs16iii": int(sal.professional_tax_16_iii),
        "DeductionUs16": int(sal.deduction_16ia(ret.regime)
                             + sal.entertainment_allowance_16_ii
                             + sal.professional_tax_16_iii),
        "IncomeFromSal": ret.income_from_salaries(),
        "PropertyDetails": [_property_json(i, hp) for i, hp in enumerate(ret.house_properties, 1)],
        "TotalIncomeChargeableUnHP": ret.hp_income(),
        "IncomeOthSrc": ret.other_sources.income_from_os(),
        "OthersInc": {"OthersIncDtlsOthSrc": os_items},
        "DeductionUs57iia": int(ret.other_sources.deduction_57_iia),
        "GrossTotIncome": s["GrossTotIncome"],
        "GrossTotIncomeIncLTCG112A": s["GrossTotIncomeIncLTCG112A"],
        "UsrDeductUndChapVIA": usr_via,
        "DeductUndChapVIA": calc_via,
        "TotalIncome": s["TotalIncome"],
        "ExemptIncAgriOthUs10": {"ExemptIncAgriOthUs10Total": 0},
    }

    tax_computation = {
        "TotalTaxPayable": s["TotalTaxPayable"],
        "Rebate87A": s["Rebate87A"],
        "TaxPayableOnRebate": s["TaxPayableOnRebate"],
        "EducationCess": s["EducationCess"],
        "GrossTaxLiability": s["GrossTaxLiability"],
        "Section89": s["Section89"],
        "NetTaxLiability": s["NetTaxLiability"],
        "IntrstPay": {
            "IntrstPayUs234A": s["IntrstPay234A"],
            "IntrstPayUs234B": s["IntrstPay234B"],
            "IntrstPayUs234C": s["IntrstPay234C"],
            "LateFilingFee234F": s["FeeIncUS234F"],
        },
        "TotalIntrstPay": s["TotalIntrstPay"],
        "TotTaxPlusIntrstPay": s["TotTaxPlusIntrstPay"],
    }

    # ---------------- taxes paid / refund / verification ----------------
    tds_sal = {
        "TotalTDSonSalaries": int(sum(t.tds_amount for t in ret.taxes_paid.tds_salary)),
        "TDSonSalary": [
            {"EmployerOrDeductorOrCollectDetl": {
                "TAN": t.tan, "EmployerOrDeductorOrCollecterName": t.employer_name},
             "IncChrgSal": 0, "TotalTDSSal": int(t.tds_amount)}
            for t in ret.taxes_paid.tds_salary
        ],
    }
    tds_other = {
        "TotalTDSonOthThanSals": int(sum(t.tds_amount for t in ret.taxes_paid.tds_other)),
        "TDSonOthThanSal": [
            {"EmployerOrDeductorOrCollectDetl": {
                "TAN": t.tan, "EmployerOrDeductorOrCollecterName": t.deductor_name},
             "SectionCode": t.section, "TotalTDSOnOthThanSal": int(t.tds_amount)}
            for t in ret.taxes_paid.tds_other
        ],
    }
    tcs = {
        "TotalSchTCS": int(sum(t.tcs_amount for t in ret.taxes_paid.tcs)),
        "TCS": [
            {"EmployerOrDeductorOrCollectDetl": {
                "TAN": t.tan, "EmployerOrDeductorOrCollecterName": t.collector_name},
             "AmtTaxCollected": int(t.tcs_amount), "CollectedYr": "2025",
             "TotalTCS": int(t.tcs_amount), "AmtTCSClaimedThisYear": int(t.tcs_amount)}
            for t in ret.taxes_paid.tcs
        ],
    }
    tax_payments = {
        "TotalTaxPayments": int(ret.taxes_paid.total_it),
        "TaxPayment": [
            {"BSRCode": c.bsr_code, "DateDep": _iso(c.date_of_deposit),
             "SrlNoOfChaln": int(c.challan_serial_no or 0), "Amt": int(c.amount)}
            for c in ret.taxes_paid.it_challans
        ],
    }
    tp = ret.taxes_paid
    tax_paid = {
        "TaxesPaid": {
            "AdvanceTax": int(tp.advance_tax),
            "TDS": int(tp.total_tds),
            "TCS": int(tp.total_tcs),
            "SelfAssessmentTax": int(tp.self_assessment_tax),
            "TotalTaxesPaid": int(ret.total_taxes_paid()),
        },
        "BalTaxPayable": int(max(0, -s["RefundOrPayable"])),
    }
    refund_due = int(max(0, s["RefundOrPayable"]))
    if not ret.bank.ifsc:
        raise ValueError("Bank account (IFSC) is mandatory - the official schema "
                         "requires at least one bank account for refund credit")
    refund = {
        "RefundDue": refund_due,
        "BankAccountDtls": {
            "AddtnlBankDetails": [{
                "IFSCCode": ret.bank.ifsc,
                "BankAccountNo": ret.bank.account_no,
                "BankName": ret.bank.bank_name,
                "AccountType": "SB",
                "UseForRefund": "true" if refund_due > 0 else "false",
            }],
        },
    }
    verification = {
        "Declaration": {
            "AssesseeVerName": f"{p.first_name} {p.last_name}".strip(),
            "FatherName": getattr(p, "father_name", "") or "NA",
            "AssesseeVerPAN": p.pan,
        },
        "Capacity": "S" if ret.verification.capacity.lower().startswith("self") else "R",
        "Place": pa.area_locality or pa.city if hasattr(pa, "city") else pa.area_locality or "-",
    }

    itr1 = {
        "Form_ITR1": {
            "FormName": C.FORM_NAME_ITR1,
            "Description": "ITR-1 SAHAJ for Resident individuals - AY 2026-27",
            "AssessmentYear": C.ASSESSMENT_YEAR,
            "SchemaVer": "Ver1.0",
            "FormVer": "Ver1.0",
        },
        "FilingStatus": filing_status,
        "PersonalInfo": personal_info,
        "ITR1_IncomeDeductions": income,
        "ITR1_TaxComputation": tax_computation,
        "TaxPaid": tax_paid,
        "LTCG112A": {
            "TotSaleCnsdrn": int(ret.ltcg_112a.total_sale_consideration),
            "TotCstAcqisn": int(ret.ltcg_112a.total_cost_of_acquisition),
            "LongCap112A": int(ret.ltcg_112a.gain),
        },
        "Refund": refund,
        "Verification": verification,
    }
    # schemas require minItems=1 for these arrays -> omit sections when empty
    if tds_sal["TDSonSalary"]:
        itr1["TDSonSalaries"] = tds_sal
    if tds_other["TDSonOthThanSal"]:
        itr1["TDSonOthThanSals"] = tds_other
    if tcs["TCS"]:
        itr1["ScheduleTCS"] = tcs
    if tax_payments["TaxPayment"]:
        itr1["TaxPayments"] = tax_payments
    itr1["CreationInfo"] = _creation_info(itr1)
    return {"ITR": {"ITR1": itr1}}


def _yes(v) -> bool:
    return str(v).replace("(Select)", "").strip().lower().startswith("yes")


def _property_json(i, hp):
    addr = {
        "AddrDetail": getattr(hp, "address", "") or "-",
        "CityOrTownOrDistrict": getattr(hp, "city", "") or "-",
        "StateCode": _state_code(getattr(hp, "state", "")) or "11",
        "CountryCode": C.COUNTRY_INDIA,
    }
    if int(getattr(hp, "pincode", 0) or 0) >= 100000:
        addr["PinCode"] = int(hp.pincode)
    unrealised_plus_tax = int(hp.unrealised_rent + hp.tax_paid_local_authorities)
    return {
        "HPSNo": i,
        "AddressDetailWithZipCode": addr,
        "PropertyOwner": "SE",
        "PropCoOwnedFlg": "NO",
        "AsseseeShareProperty": float(hp.ownership_share_pct),
        "ifLetOut": "S" if hp.is_self_occupied else ("L" if hp.property_type == "Let Out" else "D"),
        "Rentdetails": {
            "AnnualLetableValue": int(hp.gross_annual_value),
            "RentNotRealized": int(hp.unrealised_rent),
            "LocalTaxes": int(hp.tax_paid_local_authorities),
            "TotalUnrealizedAndTax": unrealised_plus_tax,
            "BalanceALV": hp.annual_value(),
            "AnnualOfPropOwned": hp.share_of_annual_value(),
            "ThirtyPercentOfBalance": hp.standard_deduction_30pct(),
            "IntOnBorwCap": int(hp.interest_24b),
            "TotalDeduct": int(hp.standard_deduction_30pct() + hp.interest_24b),
            "ArrearsUnrealizedRentRcvd": int(hp.arrears_unrealised_received),
            "IncomeOfHP": hp.income(),
        },
    }


_SEC10_CODES = {
    "10(5)": "10(5)", "10(10)": "10(10)", "10(10A)": "10(10A)",
    "10(10AA)": "10(10AA)", "10(13A)": "10(10AA)",  # HRA reported via Schedule EA
}


def _sec10_code(nature: str):
    for k, v in _SEC10_CODES.items():
        if k in nature:
            return v
    return None


def _os_code(nature: str) -> str:
    n = nature.lower()
    if "saving" in n:
        return "SAV"
    if "deposit" in n or "interest" in n:
        return "IFD"
    if "dividend" in n:
        return "DIV"
    if "commission" in n or "any other" in n:
        return "Oth"
    return "IFD"


def _creation_info(form: dict) -> dict:
    body = json.dumps(form, sort_keys=True)
    digest = base64.b64encode(hashlib.sha256(body.encode()).digest()).decode()
    return {
        "SWVersionNo": "V" + C.FORM_VER.replace(".", "-"),
        "SWCreatedBy": "SW00000001",
        "JSONCreatedBy": "SW00000001",
        "JSONCreationDate": datetime.now().strftime("%Y-%m-%d"),
        "IntermediaryCity": "Surat",
        "Digest": digest,
    }


def validate_against_schema(payload: dict, schema_path: str) -> list:
    """Validate generated JSON against the official CBDT JSON schema (draft-04)."""
    try:
        import jsonschema
    except ImportError:
        return ["jsonschema library not installed"]
    schema = json.load(open(schema_path))
    v = jsonschema.Draft4Validator(schema)
    return [f"{'/'.join(map(str, e.path))}: {e.message}" for e in v.iter_errors(payload)]
