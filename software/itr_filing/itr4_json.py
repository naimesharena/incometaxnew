"""CBDT JSON builder for ITR-4 (SUGAM) - contract per ITR-4_2026_Main_V1.1.json.

Sections per schema: CreationInfo, Form_ITR4, FilingStatus (incl. Form 10IEA
regime flags), PersonalInfo (Status I/H/F), IncomeDeductions (incl.
IncomeFromBusinessProf), TaxComputation, ScheduleBP (presumptive details),
TaxPaid, Refund, Verification (+ optional schedules).
"""
from datetime import datetime
import base64, hashlib, json

from . import constants as C
from .itr4 import ITR4Return, ITR4_DUE_DATE
from .json_builder import (_iso, _empcat, _state_code, _property_json,
                           validate_against_schema, TAN_PREFIXES)  # noqa: F401

VIA_KEYS = ["Section80C", "Section80CCC", "Section80CCDEmployeeOrSE",
            "Section80CCD1B", "Section80CCDEmployer", "Section80D",
            "Section80DD", "Section80DDB", "Section80E", "Section80G",
            "Section80GG", "Section80GGC", "Section80U", "Section80TTA",
            "Section80TTB", "AnyOthSec80CCH", "TotalChapVIADeductions"]


def build_itr4_json(ret: ITR4Return) -> dict:
    s = ret.compute_summary()
    p, f = ret.personal, ret.filing_sec_code

    filing_status = {
        "ReturnFileSec": f,
        "Form10IEAEarlierAYOldRegime": "NA",
        "F10IEACurrAYNewRegime": "N" if ret.regime == C.REGIME_OLD else "Y",
        "SeventhProvisio139": "N",
        "DepAmtAggAmtExcd1CrPrYrFlg": "N",
        "IncrExpAggAmt2LkTrvFrgnCntryFlg": "N",
        "IncrExpAggAmt1LkElctrctyPrYrFlg": "N",
        "clauseiv7provisio139i": "N",
        "AsseseeRepFlg": "N",
        "ItrFilingDueDate": _iso(ITR4_DUE_DATE),
    }

    pa = p.primary_address
    personal_info = {
        "AssesseeName": {"FirstName": p.first_name, "MiddleName": p.middle_name,
                         "SurNameOrOrgName": p.last_name},
        "PAN": p.pan,
        "AadhaarCardNo": p.aadhaar_number,
        "DOB": _iso(p.date_of_birth or ret.dob),
        "EmployerCategory": _empcat(p.nature_of_employment),
        "Status": ret.entity,
        "Address": {
            "ResidenceNo": pa.flat_door_block or "-",
            "ResidenceName": pa.name_of_premises,
            "RoadOrStreet": pa.road_street_post_office,
            "LocalityOrArea": pa.area_locality or "-",
            "CityOrTownOrDistrict": pa.city or pa.area_locality or "-",
            "StateCode": _state_code(pa.state),
            "CountryCode": C.COUNTRY_INDIA,
            "PinCode": int(pa.pin_code) if pa.pin_code.isdigit() else 0,
            "CountryCodeMobile": int(C.COUNTRY_INDIA),
            "MobileNo": int(p.primary_mobile) if p.primary_mobile.isdigit() else 0,
            "EmailAddress": p.primary_email,
        },
        "SecondaryAdd": "N",
    }

    usr_via = {k: int(ret.chapter_via_breakup.get(k, 0)) for k in VIA_KEYS}
    calc_via = {k: 0 for k in VIA_KEYS}
    calc_via["TotalChapVIADeductions"] = int(ret.chapter_via_total)

    income = {
        "IncomeFromBusinessProf": s["PresumptiveIncome"],
        "GrossSalary": int(ret.salary),
        "NetSalary": int(ret.salary),
        "DeductionUs16": 0,
        "IncomeFromSal": int(ret.salary),
        "PropertyDetails": [_property_json(i, hp)
                            for i, hp in enumerate(ret.house_properties, 1)],
        "TotalIncomeChargeableUnHP": ret.house_properties and
            int(sum(hp.income() for hp in ret.house_properties)) or 0,
        "IncomeOthSrc": int(max(0, ret.other_sources)),
        "GrossTotIncome": s["GrossTotIncome"],
        "GrossTotIncomeIncLTCG112A": s["GrossTotIncome"],
        "UsrDeductUndChapVIA": usr_via,
        "DeductUndChapVIA": calc_via,
        "TotalIncome": s["TotalIncome"],
    }

    tax_computation = {
        "TotalTaxPayable": s["TotalTaxPayable"],
        "Rebate87A": s["Rebate87A"],
        "TaxPayableOnRebate": max(0, s["TotalTaxPayable"] - s["Rebate87A"]),
        "EducationCess": s["EducationCess"],
        "GrossTaxLiability": s["GrossTaxLiability"],
        "NetTaxLiability": s["NetTaxLiability"],
        "IntrstPay": {
            "IntrstPayUs234A": s["IntrstPay234A"],
            "IntrstPayUs234B": s["IntrstPay234B"],
            "IntrstPayUs234C": 0,
            "LateFilingFee234F": s["FeeIncUS234F"],
        },
        "TotTaxPlusIntrstPay": int(s["NetTaxLiability"] + s["IntrstPay234A"]
                                   + s["IntrstPay234B"] + s["FeeIncUS234F"]),
    }

    schedule_bp = {
        "PersumptiveInc44AD": {
            "GrsTotalTrnOver": int(ret.b44ad.turnover_digital + ret.b44ad.turnover_other),
            "GrsTrnOverBank": int(ret.b44ad.turnover_digital),
            "GrsTrnOverAnyOthMode": int(ret.b44ad.turnover_other),
            "PersumptiveInc44AD6Per": int(ret.b44ad.income_digital),
            "PersumptiveInc44AD8Per": int(ret.b44ad.income_other),
            "TotPersumptiveInc44AD": int(ret.b44ad.income),
        },
        "PersumptiveInc44ADA": {
            "GrsReceipt": int(ret.p44ada.gross_receipts),
            "TotPersumptiveInc44ADA": int(ret.p44ada.income),
        },
        "PersumptiveInc44AE": {
            "TotPersumInc44AE": int(ret.b44ae.income),
            "SalInterestByFirm": 0,
            "TotalPersumptiveInc": int(ret.presumptive_income()),
            "IncChargeableUnderBus": int(ret.presumptive_income()),
        },
    }

    tp = ret.taxes_paid
    tax_paid = {
        "TaxesPaid": {
            "AdvanceTax": int(ret.advance_tax),
            "TDS": int(ret.tds),
            "TCS": int(ret.tcs),
            "SelfAssessmentTax": int(tp.self_assessment_tax),
            "TotalTaxesPaid": int(ret.total_taxes_paid()),
        },
        "BalTaxPayable": int(max(0, -s["RefundOrPayable"])),
    }
    refund_due = int(max(0, s["RefundOrPayable"]))
    if not ret.bank.ifsc:
        raise ValueError("Bank account (IFSC) is mandatory for ITR-4 refund credit")
    refund = {
        "RefundDue": refund_due,
        "BankAccountDtls": {"AddtnlBankDetails": [{
            "IFSCCode": ret.bank.ifsc,
            "BankAccountNo": ret.bank.account_no,
            "BankName": ret.bank.bank_name,
            "AccountType": "SB",
            "UseForRefund": "true" if refund_due > 0 else "false",
        }]},
    }
    verification = {
        "Declaration": {
            "AssesseeVerName": f"{p.first_name} {p.last_name}".strip() or "-",
            "FatherName": p.father_name or "NA",
            "AssesseeVerPAN": p.pan,
        },
        "Capacity": "S",
        "Place": pa.area_locality or "-",
    }

    itr4 = {
        "Form_ITR4": {
            "FormName": "ITR-4",
            "Description": "ITR-4 SUGAM for presumptive income - AY 2026-27",
            "AssessmentYear": C.ASSESSMENT_YEAR,
            "SchemaVer": "Ver1.0",
            "FormVer": "Ver1.0",
        },
        "FilingStatus": filing_status,
        "PersonalInfo": personal_info,
        "IncomeDeductions": income,
        "TaxComputation": tax_computation,
        "ScheduleBP": schedule_bp,
        "TaxPaid": tax_paid,
        "Refund": refund,
        "Verification": verification,
    }
    # conditional schedules (arrays with minItems=1 in schema)
    if tp.tds_salary:
        itr4["TDSonSalaries"] = {
            "TotalTDSonSalaries": int(sum(t.tds_amount for t in tp.tds_salary)),
            "TDSonSalary": [
                {"EmployerOrDeductorOrCollectDetl": {
                    "TAN": t.tan, "EmployerOrDeductorOrCollecterName": t.employer_name},
                 "IncChrgSal": 0, "TotalTDSSal": int(t.tds_amount)}
                for t in tp.tds_salary],
        }
    if tp.tcs:
        itr4["ScheduleTCS"] = {
            "TotalSchTCS": int(sum(t.tcs_amount for t in tp.tcs)),
            "TCS": [{"EmployerOrDeductorOrCollectDetl": {
                        "TAN": t.tan, "EmployerOrDeductorOrCollecterName": t.collector_name},
                     "AmtTaxCollected": int(t.tcs_amount), "CollectedYr": "2025",
                     "TotalTCS": int(t.tcs_amount),
                     "AmtTCSClaimedThisYear": int(t.tcs_amount)}
                    for t in tp.tcs],
        }
    if tp.it_challans:
        itr4["ScheduleIT"] = {
            "TotalTaxPayments": int(tp.total_it),
            "TaxPayment": [{"BSRCode": c.bsr_code, "DateDep": _iso(c.date_of_deposit),
                            "SrlNoOfChaln": int(c.challan_serial_no or 0),
                            "Amt": int(c.amount)} for c in tp.it_challans],
        }

    body = json.dumps(itr4, sort_keys=True)
    itr4["CreationInfo"] = {
        "SWVersionNo": "V1-1",
        "SWCreatedBy": "SW00000001",
        "JSONCreatedBy": "SW00000001",
        "JSONCreationDate": datetime.now().strftime("%Y-%m-%d"),
        "IntermediaryCity": "Surat",
        "Digest": base64.b64encode(hashlib.sha256(body.encode()).digest()).decode(),
    }
    return {"ITR": {"ITR4": itr4}}
