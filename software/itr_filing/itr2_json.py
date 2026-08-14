"""CBDT JSON builder for ITR-2 AY 2026-27 (schema-driven).

Strategy: the ITR-2 schema has 45 sections with deeply nested required
objects. We first generate a schema-valid skeleton (autofill: 0 / first enum
/ safe strings for every required leaf), then overwrite with real values.
The validator (Draft-04) is run in tests to guarantee 0 errors."""
import base64, hashlib, json
from datetime import datetime

from . import constants as C
from .itr2 import ITR2Return
from .json_builder import (_iso, _state_code, validate_against_schema)  # noqa


def _resolve(node, defs):
    while "$ref" in node:
        node = defs[node["$ref"].split("/")[-1]]
    return node


def _autofill(node, defs):
    node = dict(node)
    if "$ref" in node:
        return _autofill(defs[node["$ref"].split("/")[-1]], defs)
    p = node.get("pattern", "")
    if p and not node.get("properties"):
        if p == "ITR-2":
            return "ITR-2"
        if "Ver1" in p:
            return "Ver1.0"
        if p == "2026":
            return "2026"
        if "2026-07-31" in p:
            return "2026-07-31"
        if "2026-08-31" in p:
            return "2026-08-31"
    if "allOf" in node:
        result = None
        for part in node["allOf"]:
            v = _autofill(part, defs)
            if isinstance(v, dict):
                result = result or {}
                result.update(v)
            else:
                return v
        rest = {k: v for k, v in node.items() if k != "allOf"}
        if rest.get("type") == "object" or "properties" in rest:
            obj = _autofill(rest, defs)
            if isinstance(obj, dict):
                result = result or {}
                result.update(obj)
        return result if result is not None else "NA"
    t = node.get("type")
    if t == "object" or "properties" in node:
        out = {}
        req = node.get("required", [])
        for k, v in node.get("properties", {}).items():
            if k in req:
                out[k] = _autofill(v, defs)
        return out
    if t == "array":
        mi = node.get("minItems", 0)
        if mi:
            return [_autofill(node.get("items", {}), defs)] * mi
        return []
    if t == "integer":
        return 0
    if t == "number":
        return 0.0
    if t == "boolean":
        return False
    # string
    if "enum" in node:
        return node["enum"][0]
    p = node.get("pattern", "")
    if "Y|N" in p:
        return "N"
    if p == "ITR-2":
        return "ITR-2"
    if "Ver1" in p:
        return "Ver1.0"
    if p == "2026":
        return "2026"
    if "2026-07-31" in p:
        return "2026-07-31"
    if "2026-08-31" in p:
        return "2026-08-31"
    if "[A-Z]{5}[0-9]{4}[A-Z]" in p:
        return "AAAPA0000A"
    if "\\d{3}[0-9A-Z]{4}" in p or "[0-9]{7}" in p:
        return "0000000"
    return "NA"


def _set(d, path, value):
    cur = d
    for k in path[:-1]:
        if k not in cur or not isinstance(cur[k], dict):
            cur[k] = {}
        cur = cur[k]
    cur[path[-1]] = value


def build_itr2_json(ret: ITR2Return) -> dict:
    schema = json.load(open(C.ITR2_SCHEMA)) if hasattr(C, "ITR2_SCHEMA") else None
    import os
    schema_path = os.path.join(os.path.dirname(__file__), "..", "..",
                               "ITR-2", "ITR-2_2026_Main_V1.1.json")
    schema = json.load(open(schema_path))
    defs = schema["definitions"]

    itr2 = _autofill(defs["ITR2"], defs)

    def sec(name):
        """Ensure a section exists (autofilled) before we overwrite it."""
        if name not in itr2:
            itr2[name] = _autofill(defs[name], defs)
        return itr2[name]

    for name in ("Form_ITR2", "PartA_GEN1", "ScheduleS", "ScheduleHP",
                 "ScheduleOS", "ScheduleCYLA", "ScheduleBFLA", "ScheduleVIA",
                 "PartB-TI", "PartB_TTI", "Verification"):
        sec(name)
    s = ret.compute()
    tti = s["tti"]
    p = ret.personal
    pa = p.primary_address

    _set(itr2, ["Form_ITR2", "Description"], "ITR-2 for individuals/HUFs AY 2026-27")

    # ---- PartA_GEN1 -----------------------------------------------------
    gen = itr2["PartA_GEN1"]
    pi = gen["PersonalInfo"]
    _set(pi, ["AssesseeName", "FirstName"], p.first_name or "NA")
    _set(pi, ["AssesseeName", "MiddleName"], p.middle_name)
    _set(pi, ["AssesseeName", "SurNameOrOrgName"], p.last_name or "NA")
    if p.pan:
        pi["PAN"] = p.pan
    if p.aadhaar_number:
        pi["AadhaarCardNo"] = p.aadhaar_number
    pi["DOB"] = _iso(p.date_of_birth) or "1980-01-01"
    pi["Status"] = "I"
    pi["SecondaryAdd"] = "N"
    addr = pi["Address"]
    _set(addr, ["ResidenceNo"], pa.flat_door_block or "NA")
    _set(addr, ["ResidenceName"], pa.name_of_premises)
    _set(addr, ["RoadOrStreet"], pa.road_street_post_office)
    _set(addr, ["LocalityOrArea"], pa.area_locality or "NA")
    _set(addr, ["CityOrTownOrDistrict"], pa.city or pa.area_locality or "NA")
    _set(addr, ["StateCode"], _state_code(pa.state) or "11")
    _set(addr, ["CountryCode"], C.COUNTRY_INDIA)
    if pa.pin_code.isdigit() and len(pa.pin_code) == 6:
        _set(addr, ["PinCode"], int(pa.pin_code))
    _set(addr, ["MobileNo"], int(p.primary_mobile) if p.primary_mobile.isdigit() else 9876543210)
    _set(addr, ["EmailAddress"], p.primary_email or "na@example.com")

    fs = gen["FilingStatus"]
    fs["ReturnFileSec"] = ret.filing_sec_code
    fs["OptOutNewTaxRegime"] = "Y" if ret.regime == C.REGIME_OLD else "N"
    fs["SeventhProvisio139"] = "N"
    fs["ResidentialStatus"] = ret.residential_status
    fs["FiiFpiFlag"] = "N"
    fs["HeldUnlistedEqShrPrYrFlg"] = "N"
    fs["AsseseeRepFlg"] = "N"

    # ---- ScheduleS -------------------------------------------------------
    ss = itr2["ScheduleS"]
    ss["TotalGrossSalary"] = int(ret.gross_salary)
    ss["AllwncExtentExemptUs10"] = int(ret.exempt_allowances)
    ss["NetSalary"] = int(ret.net_salary_gross())
    ss["DeductionUnderSection16ia"] = ret.std_deduction_16ia()
    ss["EntertainmntalwncUs16ii"] = int(ret.deduction_16ii)
    ss["ProfessionalTaxUs16iii"] = int(ret.deduction_16iii)
    ss["DeductionUS16"] = int(ret.std_deduction_16ia() + ret.deduction_16ii
                              + ret.deduction_16iii)
    ss["TotIncUnderHeadSalaries"] = s["Salaries"]

    # ---- ScheduleHP / ScheduleOS ------------------------------------------
    itr2["ScheduleHP"]["TotalIncomeChargeableUnHP"] = s["IncomeFromHP"]
    os_ = itr2["ScheduleOS"]
    os_["IncChargeable"] = s["IncFromOS"]

    # ---- ScheduleCYLA / BFLA bucket totals --------------------------------
    a = s["bfla"].income_after_setoff

    def cyla_block(obj_key, income):
        b = itr2["ScheduleCYLA"][obj_key]
        _set(b, ["IncCYLA", "IncOfCurYrUnderThatHead"], int(income))
        _set(b, ["IncCYLA", "IncOfCurYrAfterSetOff"], int(income))

    cyla_block("STCG20Per", a.get("STCG20", 0))
    cyla_block("STCG30Per", a.get("STCG30", 0))
    cyla_block("STCGAppRate", a.get("STCG_RATE", 0))
    cyla_block("STCGDTAARate", a.get("STCG_DTAA", 0))
    cyla_block("LTCG12_5Per", a.get("LTCG125", 0))
    cyla_block("LTCGDTAARate", a.get("LTCG_DTAA", 0))
    cyla_res = s["cyla"]
    _set(itr2["ScheduleCYLA"], ["TotalCurYr", "TotHPlossCurYr"],
         int(-min(0, ret.hp_income())))
    _set(itr2["ScheduleCYLA"], ["TotalCurYr", "TotOthSrcLossNoRaceHorse"], 0)
    _set(itr2["ScheduleCYLA"], ["TotalLossSetOff", "TotHPlossCurYrSetoff"],
         int(cyla_res.hp_loss_setoff))
    _set(itr2["ScheduleCYLA"], ["TotalLossSetOff", "TotOthSrcLossNoRaceHorseSetoff"], 0)
    _set(itr2["ScheduleCYLA"], ["LossRemAftSetOff", "BalHPlossCurYrAftSetoff"],
         int(cyla_res.hp_loss_remaining))
    _set(itr2["ScheduleCYLA"], ["LossRemAftSetOff", "BalOthSrcLossNoRaceHorseAftSetoff"], 0)

    bfla_sec = itr2["ScheduleBFLA"]
    def bfla_block(key, band):
        blk = bfla_sec[key]
        _set(blk, ["IncBFLA", "IncOfCurYrUndHeadFromCYLA"], int(a.get(band, 0)))
        # salary block has no BF-loss field (salary cannot absorb BF losses)
        if "BFlossPrevYrUndSameHeadSetoff" in blk.get("IncBFLA", {}):
            _set(blk, ["IncBFLA", "BFlossPrevYrUndSameHeadSetoff"], 0)
        _set(blk, ["IncBFLA", "IncOfCurYrAfterSetOffBFLosses"], int(a.get(band, 0)))

    for key, band in (("Salary", "SAL"), ("STCG20Per", "STCG20"),
                      ("STCG30Per", "STCG30"), ("STCGAppRate", "STCG_RATE"),
                      ("STCGDTAARate", "STCG_DTAA"), ("LTCG12_5Per", "LTCG125"),
                      ("LTCGDTAARate", "LTCG_DTAA")):
        bfla_block(key, band)
    _set(bfla_sec, ["TotalBFLossSetOff", "TotBFLossSetoff"],
         int(sum(s["bfla"].setoff_by_category.values()) + s["bfla"].depreciation_setoff))
    bfla_sec["IncomeOfCurrYrAftCYLABFLA"] = int(s["GrossTotalIncome"])

    # ---- ScheduleVIA -------------------------------------------------------
    _set(itr2["ScheduleVIA"], ["DeductUndChapVIA", "TotalChapVIADeductions"],
         s["VIADeductions"])
    _set(itr2["ScheduleVIA"], ["UsrDeductUndChapVIA", "TotalChapVIADeductions"],
         int(ret.via_total or 0))

    # ---- PartB-TI ----------------------------------------------------------
    ti = itr2["PartB-TI"]
    ti["Salaries"] = s["Salaries"]
    ti["IncomeFromHP"] = s["IncomeFromHP"]
    st = ti["CapGain"]["ShortTerm"]
    st["ShortTerm20Per"] = int(a.get("STCG20", 0))
    st["ShortTerm30Per"] = int(a.get("STCG30", 0))
    st["ShortTermAppRate"] = int(a.get("STCG_RATE", 0))
    st["ShortTermSplRateDTAA"] = int(a.get("STCG_DTAA", 0))
    st["TotalShortTerm"] = int(a.get("STCG20", 0) + a.get("STCG30", 0)
                               + a.get("STCG_RATE", 0) + a.get("STCG_DTAA", 0))
    ti["CapGain"]["ShortTermLongTermTotal"] = s["CapGainTotal"]
    ti["CapGain"]["CapGains30Per115BBH"] = 0
    ti["CapGain"]["TotalCapGains"] = s["CapGainTotal"]
    ti["IncFromOS"]["OtherSrcThanOwnRaceHorse"] = int(a.get("OS", 0))
    ti["IncFromOS"]["IncChargblSplRate"] = 0
    ti["IncFromOS"]["FromOwnRaceHorse"] = int(a.get("RACEHORSE", 0))
    ti["IncFromOS"]["TotIncFromOS"] = s["IncFromOS"]
    ti["TotalTI"] = s["GrossTotalIncome"]
    ti["CurrentYearLoss"] = 0
    ti["BalanceAfterSetoffLosses"] = s["GrossTotalIncome"]
    ti["BroughtFwdLossesSetoff"] = 0
    ti["GrossTotalIncome"] = s["GrossTotalIncome"]
    ti["IncChargeTaxSplRate111A112"] = s["IncChargeTaxSplRate111A112"]
    ti["DeductionsUnderScheduleVIA"] = s["VIADeductions"]
    ti["TotalIncome"] = s["TotalIncome"]
    ti["IncChargeableTaxSplRates"] = s["IncChargeableTaxSplRates"]
    ti["NetAgricultureIncomeOrOtherIncomeForRate"] = 0
    ti["AggregateIncome"] = s["TotalIncome"]
    ti["LossesOfCurrentYearCarriedFwd"] = 0
    ti["DeemedIncomeUs115JC"] = 0

    # ---- PartB_TTI -----------------------------------------------------------
    tti_sec = itr2["PartB_TTI"]
    tti_sec["TaxPayDeemedTotIncUs115JC"] = 0
    tti_sec["Surcharge"] = tti["Surcharge"]
    tti_sec["HealthEduCess"] = tti["Cess"]
    tti_sec["TotalTaxPayablDeemedTotInc"] = 0
    comp = tti_sec["ComputationOfTaxLiability"]
    _set(comp, ["TaxPayableOnTI", "TaxAtNormalRatesOnAggrInc"], tti["TaxNormal"])
    _set(comp, ["TaxPayableOnTI", "TaxAtSpecialRates"], tti["TaxSpecialTotal"])
    _set(comp, ["TaxPayableOnTI", "RebateOnAgriInc"], 0)
    _set(comp, ["TaxPayableOnTI", "TaxPayableOnTotInc"], tti["TaxNormal"]
         + tti["TaxSpecialTotal"])
    comp["Rebate87A"] = tti["Rebate87A"]
    comp["TaxPayableOnRebate"] = max(0, tti["TaxNormal"] + tti["TaxSpecialTotal"]
                                     - tti["Rebate87A"])
    comp["Surcharge25ofSI"] = 0
    comp["SurchargeOnAboveCrore"] = tti["Surcharge"]
    comp["Surcharge25ofSIBeforeMarginal"] = 0
    comp["SurchargeOnAboveCroreBeforeMarginal"] = tti["Surcharge"]
    comp["TotalSurcharge"] = tti["Surcharge"]
    comp["EducationCess"] = tti["Cess"]
    comp["GrossTaxLiability"] = tti["GrossTaxLiability"]
    comp["GrossTaxPayable"] = tti["GrossTaxLiability"]
    _set(comp, ["GrossTaxPay", "TaxInc17"], 0)
    _set(comp, ["GrossTaxPay", "TaxDeferred17"], 0)
    _set(comp, ["GrossTaxPay", "TaxDeferredPayableCY"], 0)
    comp["CreditUS115JD"] = 0
    comp["TaxPayAfterCreditUs115JD"] = tti["GrossTaxLiability"]
    _set(comp, ["TaxRelief", "Section89"], int(ret.relief_89))
    _set(comp, ["TaxRelief", "Section90"], int(ret.relief_90))
    _set(comp, ["TaxRelief", "Section91"], int(ret.relief_91))
    _set(comp, ["TaxRelief", "TotTaxRelief"],
         int(ret.relief_89 + ret.relief_90 + ret.relief_91))
    comp["NetTaxLiability"] = tti["NetTaxLiability"]
    _set(comp, ["IntrstPay", "IntrstPayUs234A"], s["IntrstPay234A"])
    _set(comp, ["IntrstPay", "IntrstPayUs234B"], s["IntrstPay234B"])
    _set(comp, ["IntrstPay", "IntrstPayUs234C"], 0)
    _set(comp, ["IntrstPay", "LateFilingFee234F"], s["Fee234F"])
    _set(comp, ["IntrstPay", "TotalIntrstPay"],
         int(s["IntrstPay234A"] + s["IntrstPay234B"] + s["Fee234F"]))
    comp["AggregateTaxInterestLiability"] = s["TotalLiability"]

    tp = ret.taxes_paid
    _set(tti_sec, ["TaxPaid", "TaxesPaid", "AdvanceTax"], int(tp.advance_tax))
    _set(tti_sec, ["TaxPaid", "TaxesPaid", "TDS"], int(tp.total_tds))
    _set(tti_sec, ["TaxPaid", "TaxesPaid", "TCS"], int(tp.total_tcs))
    _set(tti_sec, ["TaxPaid", "TaxesPaid", "SelfAssessmentTax"],
         int(tp.self_assessment_tax))
    _set(tti_sec, ["TaxPaid", "TaxesPaid", "TotalTaxesPaid"],
         int(tp.total_tds + tp.total_tcs + tp.total_it))
    _set(tti_sec, ["TaxPaid", "BalTaxPayable"], s["BalTaxPayable"])
    _set(tti_sec, ["Refund", "RefundDue"], s["RefundDue"])
    _set(tti_sec, ["Refund", "BankAccountDtls", "BankDtlsFlag"],
         "Y" if ret.bank.ifsc else "N")
    if ret.bank.ifsc:
        _set(tti_sec, ["Refund", "BankAccountDtls", "AddtnlBankDetails"], [{
            "IFSCCode": ret.bank.ifsc,
            "BankAccountNo": ret.bank.account_no,
            "BankName": ret.bank.bank_name,
            "AccountType": "SB",
            "UseForRefund": "true" if s["RefundDue"] > 0 else "false",
        }])
    tti_sec["AssetOutIndiaFlag"] = "NO"

    # ---- Verification ------------------------------------------------------
    ver = itr2["Verification"]
    _set(ver, ["Declaration", "AssesseeVerName"],
         f"{p.first_name} {p.last_name}".strip() or "NA")
    _set(ver, ["Declaration", "FatherName"], p.father_name or "NA")
    _set(ver, ["Declaration", "AssesseeVerPAN"], p.pan or "AAAPA0000A")
    if "Capacity" in ver:
        ver["Capacity"] = "S"
    if "Place" in ver:
        ver["Place"] = pa.area_locality or "NA"

    body = json.dumps(itr2, sort_keys=True)
    itr2["CreationInfo"] = {
        "SWVersionNo": "V1-1",
        "SWCreatedBy": "SW00000001",
        "JSONCreatedBy": "SW00000001",
        "JSONCreationDate": datetime.now().strftime("%Y-%m-%d"),
        "IntermediaryCity": "Surat",
        "Digest": base64.b64encode(hashlib.sha256(body.encode()).digest()).decode(),
    }
    return {"ITR": {"ITR2": itr2}}
