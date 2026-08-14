"""CBDT JSON builder for ITR-3 AY 2026-27 (schema-driven, same autofill
technique as itr2_json)."""
import base64, hashlib, json, os
from datetime import datetime

from . import constants as C
from .itr3 import ITR3Return
from .itr2_json import _autofill, _set
from .json_builder import (_iso, _state_code, validate_against_schema)  # noqa


def build_itr3_json(ret: ITR3Return) -> dict:
    schema_path = os.path.join(os.path.dirname(__file__), "..", "..",
                               "ITR-3", "ITR-3_2026_Main_V1.1.json")
    schema = json.load(open(schema_path))
    defs = schema["definitions"]

    itr3 = _autofill(defs["ITR3"], defs)

    def sec(name):
        if name not in itr3:
            itr3[name] = _autofill(defs[name], defs)
        return itr3[name]

    for name in ("Form_ITR3", "PartA_GEN1", "PartA_GEN2", "PARTA_BS", "PARTA_PL",
                 "ITR3ScheduleBP", "ScheduleCYLA", "ScheduleBFLA",
                 "PartB-TI", "PartB_TTI", "Verification"):
        sec(name)
    s = ret.compute()
    tti = s["tti"]
    p = ret.personal
    pa = p.primary_address
    bp = s["bp"]
    a = s["bfla"].income_after_setoff

    _set(itr3, ["Form_ITR3", "Description"], "ITR-3 for individuals/HUFs with business income AY 2026-27")

    # ---- PartA_GEN1 ------------------------------------------------------
    gen = itr3["PartA_GEN1"]
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
    if "IncFrmBusOrProf" in fs:
        fs["IncFrmBusOrProf"] = "Y"
    fs["SeventhProvisio139"] = "N"
    fs["ResidentialStatus"] = ret.residential_status
    fs["HeldUnlistedEqShrPrYrFlg"] = "N"
    if "ForeignExchangeFlag" in fs:
        fs["ForeignExchangeFlag"] = "N"
    fs["FiiFpiFlag"] = "N"
    fs["ItrFilingDueDate"] = "2026-08-31"

    # ---- PartA_GEN2 audit flags -------------------------------------------
    audit = itr3["PartA_GEN2"]["AuditInfo"]
    for k in audit:
        audit[k] = "N"

    # ---- PARTA_PL ----------------------------------------------------------
    pl = itr3["PARTA_PL"]
    net_pl = sum(b.net_profit_as_per_pl for b in ret.businesses)
    _set(pl, ["CreditsToPL", "TotCreditsToPL"], int(net_pl))
    if "TotDebitsToPL" in pl.get("DebitsToPL", {}):
        _set(pl, ["DebitsToPL", "TotDebitsToPL"], 0)
    _set(pl, ["TaxProvAppr", "ProfitAfterTax"], int(net_pl))
    nob = pl.get("NoBooksOfAccPL", {})
    if "NetProfit" in nob:
        nob["NetProfit"] = int(s["BusinessIncome"])
    if "GrossReceipt" in nob:
        nob["GrossReceipt"] = int(sum(b.income_before_depreciation() for b in ret.businesses))
    pl["TurnverFrmSpecActivity"] = int(bp.specified)
    pl["NetIncomeFrmSpecActivity"] = int(bp.specified)
    if "TotalPrsumptvIncUs44E" in pl:
        pl["TotalPrsumptvIncUs44E"] = int(bp.presumptive_44ae)

    # ---- ITR3ScheduleBP ------------------------------------------------------
    bps = itr3["ITR3ScheduleBP"]
    bio = bps["BusinessIncOthThanSpec"]
    bio["ProfBfrTaxPL"] = int(net_pl)
    bio["NetPLFromSpecBus"] = int(bp.speculation)
    bio["NetPLFromSpecifiedBus"] = int(bp.specified)
    _set(bio, ["ProfitLossInclRefrdSec", "ProfitLossUs44AD"], int(bp.presumptive_44ad))
    _set(bio, ["ProfitLossInclRefrdSec", "ProfitLossUs44ADA"], int(bp.presumptive_44ada))
    _set(bio, ["ProfitLossInclRefrdSec", "ProfitLossUs44AE"], int(bp.presumptive_44ae))
    presumptive_total = int(bp.presumptive_44ad + bp.presumptive_44ada + bp.presumptive_44ae)
    bio["TotalProfitFrmActCvrd"] = presumptive_total
    bio["BalancePLOthThanSpecBus"] = int(net_pl)
    bio["AdjustedPLOthThanSpecBus"] = int(net_pl)
    dep_it = int(sum(b.total_depreciation() for b in ret.businesses))
    _set(bio, ["DepreciationAllowITAct32", "DepreciationAllowUs32_1_ii"], dep_it)
    _set(bio, ["DepreciationAllowITAct32", "DepreciationAllowUs32_1_i"], 0)
    _set(bio, ["DepreciationAllowITAct32", "TotDeprAllowITAct"], dep_it)
    bio["NetPLAftAdjBusOthThanSpec"] = int(bp.normal)
    bio["NetPLBusOthThanSpec7A7B7C"] = int(max(0, bp.normal))
    _set(bps, ["SpecBusinessInc", "NetPLFrmSpecBus"], int(bp.speculation))
    _set(bps, ["SpecBusinessInc", "AdjustedPLFrmSpecuBus"], int(bp.speculation))
    _set(bps, ["SpecifiedBusinessInc", "NetPLFrmSpecifiedBus"], int(bp.specified)) \
        if "NetPLFrmSpecifiedBus" in bps.get("SpecifiedBusinessInc", {}) else None
    bps["IncChrgUnHdProftGain"] = int(s["BusinessIncome"] + s["SpeculationIncome"])
    _set(bps, ["BusSetoffCurrYr", "LossSetOffOnBusLoss"], 0)
    _set(bps, ["BusSetoffCurrYr", "TotLossSetOffOnBus"], 0)
    _set(bps, ["BusSetoffCurrYr", "LossRemainSetOffOnBus"], int(max(0, -bp.normal)))

    # ---- ScheduleCYLA / BFLA (CG bands only for ITR-3 too) -------------------
    def cyla_block(obj_key, income):
        b = itr3["ScheduleCYLA"][obj_key]
        _set(b, ["IncCYLA", "IncOfCurYrUnderThatHead"], int(income))
        _set(b, ["IncCYLA", "IncOfCurYrAfterSetOff"], int(income))

    cyla_block("STCG20Per", a.get("STCG20", 0))
    cyla_block("STCG30Per", a.get("STCG30", 0))
    cyla_block("STCGAppRate", a.get("STCG_RATE", 0))
    cyla_block("STCGDTAARate", a.get("STCG_DTAA", 0))
    cyla_block("LTCG12_5Per", a.get("LTCG125", 0))
    cyla_block("LTCGDTAARate", a.get("LTCG_DTAA", 0))
    cyla_res = s["cyla"]
    _set(itr3["ScheduleCYLA"], ["TotalCurYr", "TotHPlossCurYr"],
         int(-min(0, sum(hp.income() for hp in ret.house_properties))))
    _set(itr3["ScheduleCYLA"], ["TotalCurYr", "TotOthSrcLossNoRaceHorse"], 0)
    _set(itr3["ScheduleCYLA"], ["TotalLossSetOff", "TotHPlossCurYrSetoff"],
         int(cyla_res.hp_loss_setoff))
    _set(itr3["ScheduleCYLA"], ["TotalLossSetOff", "TotOthSrcLossNoRaceHorseSetoff"], 0)
    lra = itr3["ScheduleCYLA"]["LossRemAftSetOff"]
    if "BalHPlossCurYrAftSetoff" in lra:
        lra["BalHPlossCurYrAftSetoff"] = int(cyla_res.hp_loss_remaining)
    if "BalOthSrcLossNoRaceHorseAftSetoff" in lra:
        lra["BalOthSrcLossNoRaceHorseAftSetoff"] = 0

    bfla_sec = itr3["ScheduleBFLA"]

    def bfla_block(key, band):
        blk = bfla_sec[key]
        _set(blk, ["IncBFLA", "IncOfCurYrUndHeadFromCYLA"], int(a.get(band, 0)))
        if "BFlossPrevYrUndSameHeadSetoff" in blk.get("IncBFLA", {}):
            _set(blk, ["IncBFLA", "BFlossPrevYrUndSameHeadSetoff"], 0)
        _set(blk, ["IncBFLA", "IncOfCurYrAfterSetOffBFLosses"], int(a.get(band, 0)))

    for key, band in (("Salary", "SAL"), ("STCG20Per", "STCG20"),
                      ("STCG30Per", "STCG30"), ("STCGAppRate", "STCG_RATE"),
                      ("STCGDTAARate", "STCG_DTAA"), ("LTCG12_5Per", "LTCG125"),
                      ("LTCGDTAARate", "LTCG_DTAA")):
        bfla_block(key, band)
    tb = bfla_sec["TotalBFLossSetOff"]
    if "TotBFLossSetoff" in tb:
        tb["TotBFLossSetoff"] = int(sum(s["bfla"].setoff_by_category.values())
                                    + s["bfla"].depreciation_setoff)
    bfla_sec["IncomeOfCurrYrAftCYLABFLA"] = int(s["GrossTotalIncome"])

    # ---- ScheduleVIA ----------------------------------------------------------
    if "ScheduleVIA" in itr3:
        _set(itr3["ScheduleVIA"], ["DeductUndChapVIA", "TotalChapVIADeductions"],
             s["VIADeductions"])
        _set(itr3["ScheduleVIA"], ["UsrDeductUndChapVIA", "TotalChapVIADeductions"],
             int(ret.via_total or 0))

    # ---- PartB-TI ---------------------------------------------------------------
    ti = itr3["PartB-TI"]
    ti["Salaries"] = s["Salaries"]
    ti["IncomeFromHP"] = s["IncomeFromHP"]
    _set(ti, ["ProfBusGain", "ProfGainNoSpecBus"], int(s["BusinessIncome"]))
    _set(ti, ["ProfBusGain", "ProfGainSpecBus"], int(s["SpeculationIncome"]))
    _set(ti, ["ProfBusGain", "ProfGainSpecifiedBus"], 0)
    _set(ti, ["ProfBusGain", "ProfIncome115BBF"], 0)
    _set(ti, ["ProfBusGain", "TotProfBusGain"],
         int(s["BusinessIncome"] + s["SpeculationIncome"]))
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
    _set(ti, ["DeductionsUndSchVIADtl", "PartBchapterVIA"], s["VIADeductions"])
    _set(ti, ["DeductionsUndSchVIADtl", "PartCchapterVIA"], 0)
    _set(ti, ["DeductionsUndSchVIADtl", "TotDeductUndSchVIA"], s["VIADeductions"])
    ti["DeductionsUnder10Aor10AA"] = 0
    ti["TotalIncome"] = s["TotalIncome"]
    ti["IncChargeableTaxSplRates"] = s["IncChargeableTaxSplRates"]
    ti["NetAgricultureIncomeOrOtherIncomeForRate"] = 0
    ti["AggregateIncome"] = s["TotalIncome"]
    ti["LossesOfCurrentYearCarriedFwd"] = 0
    ti["DeemedIncomeUs115JC"] = 0

    # ---- PartB_TTI -----------------------------------------------------------------
    tti_sec = itr3["PartB_TTI"]
    comp = tti_sec["ComputationOfTaxLiability"]
    if isinstance(comp.get("TaxPayableOnDeemedTI"), dict):
        for k in list(comp["TaxPayableOnDeemedTI"].keys()):
            comp["TaxPayableOnDeemedTI"][k] = 0
    tpt = comp["TaxPayableOnTI"]
    tpt["TaxAtNormalRatesOnAggrInc"] = tti["TaxNormal"]
    tpt["TaxAtSpecialRates"] = tti["TaxSpecialTotal"]
    tpt["RebateOnAgriInc"] = 0
    tpt["TaxPayableOnTotInc"] = tti["TaxNormal"] + tti["TaxSpecialTotal"]
    tpt["Rebate87A"] = tti["Rebate87A"]
    tpt["TaxPayableOnRebate"] = max(0, tti["TaxNormal"] + tti["TaxSpecialTotal"]
                                    - tti["Rebate87A"])
    if "Surcharge25ofSI" in tpt:
        tpt["Surcharge25ofSI"] = 0
    if "SurchargeOnAboveCrore" in tpt:
        tpt["SurchargeOnAboveCrore"] = tti["Surcharge"]
    if "Surcharge25ofSIBeforeMarginal" in tpt:
        tpt["Surcharge25ofSIBeforeMarginal"] = 0
    if "SurchargeOnAboveCroreBeforeMarginal" in tpt:
        tpt["SurchargeOnAboveCroreBeforeMarginal"] = tti["Surcharge"]
    if "TotalSurcharge" in tpt:
        tpt["TotalSurcharge"] = tti["Surcharge"]
    if "EducationCess" in tpt:
        tpt["EducationCess"] = tti["Cess"]
    if "GrossTaxLiability" in tpt:
        tpt["GrossTaxLiability"] = tti["GrossTaxLiability"]
    comp["GrossTaxPayable"] = tti["GrossTaxLiability"]
    _set(comp, ["GrossTaxPay", "TaxInc17"], 0)
    _set(comp, ["GrossTaxPay", "TaxDeferred17"], 0)
    _set(comp, ["GrossTaxPay", "TaxDeferredPayableCY"], 0)
    comp["CreditUS115JD"] = 0
    comp["TaxPayAfterCreditUs115JD"] = tti["GrossTaxLiability"]
    comp["NetTaxLiability"] = tti["NetTaxLiability"]
    _set(comp, ["IntrstPay", "IntrstPayUs234A"], s["IntrstPay234A"])
    _set(comp, ["IntrstPay", "IntrstPayUs234B"], s["IntrstPay234B"])
    _set(comp, ["IntrstPay", "IntrstPayUs234C"], 0)
    _set(comp, ["IntrstPay", "LateFilingFee234F"], s["Fee234F"])
    if "TotalIntrstPay" in comp.get("IntrstPay", {}):
        _set(comp, ["IntrstPay", "TotalIntrstPay"],
             int(s["IntrstPay234A"] + s["IntrstPay234B"] + s["Fee234F"]))
    if "TaxRelief" in comp and isinstance(comp["TaxRelief"], dict):
        if "Section89" in comp["TaxRelief"]:
            comp["TaxRelief"]["Section89"] = int(ret.relief_89)
        if "Section90" in comp["TaxRelief"]:
            comp["TaxRelief"]["Section90"] = int(ret.relief_90)
        if "Section91" in comp["TaxRelief"]:
            comp["TaxRelief"]["Section91"] = int(ret.relief_91)
        if "TotTaxRelief" in comp["TaxRelief"]:
            comp["TaxRelief"]["TotTaxRelief"] = int(ret.relief_89 + ret.relief_90
                                                    + ret.relief_91)
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
    bankd = tti_sec["Refund"]["BankAccountDtls"]
    if "BankDtlsFlag" in bankd:
        bankd["BankDtlsFlag"] = "Y" if ret.bank.ifsc else "N"
    if ret.bank.ifsc:
        _set(tti_sec, ["Refund", "BankAccountDtls", "AddtnlBankDetails"], [{
            "IFSCCode": ret.bank.ifsc,
            "BankAccountNo": ret.bank.account_no,
            "BankName": ret.bank.bank_name,
            "AccountType": "SB",
            "UseForRefund": "true" if s["RefundDue"] > 0 else "false",
        }])
    if "AssetOutIndiaFlag" in tti_sec:
        tti_sec["AssetOutIndiaFlag"] = "NO"

    # ---- Verification -----------------------------------------------------------------
    ver = itr3["Verification"]
    _set(ver, ["Declaration", "AssesseeVerName"],
         f"{p.first_name} {p.last_name}".strip() or "NA")
    _set(ver, ["Declaration", "FatherName"], p.father_name or "NA")
    _set(ver, ["Declaration", "AssesseeVerPAN"], p.pan or "AAAPA0000A")
    if "Capacity" in ver:
        ver["Capacity"] = "S"
    if "Place" in ver:
        ver["Place"] = pa.area_locality or "NA"
    if "Date" in ver:
        ver["Date"] = _iso(ret.verification_date) or "2026-07-31"

    body = json.dumps(itr3, sort_keys=True)
    itr3["CreationInfo"] = {
        "SWVersionNo": "V1-1",
        "SWCreatedBy": "SW00000001",
        "JSONCreatedBy": "SW00000001",
        "JSONCreationDate": datetime.now().strftime("%Y-%m-%d"),
        "IntermediaryCity": "Surat",
        "Digest": base64.b64encode(hashlib.sha256(body.encode()).digest()).decode(),
    }
    return {"ITR": {"ITR3": itr3}}
