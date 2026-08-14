"""Map flat UI input dicts -> return objects (ITR-1..4) and format results."""
from datetime import date

from itr_filing import constants as C
from itr_filing.itr1 import (ITR1Return, PersonalInfo, FilingStatus, Salary,
                             ChapterVIA, TaxesPaid, TDSonSalary, Verification,
                             BankAccount)
from itr_filing.itr2 import ITR2Return
from itr_filing.itr3 import ITR3Return
from itr_filing.itr4 import ITR4Return, Business44AD, Profession44ADA, Business44AE
from itr_filing.house_property import HouseProperty
from itr_filing.capital_gains import CapitalAsset
from itr_filing.business import BusinessPL, DepreciationBlock
from itr_filing.schedules import Schedule80D


def _d(s):
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return date.fromisoformat(s) if fmt == "%Y-%m-%d" else \
                date(*map(int, reversed(s.split("/"))))
        except Exception:
            continue
    return None


def _f(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _personal(i: dict) -> PersonalInfo:
    p = PersonalInfo(
        first_name=i.get("first_name", ""), middle_name=i.get("middle_name", ""),
        last_name=i.get("last_name", ""), pan=i.get("pan", ""),
        aadhaar_number=i.get("aadhaar", ""), date_of_birth=_d(i.get("dob")),
        nature_of_employment=i.get("employment", "PE"),
        primary_email=i.get("email", ""), primary_mobile=i.get("mobile", ""),
        father_name=i.get("father_name", ""))
    p.primary_address.flat_door_block = i.get("addr1", "")
    p.primary_address.name_of_premises = i.get("addr2", "")
    p.primary_address.area_locality = i.get("locality", "")
    p.primary_address.city = i.get("city", "")
    p.primary_address.state = i.get("state", "")
    p.primary_address.pin_code = i.get("pincode", "")
    return p


def _bank(i: dict) -> BankAccount:
    return BankAccount(ifsc=i.get("ifsc", ""), account_no=i.get("account_no", ""),
                       bank_name=i.get("bank_name", ""),
                       is_for_refund="Yes" if i.get("ifsc") else "No")


def _tds(i: dict) -> TaxesPaid:
    from itr_filing.itr1 import ITChallan
    from datetime import date as _date
    tp = TaxesPaid()
    if _f(i.get("tds_amount")):
        tp.tds_salary = [TDSonSalary(employer_name=i.get("tds_deductor", "Employer"),
                                     tan=i.get("tds_tan", ""), tds_amount=_f(i["tds_amount"]))]
    if _f(i.get("advance_tax")):
        tp.it_challans.append(ITChallan(bsr_code="0000000", date_of_deposit=_date(2026, 3, 10),
                                        challan_serial_no="1", amount=_f(i["advance_tax"]),
                                        mode="Advance Tax"))
    if _f(i.get("self_tax")):
        tp.it_challans.append(ITChallan(bsr_code="0000000", date_of_deposit=_date(2026, 7, 15),
                                        challan_serial_no="2", amount=_f(i["self_tax"]),
                                        mode="Self Assessment Tax"))
    return tp


def build_itr1(i: dict) -> ITR1Return:
    r = ITR1Return()
    r.personal = _personal(i)
    regime = C.REGIME_OLD if i.get("regime") == "old" else C.REGIME_NEW
    r.filing = FilingStatus(opt_out_new_regime="Yes" if regime == C.REGIME_OLD else "No",
                            filing_sec_code=int(i.get("filing_sec", 11)))
    r.salary = Salary(
        salary_17_1=_f(i.get("gross_salary")),
        basic_salary=_f(i.get("basic_salary")),
        dearness_allowance=_f(i.get("da")),
        hra_received=_f(i.get("hra")),
        rent_paid=_f(i.get("rent_paid")),
        metro=bool(i.get("metro")),
        professional_tax_16_iii=_f(i.get("professional_tax")))
    if _f(i.get("exempt_allowances")):
        r.salary.exempt_allowances = [("Sec 10(14)(i) Prescribed allowances/benefits",
                                       _f(i.get("exempt_allowances")))]
    if i.get("hp_type") and i.get("hp_type") != "(Select)":
        r.house_properties = [HouseProperty(
            property_type=i["hp_type"], gross_annual_value=_f(i.get("hp_rent")),
            tax_paid_local_authorities=_f(i.get("hp_municipal_tax")),
            interest_24b=_f(i.get("hp_interest")),
            address=i.get("addr1", ""), city=i.get("city", ""),
            state=i.get("state", ""), pincode=int(i.get("pincode") or 0))]
    from itr_filing.itr1 import OtherSources
    os_ = OtherSources()
    if _f(i.get("savings_interest")):
        os_.items.append(("Interest from Savings Account", "", _f(i.get("savings_interest"))))
    if _f(i.get("other_interest")):
        os_.items.append(("Interest from Deposits", "", _f(i.get("other_interest"))))
    if _f(i.get("dividend")):
        os_.dividend_by_quarter = [_f(i.get("dividend")), 0, 0, 0, 0]
    if _f(i.get("family_pension")):
        os_.items.append(("Family Pension", "", _f(i.get("family_pension"))))
        os_.deduction_57_iia = min(15000, _f(i.get("family_pension")))
    r.other_sources = os_
    r.ltcg_112a.total_sale_consideration = _f(i.get("ltcg_sale"))
    r.ltcg_112a.total_cost_of_acquisition = _f(i.get("ltcg_cost"))
    cv = ChapterVIA(s80c=_f(i.get("s80c")), s80ccc=_f(i.get("s80ccc")),
                    s80ccd1=_f(i.get("s80ccd1")), s80ccd1b=_f(i.get("s80ccd1b")),
                    s80e=_f(i.get("s80e")), s80tta=_f(i.get("s80tta")),
                    s80ttb=_f(i.get("s80ttb")))
    if _f(i.get("s80d_premium")) or i.get("s80d_selection"):
        cv.schedule_80d = Schedule80D(selection=int(i.get("s80d_selection") or 1),
                                      premium_paid=_f(i.get("s80d_premium")),
                                      taxpayer_age=r.personal.age)
    cv.disability_80dd = int(i.get("s80dd_type") or 0)
    cv.severity_80u = int(i.get("s80u_type") or 0)
    cv.selection_80ddb = int(i.get("s80ddb_selection") or 0)
    cv.s80ddb_schedule_amount = _f(i.get("s80ddb_amount"))
    cv.user_total = 10**12  # allow computed caps to govern
    r.chapter_via = cv
    r.taxes_paid = _tds(i)
    r.bank = _bank(i)
    r.relief_89 = _f(i.get("relief_89"))
    r.verification = Verification(verification_date=_d(i.get("verification_date")) or date.today())
    return r


def build_itr4(i: dict) -> ITR4Return:
    r = ITR4Return()
    r.personal = _personal(i)
    r.entity = i.get("entity", "I")
    r.dob = r.personal.date_of_birth
    r.regime = C.REGIME_OLD if i.get("regime") == "old" else C.REGIME_NEW
    r.salary = _f(i.get("gross_salary"))
    r.other_sources = _f(i.get("other_income"))
    r.b44ad = Business44AD(turnover_digital=_f(i.get("turnover_digital")),
                           turnover_other=_f(i.get("turnover_other")),
                           income_digital=_f(i.get("income_digital")),
                           income_other=_f(i.get("income_other")))
    r.p44ada = Profession44ADA(gross_receipts=_f(i.get("gross_receipts_44ada")),
                               income=_f(i.get("income_44ada")))
    r.chapter_via_total = _f(i.get("via_total"))
    r.tds = _f(i.get("tds_amount"))
    r.advance_tax = _f(i.get("advance_tax"))
    r.taxes_paid = _tds(i)
    r.bank = _bank(i)
    r.verification_date = _d(i.get("verification_date")) or date.today()
    return r


def build_itr2(i: dict) -> ITR2Return:
    r = ITR2Return()
    r.personal = _personal(i)
    r.regime = C.REGIME_OLD if i.get("regime") == "old" else C.REGIME_NEW
    r.gross_salary = _f(i.get("gross_salary"))
    r.exempt_allowances = _f(i.get("exempt_allowances"))
    if i.get("hp_type") and i.get("hp_type") != "(Select)":
        r.house_properties = [HouseProperty(property_type=i["hp_type"],
                                            gross_annual_value=_f(i.get("hp_rent")),
                                            interest_24b=_f(i.get("hp_interest")))]
    r.other_sources_normal = _f(i.get("other_income"))
    if _f(i.get("stcg_amount")):
        r.cg_assets.append(CapitalAsset(
            asset_class="securities", listed_stt_paid=True,
            acquisition_date=_d(i.get("stcg_acq_date")) or date(2025, 11, 1),
            transfer_date=_d(i.get("verification_date")) or date(2026, 3, 1),
            sale_consideration=_f(i.get("stcg_amount")) + _f(i.get("stcg_cost")),
            cost_of_acquisition=_f(i.get("stcg_cost"))))
    r.via_total = _f(i.get("via_total"))
    r.taxes_paid = _tds(i)
    r.bank = _bank(i)
    r.verification_date = _d(i.get("verification_date")) or date.today()
    return r


def build_itr3(i: dict) -> ITR3Return:
    r = ITR3Return()
    r.personal = _personal(i)
    r.regime = C.REGIME_OLD if i.get("regime") == "old" else C.REGIME_NEW
    r.gross_salary = _f(i.get("gross_salary"))
    r.exempt_allowances = _f(i.get("exempt_allowances"))
    r.other_sources_normal = _f(i.get("other_income"))
    blocks = []
    if _f(i.get("dep_opening_wdv")) or _f(i.get("dep_additions")):
        blocks = [DepreciationBlock(rate=_f(i.get("dep_rate")) or 0.15,
                                    opening_wdv=_f(i.get("dep_opening_wdv")),
                                    additions_full=_f(i.get("dep_additions")),
                                    additions_half=_f(i.get("dep_additions_half")))]
    r.businesses = [BusinessPL(net_profit_as_per_pl=_f(i.get("net_profit_pl")),
                               inadmissible_expenses=_f(i.get("inadmissible")),
                               depreciation_as_per_pl=_f(i.get("dep_as_per_pl")),
                               depreciation_blocks=blocks)]
    r.via_total = _f(i.get("via_total"))
    r.taxes_paid = _tds(i)
    r.bank = _bank(i)
    r.verification_date = _d(i.get("verification_date")) or date.today()
    return r


BUILDERS = {"ITR-1": build_itr1, "ITR-2": build_itr2,
            "ITR-3": build_itr3, "ITR-4": build_itr4}


def format_result(form: str, ret) -> dict:
    """Uniform result shape for the UI."""
    if form == "ITR-1":
        s = ret.compute_summary()
        return {
            "Gross Total Income": s["GrossTotIncome"],
            "Chapter VI-A Deductions": s["TotalChapterVIA"],
            "Total Income": s["TotalIncome"],
            "Tax on Total Income (D1)": s["TotalTaxPayable"],
            "Rebate u/s 87A (D2)": s["Rebate87A"],
            "Health & Education Cess (D4)": s["EducationCess"],
            "Gross Tax Liability (D5)": s["GrossTaxLiability"],
            "Interest 234A": s["IntrstPay234A"],
            "Interest 234B": s["IntrstPay234B"],
            "Fee 234F": s["FeeIncUS234F"],
            "Total Liability": s["TotTaxPlusIntrstPay"],
            "Refund / (Payable)": s["RefundOrPayable"],
        }
    if form == "ITR-4":
        s = ret.compute_summary()
        return {
            "Presumptive Income": s["PresumptiveIncome"],
            "Gross Total Income": s["GrossTotIncome"],
            "Total Income": s["TotalIncome"],
            "Tax Payable": s["TotalTaxPayable"],
            "Rebate u/s 87A": s["Rebate87A"],
            "Cess": s["EducationCess"],
            "Gross Tax Liability": s["GrossTaxLiability"],
            "Refund / (Payable)": s["RefundOrPayable"],
        }
    s = ret.compute()
    tti = s["tti"]
    return {
        "Gross Total Income": s["GrossTotalIncome"],
        "VI-A Deductions": s["VIADeductions"],
        "Total Income": s["TotalIncome"],
        "Tax at Normal Rates": tti["TaxNormal"],
        "Tax at Special Rates": tti["TaxSpecialTotal"],
        "Rebate u/s 87A": tti["Rebate87A"],
        "Surcharge": tti["Surcharge"],
        "Cess": tti["Cess"],
        "Gross Tax Liability": tti["GrossTaxLiability"],
        "Interest 234A": s["IntrstPay234A"],
        "Interest 234B": s["IntrstPay234B"],
        "Fee 234F": s["Fee234F"],
        "Total Liability": s["TotalLiability"],
        "Refund / (Payable)": s["RefundDue"] if s["RefundDue"] else -s["BalTaxPayable"],
    }
