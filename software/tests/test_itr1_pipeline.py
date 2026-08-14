"""End-to-end ITR-1 pipeline tests + official-schema validation."""
import sys, os, json
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import date
from itr_filing.itr1 import (ITR1Return, PersonalInfo, FilingStatus, Salary,
                             OtherSources, ChapterVIA, TaxesPaid, TDSonSalary,
                             Verification, BankAccount)
from itr_filing.house_property import HouseProperty
from itr_filing import constants as C

REPO = os.path.join(os.path.dirname(__file__), "..", "..")


def make_basic(salary_amt=900000, opt_out="No", dob=date(1986, 8, 14)):
    r = ITR1Return()
    r.personal = PersonalInfo(first_name="Ramesh", last_name="Shah",
                              pan="ABCPS1234K", aadhaar_number="123456789012",
                              date_of_birth=dob, nature_of_employment="Salaried",
                              primary_mobile="9876543210", primary_email="r@x.com")
    r.personal.primary_address.state = "24-GUJARAT"
    r.personal.primary_address.pin_code = "395007"
    r.filing = FilingStatus(opt_out_new_regime=opt_out, filing_sec_code=11)
    r.salary = Salary(salary_17_1=salary_amt)
    r.verification = Verification(verification_date=date(2026, 7, 20))
    return r


def test_salary_9L_new_regime_zero_tax():
    r = make_basic(900000)
    s = r.compute_summary()
    assert r.income_from_salaries() == 825000          # 9L - 75k std ded
    assert s["GrossTotIncome"] == 825000
    assert s["TotalTaxPayable"] == 22500
    assert s["Rebate87A"] == 22500                     # TI <= 12L
    assert s["GrossTaxLiability"] == 0


def test_salary_12L_new_regime_marginal_relief():
    r = make_basic(1275000)
    s = r.compute_summary()
    assert r.income_from_salaries() == 1200000         # 12.75L - 75k
    assert s["TotalTaxPayable"] == 60000
    assert s["Rebate87A"] == 60000
    assert s["GrossTaxLiability"] == 0                 # exactly at 12L


def test_salary_13L5_new_regime_marginal_relief_partial():
    r = make_basic(1305000)                            # TI = 12.3L
    s = r.compute_summary()
    assert s["TotalTaxPayable"] == 64500
    assert s["Rebate87A"] == 34500
    assert s["TaxPayableOnRebate"] == 30000
    assert s["EducationCess"] == 1200
    assert s["GrossTaxLiability"] == 31200


def test_old_regime_senior_with_deductions():
    r = make_basic(1200000, opt_out="Yes", dob=date(1960, 1, 1))  # age 66
    r.chapter_via = ChapterVIA(s80c=150000, s80d_schedule_amount=25000, user_total=175000)
    s = r.compute_summary()
    # old regime std deduction 50k -> net taxable path
    assert r.salary.deduction_16ia(C.REGIME_OLD) == 50000
    assert r.income_from_salaries() == 1150000
    assert s["TotalChapterVIA"] == 175000
    assert s["TotalIncome"] == 975000
    # senior slabs: (975000-500000)*0.2 + 10000 = 105000
    assert s["TotalTaxPayable"] == 105000
    assert s["Rebate87A"] == 0                         # TI > 5L


def test_new_regime_blocks_80dd_80u():
    r = make_basic(1500000)
    r.chapter_via = ChapterVIA(s80dd_schedule_amount=75000, s80u_schedule_amount=125000,
                               user_total=200000)
    s = r.compute_summary()
    assert s["ChapterVIAComponents"]["80DD"] == 0
    assert s["ChapterVIAComponents"]["80U"] == 0
    assert s["TotalChapterVIA"] == 0


def test_hp_letout_and_setoff():
    r = make_basic(1000000)
    r.house_properties = [HouseProperty(property_type="Let Out", gross_annual_value=300000,
                                        tax_paid_local_authorities=10000, interest_24b=50000)]
    s = r.compute_summary()
    assert r.hp_income() == 153000
    assert s["GrossTotIncome"] == 925000 + 153000 - 1000000 + 1000000  # = 1078000
    assert s["GrossTotIncome"] == 1078000


def test_refund_case():
    r = make_basic(1200000)
    r.taxes_paid = TaxesPaid(tds_salary=[TDSonSalary(employer_name="Acme Ltd",
                                                     tan="MUMC12345A", tds_amount=40000)])
    r.bank = BankAccount(ifsc="SBIN0001234", account_no="1234567890",
                         bank_name="SBI", is_for_refund="Yes")
    s = r.compute_summary()
    assert s["GrossTaxLiability"] == 0
    assert s["RefundOrPayable"] == 40000


def test_ltcg_112a_eligibility_included_in_gti_new():
    r = make_basic(900000)
    r.ltcg_112a.total_sale_consideration = 200000
    r.ltcg_112a.total_cost_of_acquisition = 120000    # gain 80k <= 1.25L
    s = r.compute_summary()
    assert s["GrossTotIncome"] == 825000               # 112A excluded from GTI
    assert s["GrossTotIncomeIncLTCG112A"] == 905000    # included in _New
    assert s["TotalIncomeIncl112A"] == 905000


def test_json_builds_and_matches_schema():
    """STRICT: output must be 100% valid against the official CBDT schema."""
    from itr_filing.json_builder import build_itr1_json, validate_against_schema
    r = make_basic(900000)
    r.personal.father_name = "Kantilal Shah"
    r.personal.nature_of_employment = "PE"
    r.personal.primary_address.flat_door_block = "B-12"
    r.personal.primary_address.area_locality = "Adajan"
    r.personal.primary_address.city = "Surat"
    r.personal.primary_address.state = "11"
    r.personal.primary_address.pin_code = "395009"
    r.taxes_paid = TaxesPaid(tds_salary=[TDSonSalary(employer_name="Acme Ltd",
                                                     tan="MUMC12345A", tds_amount=1000)])
    r.bank = BankAccount(ifsc="SBIN0001234", account_no="30012345678",
                         bank_name="State Bank of India", is_for_refund="Yes")
    payload = build_itr1_json(r)
    assert payload["ITR"]["ITR1"]["Form_ITR1"]["FormName"] == "ITR-1"
    assert payload["ITR"]["ITR1"]["Form_ITR1"]["AssessmentYear"] == "2026"
    assert payload["ITR"]["ITR1"]["ITR1_TaxComputation"]["TotalTaxPayable"] == 22500
    schema_path = os.path.join(REPO, "ITR-1", "ITR-1_2026_Main_V1.1.json")
    errs = validate_against_schema(payload, schema_path)
    for e in errs[:20]:
        print("SCHEMA:", e)
    assert errs == [], errs


def test_json_with_house_property_schema_valid():
    """Regression: PropertyDetails/Rentdetails field names per official schema."""
    from itr_filing.json_builder import build_itr1_json, validate_against_schema
    r = make_basic(1000000)
    r.personal.father_name = "Kantilal Shah"
    r.personal.primary_address.flat_door_block = "B-12"
    r.personal.primary_address.area_locality = "Adajan"
    r.personal.primary_address.city = "Surat"
    r.personal.primary_address.state = "11"
    r.personal.primary_address.pin_code = "395009"
    r.house_properties = [HouseProperty(property_type="Let Out", gross_annual_value=300000,
                                        tax_paid_local_authorities=10000, interest_24b=50000,
                                        address="12 MG Road", city="Surat", state="11",
                                        pincode=395001)]
    r.bank = BankAccount(ifsc="SBIN0001234", account_no="30012345678",
                         bank_name="State Bank of India", is_for_refund="Yes")
    payload = build_itr1_json(r)
    errs = validate_against_schema(payload, os.path.join(REPO, "ITR-1",
                                                         "ITR-1_2026_Main_V1.1.json"))
    assert errs == [], errs


def test_validation_rules():
    from itr_filing.validation import validate_itr1
    r = make_basic(900000)
    r.personal.aadhaar_number = "12345"               # invalid
    r.ltcg_112a.total_sale_consideration = 500000
    r.ltcg_112a.total_cost_of_acquisition = 100000    # gain 4L > 1.25L
    out = validate_itr1(r)
    assert any("Aadhaar" in e for e in out["errors"])
    assert any("112A" in e for e in out["errors"])
