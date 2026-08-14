"""ITR-2 model + JSON builder tests (strict schema validation)."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import date
from itr_filing.itr2 import ITR2Return
from itr_filing.itr1 import PersonalInfo, BankAccount, TaxesPaid, TDSonSalary
from itr_filing.house_property import HouseProperty
from itr_filing.capital_gains import CapitalAsset
from itr_filing.carry_forward import LossEntry
from itr_filing.itr2_json import build_itr2_json
from itr_filing.json_builder import validate_against_schema
from itr_filing import constants as C

REPO = os.path.join(os.path.dirname(__file__), "..", "..")
SCHEMA = os.path.join(REPO, "ITR-2", "ITR-2_2026_Main_V1.1.json")


def make_itr2():
    r = ITR2Return()
    r.personal = PersonalInfo(first_name="Nisha", last_name="Desai", pan="ABCPS1234K",
                              aadhaar_number="123456789012", date_of_birth=date(1988, 3, 12),
                              primary_mobile="9812345678", primary_email="n@x.com",
                              father_name="Raman Desai")
    r.personal.primary_address.flat_door_block = "A-5"
    r.personal.primary_address.area_locality = "Vesu"
    r.personal.primary_address.city = "Surat"
    r.personal.primary_address.state = "11"
    r.personal.primary_address.pin_code = "395007"
    r.gross_salary = 1600000
    r.exempt_allowances = 100000
    r.cg_assets = [CapitalAsset(asset_class="securities", listed_stt_paid=True,
                                acquisition_date=date(2025, 11, 1),
                                transfer_date=date(2026, 2, 1),
                                sale_consideration=300000, cost_of_acquisition=200000)]
    r.via_total = 150000
    r.taxes_paid = TaxesPaid(tds_salary=[TDSonSalary(employer_name="Acme",
                                                     tan="MUMC12345A", tds_amount=80000)])
    r.bank = BankAccount(ifsc="SBIN0001234", account_no="30012345678", bank_name="SBI")
    r.verification_date = date(2026, 7, 25)
    return r


def test_itr2_compute_pipeline():
    r = make_itr2()
    s = r.compute()
    # salary: 16L - 1L exempt - 75k std ded = 14,25,000
    assert s["Salaries"] == 1425000
    # CG: STCG20 band +1,00,000 (4 months holding, STT-paid)
    assert s["cg_summary"]["STCG20"] == 100000
    assert s["IncChargeTaxSplRate111A112"] == 100000
    # GTI = 14,25,000 + 1,00,000 = 15,25,000; VIA 1.5L -> TI 13,75,000
    assert s["GrossTotalIncome"] == 1525000
    assert s["TotalIncome"] == 1375000
    # tax: normal 14.25L? normal TI = 14,25,000 -> new regime = 1,15,000... check
    tti = s["tti"]
    assert tti["TaxNormal"] == 93750           # (14.25L-12L)*0.15+60000
    assert tti["TaxSpecialTotal"] == 20000     # 20% on STCG20 1L
    assert tti["Rebate87A"] == 0               # TI excl 112A = 15.25L > 12L


def test_itr2_hp_loss_cyla_and_cfl():
    r = make_itr2()
    r.house_properties = [HouseProperty(property_type="Self Occupied", interest_24b=300000)]
    s = r.compute()
    # HP loss capped 2L set off against salary; remaining 0 in new regime
    assert s["IncomeFromHP"] == 0
    assert s["Salaries"] == 1425000                       # head income pre-set-off
    assert s["bfla"].income_after_setoff["SAL"] == 1225000  # 14.25L - 2L HP loss
    r.regime = C.REGIME_OLD
    s2 = r.compute()
    assert s2["cyla"].hp_loss_remaining == 0   # fully absorbed (2L) anyway


def test_itr2_bfla_from_cfl():
    r = make_itr2()
    r.cfl.add(LossEntry(ay="2024-25", stcg=50000))
    s = r.compute()
    # brought-forward STCL absorbs current STCG20 1L
    assert s["bfla"].setoff_by_category["stcg"] == 50000
    assert s["IncChargeTaxSplRate111A112"] == 50000


def test_itr2_json_schema_valid():
    payload = build_itr2_json(make_itr2())
    errs = validate_against_schema(payload, SCHEMA)
    assert errs == [], errs


def test_itr2_json_values():
    payload = build_itr2_json(make_itr2())
    i2 = payload["ITR"]["ITR2"]
    assert i2["Form_ITR2"]["FormName"] == "ITR-2"
    assert i2["PartA_GEN1"]["PersonalInfo"]["PAN"] == "ABCPS1234K"
    assert i2["PartB-TI"]["Salaries"] == 1425000
    assert i2["PartB-TI"]["TotalIncome"] == 1375000
    assert i2["PartB_TTI"]["ComputationOfTaxLiability"]["TaxPayableOnTI"]["TaxAtSpecialRates"] == 20000
    # refund: paid 80000 vs liability
    assert i2["PartB_TTI"]["Refund"]["RefundDue"] >= 0
