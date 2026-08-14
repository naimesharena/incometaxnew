"""ITR-3 tests: business income (P&L + depreciation), pipeline, strict schema."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import date
from itr_filing.itr3 import ITR3Return
from itr_filing.business import (BusinessPL, DepreciationBlock,
                                 SpeculationBusiness, SpecifiedBusiness,
                                 compute_business_income)
from itr_filing.itr1 import PersonalInfo, BankAccount
from itr_filing.itr3_json import build_itr3_json
from itr_filing.json_builder import validate_against_schema

REPO = os.path.join(os.path.dirname(__file__), "..", "..")
SCHEMA = os.path.join(REPO, "ITR-3", "ITR-3_2026_Main_V1.1.json")


def test_depreciation_block_180_day_rule():
    b = DepreciationBlock(rate=0.15, opening_wdv=800000,
                          additions_full=200000, additions_half=100000)
    # (8L + 2L) x 15% + 1L x 7.5% = 150000 + 7500
    assert b.depreciation() == 157500


def test_business_pl_income():
    pl = BusinessPL(net_profit_as_per_pl=1200000, inadmissible_expenses=50000,
                    depreciation_as_per_pl=100000,
                    depreciation_blocks=[DepreciationBlock(rate=0.15,
                                                           opening_wdv=800000,
                                                           additions_full=200000,
                                                           additions_half=100000)])
    assert pl.income_before_depreciation() == 1350000   # +50k +100k add-back
    assert pl.income() == 1192500                       # - 157500 IT depreciation


def make_itr3():
    r = ITR3Return()
    r.personal = PersonalInfo(first_name="Kiran", last_name="Shah", pan="ABCPS9999K",
                              date_of_birth=date(1982, 5, 20),
                              primary_mobile="9898989898", primary_email="k@x.com",
                              father_name="Bharat Shah")
    r.personal.primary_address.flat_door_block = "7"
    r.personal.primary_address.area_locality = "Katargam"
    r.personal.primary_address.city = "Surat"
    r.personal.primary_address.state = "11"
    r.personal.primary_address.pin_code = "395004"
    r.businesses = [BusinessPL(net_profit_as_per_pl=1200000,
                               inadmissible_expenses=50000,
                               depreciation_as_per_pl=100000,
                               depreciation_blocks=[DepreciationBlock(
                                   rate=0.15, opening_wdv=800000,
                                   additions_full=200000, additions_half=100000)])]
    r.bank = BankAccount(ifsc="SBIN0001234", account_no="30012345678", bank_name="SBI")
    r.verification_date = date(2026, 7, 20)
    return r


def test_itr3_pipeline_business_income():
    r = make_itr3()
    s = r.compute()
    assert s["BusinessIncome"] == 1192500
    assert s["GrossTotalIncome"] == 1192500
    assert s["TotalIncome"] == 1192500
    # new regime slabs on 11.925L: (11.925L-8L)*0.10 + 20000 = 59250
    assert s["tti"]["TaxNormal"] == 59250


def test_itr3_speculation_isolated():
    r = make_itr3()
    r.speculation = [SpeculationBusiness(profit=-200000)]
    s = r.compute()
    # speculation loss cannot offset business profit
    assert s["BusinessIncome"] == 1192500
    assert s["SpeculationIncome"] == 0       # floored for income
    assert s["cyla"].income_after_setoff["SPEC"] == 0


def test_itr3_json_schema_valid():
    payload = build_itr3_json(make_itr3())
    errs = validate_against_schema(payload, SCHEMA)
    assert errs == [], errs


def test_itr3_json_values():
    payload = build_itr3_json(make_itr3())
    i3 = payload["ITR"]["ITR3"]
    assert i3["Form_ITR3"]["FormName"] == "ITR-3"
    assert i3["PartA_GEN1"]["FilingStatus"]["ItrFilingDueDate"] == "2026-08-31"
    bio = i3["ITR3ScheduleBP"]["BusinessIncOthThanSpec"]
    assert bio["ProfBfrTaxPL"] == 1200000
    assert bio["NetPLAftAdjBusOthThanSpec"] == 1192500
    assert bio["DepreciationAllowITAct32"]["TotDeprAllowITAct"] == 157500
    assert i3["PartB-TI"]["ProfBusGain"]["TotProfBusGain"] == 1192500
    assert i3["PartB-TI"]["TotalIncome"] == 1192500
