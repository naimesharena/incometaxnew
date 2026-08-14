"""Tests for the final backlog modules: 115AD proviso, relief 89, 244A,
updated-return additional tax, storage/prefill."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import date
from itr_filing.schedule_115ad import FPIEquityTransaction, Schedule115AD
from itr_filing.relief89 import relief_89
from itr_filing.interest import refund_interest_244a, additional_tax_updated_return_139_8a
from itr_filing.storage import apply_prefill, save_draft, load_draft, list_drafts
from itr_filing import constants as C


# ------------------------------------------------------------- 115AD
def test_115ad_grandfathering_and_buckets():
    t1 = FPIEquityTransaction(acquired_on_or_before_31jan2018=True, shares=1000,
                              sale_price_per_share=500, cost_per_share=100,
                              fmv_per_share_31jan2018=300,
                              transfer_date=date(2024, 5, 1))       # before 23/07/24
    # FV = 500000; cost = MAX(100000, MIN(300000, 500000)) = 300000; gain 200000
    assert t1.full_value_consideration() == 500000
    assert t1.cost_of_acquisition() == 300000
    assert t1.gain() == 200000
    t2 = FPIEquityTransaction(acquired_on_or_before_31jan2018=False, shares=1000,
                              sale_price_per_share=400, cost_per_share=350,
                              transfer_date=date(2025, 1, 15))      # after 23/07/24
    assert t2.gain() == 50000
    sch = Schedule115AD(transactions=[t1, t2])
    s = sch.summary()
    assert s["GainsTransferBefore23Jul2024"] == 200000
    assert s["GainsTransferOnAfter23Jul2024"] == 50000
    assert s["LTCGUs112A_115ADProviso"] == 250000


def test_115ad_tax_special_rates():
    t = FPIEquityTransaction(acquired_on_or_before_31jan2018=False, shares=10000,
                             sale_price_per_share=200, cost_per_share=100,
                             transfer_date=date(2025, 2, 1))
    s = Schedule115AD(transactions=[t]).summary()
    # gain 10L; 112A exempt 1.25L -> taxable 8.75L @ 12.5% = 109375
    assert s["TaxableAfter112AExempt"] == 875000
    assert s["TaxAtSpecialRates"] == 109375


# ------------------------------------------------------------- relief 89
def test_relief_89_new_regime():
    # arrears 3L; CY income 9L; PY income 4L
    # CY delta: tax(12L)-tax(9L) = 60000 - 30000 = 30000
    # PY delta: tax(7L)-tax(4L) = 15000 - 0 = 15000
    r = relief_89(300000, 900000, 400000, regime=C.REGIME_NEW)
    assert r == 15000


def test_relief_89_no_arrears():
    assert relief_89(0, 1000000, 500000) == 0


def test_relief_89_never_negative():
    # prior-year bracket jump bigger than current-year effect
    r = relief_89(100000, 500000, 350000, regime=C.REGIME_NEW)
    assert r >= 0


# ------------------------------------------------------------- 244A & 139(8A)
def test_refund_interest_244a():
    assert refund_interest_244a(100000, 6) == 3000      # 0.5% x 6 months
    assert refund_interest_244a(0, 6) == 0
    assert refund_interest_244a(100000, 0) == 0


def test_updated_return_additional_tax():
    assert additional_tax_updated_return_139_8a(100000, 1) == 25000
    assert additional_tax_updated_return_139_8a(100000, 2) == 50000
    assert additional_tax_updated_return_139_8a(100000, 3) == 0


# ------------------------------------------------------------- storage/prefill
def test_draft_roundtrip(tmp_path, monkeypatch):
    import itr_filing.storage as st
    monkeypatch.setattr(st, "DRAFT_DIR", str(tmp_path))
    save_draft("ITR-1", "ABCPS1234K", {"gross_salary": 1400000})
    d = load_draft("ITR-1", "ABCPS1234K")
    assert d["input"]["gross_salary"] == 1400000
    assert d["form"] == "ITR-1"
    assert len(list_drafts()) == 1


def test_prefill_application():
    prefill = {"ITR": {"ITR1": {
        "PersonalInfo": {"AssesseeName": {"FirstName": "Asha", "MiddleName": "",
                                           "SurNameOrOrgName": "Patel"},
                          "PAN": "ABCPP1234K", "DOB": "1990-01-01"},
        "TDSonSalaries": {"TotalTDSonSalaries": 75000, "TDSonSalary": [
            {"EmployerOrDeductorOrCollectDetl": {"TAN": "MUMC12345A",
             "EmployerOrDeductorOrCollecterName": "Acme Ltd"},
             "IncChrgSal": 0, "TotalTDSSal": 75000}]},
        "ScheduleS": {"TotalGrossSalary": 1500000},
        "TaxPayments": {"TotalTaxPayments": 20000,
                        "TaxPayment": [{"BSRCode": "1234567", "DateDep": "2026-03-01",
                                        "SrlNoOfChaln": 1, "Amt": 20000}]},
    }}}
    out = apply_prefill(prefill)
    assert out["inputs"]["pan"] == "ABCPP1234K"
    assert out["inputs"]["first_name"] == "Asha"
    assert out["inputs"]["tds_amount"] == 75000
    assert out["inputs"]["gross_salary"] == 1500000
    assert out["inputs"]["advance_tax"] == 20000
    assert "TDSonSalaries" in out["prefilled_sections"]
    assert "ScheduleS" in out["prefilled_sections"]


def test_prefill_multiple_tds_note():
    prefill = {"ITR": {"ITR1": {"TDSonSalaries": {"TDSonSalary": [
        {"EmployerOrDeductorOrCollectDetl": {"TAN": "MUMC12345A",
         "EmployerOrDeductorOrCollecterName": "A"}, "TotalTDSSal": 10000},
        {"EmployerOrDeductorOrCollectDetl": {"TAN": "MUMC12346A",
         "EmployerOrDeductorOrCollecterName": "B"}, "TotalTDSSal": 5000}]}}}}
    out = apply_prefill(prefill)
    assert out["inputs"]["tds_amount"] == 10000
    assert any("additional" in n for n in out["notes"])
