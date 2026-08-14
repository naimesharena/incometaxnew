"""ITR-4 JSON builder - strict validation against the official CBDT schema."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import date
from itr_filing.itr4 import ITR4Return, Business44AD, ENTITY_HUF
from itr_filing.itr1 import PersonalInfo, BankAccount
from itr_filing.itr4_json import build_itr4_json, validate_against_schema

REPO = os.path.join(os.path.dirname(__file__), "..", "..")
SCHEMA = os.path.join(REPO, "ITR-4", "ITR-4_2026_Main_V1.1.json")


def make_itr4():
    r = ITR4Return(entity="I", dob=date(1985, 1, 1))
    r.personal = PersonalInfo(first_name="Mehul", last_name="Patel", pan="ABCPP1234K",
                              aadhaar_number="123456789012", date_of_birth=date(1985, 1, 1),
                              nature_of_employment="NA", primary_mobile="9876543210",
                              primary_email="m@x.com", father_name="Kantilal Patel")
    r.personal.primary_address.flat_door_block = "101"
    r.personal.primary_address.area_locality = "Ring Road"
    r.personal.primary_address.city = "Surat"
    r.personal.primary_address.state = "11"
    r.personal.primary_address.pin_code = "395002"
    r.b44ad = Business44AD(turnover_digital=3000000, turnover_other=2000000,
                           income_digital=180000, income_other=160000)
    r.bank = BankAccount(ifsc="SBIN0001234", account_no="30012345678",
                         bank_name="State Bank of India", is_for_refund="Yes")
    r.tds = 10000
    r.verification_date = date(2026, 8, 20)
    return r


def test_itr4_json_schema_valid():
    r = make_itr4()
    payload = build_itr4_json(r)
    errs = validate_against_schema(payload, SCHEMA)
    assert errs == [], errs


def test_itr4_json_values():
    r = make_itr4()
    payload = build_itr4_json(r)
    i4 = payload["ITR"]["ITR4"]
    assert i4["Form_ITR4"]["FormName"] == "ITR-4"
    assert i4["Form_ITR4"]["AssessmentYear"] == "2026"
    assert i4["FilingStatus"]["ItrFilingDueDate"] == "2026-08-31"   # ITR-4 due date
    assert i4["PersonalInfo"]["Status"] == "I"
    bp = i4["ScheduleBP"]["PersumptiveInc44AD"]
    assert bp["GrsTotalTrnOver"] == 5000000
    assert bp["TotPersumptiveInc44AD"] == 340000
    assert i4["IncomeDeductions"]["IncomeFromBusinessProf"] == 340000
    # 3.4L < 4L -> zero tax, full TDS refund
    tc = i4["TaxComputation"]
    assert tc["TotalTaxPayable"] == 0
    assert i4["Refund"]["RefundDue"] == 10000


def test_itr4_huf_status_and_bank_required():
    r = make_itr4()
    r.entity = ENTITY_HUF
    payload = build_itr4_json(r)
    assert payload["ITR"]["ITR4"]["PersonalInfo"]["Status"] == "H"
    r2 = make_itr4()
    r2.bank = BankAccount()
    try:
        build_itr4_json(r2)
        assert False, "should require bank"
    except ValueError:
        pass
