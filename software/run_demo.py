#!/usr/bin/env python3
"""Demo: a resident salaried taxpayer (Surat) filing ITR-1 AY 2026-27 in the
default new regime -> full computation + schema-valid CBDT JSON."""
import json, os, sys
from datetime import date

sys.path.insert(0, os.path.dirname(__file__))
from itr_filing.itr1 import (ITR1Return, PersonalInfo, FilingStatus, Salary,
                             TaxesPaid, TDSonSalary, Verification, BankAccount)
from itr_filing.house_property import HouseProperty
from itr_filing.json_builder import build_itr1_json, validate_against_schema

r = ITR1Return()
r.personal = PersonalInfo(
    first_name="Ramesh", last_name="Shah", pan="ABCPS1234K",
    aadhaar_number="123456789012", date_of_birth=date(1986, 8, 14),
    nature_of_employment="PE", primary_mobile="9876543210",
    primary_email="ramesh@example.com", father_name="Kantilal Shah")
r.personal.primary_address.flat_door_block = "B-12"
r.personal.primary_address.name_of_premises = "Shanti Residency"
r.personal.primary_address.road_street_post_office = "Adajan Road"
r.personal.primary_address.area_locality = "Adajan"
r.personal.primary_address.city = "Surat"
r.personal.primary_address.state = "11"          # 11 = Gujarat (schema list)
r.personal.primary_address.pin_code = "395009"

r.filing = FilingStatus(opt_out_new_regime="No", filing_sec_code=11)   # 139(1)
r.salary = Salary(salary_17_1=1400000)                                  # Rs 14 L
r.house_properties = [HouseProperty(property_type="Self Occupied")]
r.taxes_paid = TaxesPaid(tds_salary=[TDSonSalary(
    employer_name="Acme Technologies Pvt Ltd", tan="MUMC12345A", tds_amount=60000)])
r.bank = BankAccount(ifsc="SBIN0001234", account_no="30012345678",
                     bank_name="State Bank of India", is_for_refund="Yes")
r.verification = Verification(verification_date=date(2026, 7, 20))

s = r.compute_summary()
print("=== ITR-1 AY 2026-27 computation (new regime) ===")
for k in ("GrossTotIncome", "TotalChapterVIA", "TotalIncome", "TotalTaxPayable",
          "Rebate87A", "TaxPayableOnRebate", "EducationCess", "GrossTaxLiability",
          "FeeIncUS234F", "TotTaxPlusIntrstPay", "RefundOrPayable"):
    print(f"  {k:<22} = {s[k]:>12}")

payload = build_itr1_json(r)
schema = os.path.join(os.path.dirname(__file__), "..", "ITR-1",
                      "ITR-1_2026_Main_V1.1.json")
errs = validate_against_schema(payload, schema)
print(f"\nSchema validation against official CBDT schema: "
      f"{'PASS (0 errors)' if not errs else len(errs)}")
out = os.path.join(os.path.dirname(__file__), "sample_output_itr1.json")
json.dump(payload, open(out, "w"), indent=1)
print("CBDT JSON written to", out)
