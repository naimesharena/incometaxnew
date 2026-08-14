"""ITR-4 engine tests (TaxCalc port + presumptive taxation rules)."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import date
from itr_filing.itr4 import (ITR4Return, Business44AD, Profession44ADA,
                             Business44AE, Vehicle44AE, ENTITY_FIRM, ENTITY_HUF,
                             ENTITY_INDIVIDUAL)
from itr_filing import constants as C


def make(entity=ENTITY_INDIVIDUAL, ti_income=2000000, regime=C.REGIME_NEW,
         age_dob=date(1985, 1, 1)):
    r = ITR4Return(entity=entity, dob=age_dob, regime=regime)
    r.b44ad = Business44AD(turnover_digital=20000000, turnover_other=10000000,
                           income_digital=ti_income // 2, income_other=ti_income // 2)
    return r


def test_44ad_minimum_rates():
    b = Business44AD(turnover_digital=20000000, turnover_other=10000000,
                     income_digital=1200000, income_other=800000)
    assert b.min_income() == 0.06 * 20000000 + 0.08 * 10000000
    assert b.validation_errors() == []
    b2 = Business44AD(turnover_digital=20000000, turnover_other=0,
                      income_digital=1000000, income_other=0)
    assert any("6%" in e for e in b2.validation_errors())


def test_44ae_minimums():
    ae = Business44AE(vehicles=[Vehicle44AE(heavy=True, units_or_tons=10, months=12)],
                      declared_income=119999)
    assert ae.min_income() == 1000 * 10 * 12
    assert ae.validation_errors()
    ae.declared_income = 120000
    assert ae.validation_errors() == []


def test_new_regime_zero_tax_upto_4L():
    r = make(ti_income=300000)          # Rs 3 lakh
    s = r.compute_summary()
    assert s["TotalIncome"] == 300000
    assert s["TotalTaxPayable"] == 0


def test_new_regime_87a_marginal_relief():
    r = make(ti_income=1230000)         # Rs 12.3 lakh
    s = r.compute_summary()
    assert s["TotalTaxPayable"] == 64500
    assert s["Rebate87A"] == 34500
    assert s["EducationCess"] == 1200


def test_firm_flat_30pct():
    r = make(entity=ENTITY_FIRM, ti_income=1000000)
    s = r.compute_summary()
    assert s["TotalTaxPayable"] == 300000
    assert s["Rebate87A"] == 0              # firms get no 87A
    assert s["EducationCess"] == 12000      # 4% of 300000


def test_huf_old_regime_slabs():
    r = make(entity=ENTITY_HUF, ti_income=1000000, regime=C.REGIME_OLD)
    assert r.d1_total_tax_payable() == 112500


def test_surcharge_marginal_relief_old_regime():
    # [TaxCalc B32-B36] TI 50.1L old regime individual
    r = make(ti_income=5010000, regime=C.REGIME_OLD)
    assert r.d1_total_tax_payable() == 1315500
    assert r.surcharge() == 7000            # relief caps the extra burden


def test_surcharge_10pct_band():
    r = make(ti_income=6000000, regime=C.REGIME_OLD)
    assert r.surcharge() == 161250          # plain 10%, no relief needed


def test_total_income_cap_50L():
    r = make(ti_income=51000000)
    assert any("50 Lakh" in e for e in r.validation_errors())


def test_44ada_rules():
    p = Profession44ADA(gross_receipts=4000000, income=1900000)
    assert any("50%" in e for e in p.validation_errors())
    p.income = 2000000
    assert p.validation_errors() == []
    p.gross_receipts = 8000000
    assert any("75 Lakh" in e for e in p.validation_errors())
