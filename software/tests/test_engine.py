"""Golden-vector tests. Expected values hand-derived from the ported Excel/VBA
rules (sources cited in each module)."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import date
from itr_filing.tax_engine import (tax_old_regime, tax_new_regime, rebate_87a,
                                   education_cess, ltcg_112a_not_chargeable, calculate_age)
from itr_filing import chapter_via as via
from itr_filing.interest import interest_234a, interest_234b, fee_234f
from itr_filing.house_property import HouseProperty, total_hp_income
from itr_filing import constants as C


# ---------------------------------------------------------------- slabs
def test_new_regime_slabs_AY26_27():
    # [VBA:calcTaxPayableOnTINTR] 4/8/12/16/20/24L
    assert tax_new_regime(400000) == 0
    assert tax_new_regime(800000) == 20000
    assert tax_new_regime(900000) == 30000
    assert tax_new_regime(1200000) == 60000
    assert tax_new_regime(1600000) == 120000
    assert tax_new_regime(2000000) == 200000
    assert tax_new_regime(2400000) == 300000
    assert tax_new_regime(2500000) == 330000


def test_old_regime_slabs_by_age():
    # [VBA:calcTaxPayableOnTI]
    assert tax_old_regime(1000000, 40) == 112500      # <60
    assert tax_old_regime(1000000, 65) == 110000      # senior 60-79
    assert tax_old_regime(1000000, 82) == 100000      # super senior >=80: 5L exempt, 20% band
    assert tax_old_regime(1500000, 82) == 250000      # >10L: 30% + 100000
    assert tax_old_regime(250000, 40) == 0
    assert tax_old_regime(300000, 65) == 0
    assert tax_old_regime(500000, 82) == 0


# ---------------------------------------------------------------- 87A
def test_rebate_87a_new_regime_full():
    # TI <= 12L -> rebate = min(tax, 60000)  [AO177]
    tax = tax_new_regime(825000)            # 22500
    assert rebate_87a(tax, 825000, 825000, C.REGIME_NEW) == 22500


def test_rebate_87a_new_regime_marginal_relief():
    # TI = 12.3L -> tax 64500; relief caps effective tax at (TI-12L)=30000
    tax = tax_new_regime(1230000)           # 64500
    r = rebate_87a(tax, 1230000, 1230000, C.REGIME_NEW)
    assert r == 34500
    assert tax - r == 30000


def test_rebate_87a_old_regime():
    # tested on TI incl. 112A [AO177 BacValue=2 branch]
    tax = tax_old_regime(490000, 40)        # 12000
    assert rebate_87a(tax, 490000, 490000, C.REGIME_OLD) == 12000
    assert rebate_87a(tax, 510000, 510000, C.REGIME_OLD) == 0


def test_cess_4pct():
    assert education_cess(30000) == 1200    # [AO180]


def test_112a_not_chargeable():
    # [AO175] gain <= 1.25L -> whole gain not chargeable; > 1.25L -> 0
    assert ltcg_112a_not_chargeable(300000, 100000) == 0       # gain 2.0L > 1.25L
    assert ltcg_112a_not_chargeable(200000, 100000) == 100000  # gain 1.0L <= 1.25L
    assert ltcg_112a_not_chargeable(225000, 100000) == 125000  # gain exactly 1.25L
    assert ltcg_112a_not_chargeable(100000, 200000) == 0       # no gain


# ---------------------------------------------------------------- VI-A caps
def test_cap_80c():
    assert via.cap_80c(200000, 1000000) == 150000
    assert via.cap_80c(200000, 100000) == 100000        # capped by TI
    assert via.cap_80c(90000, 1000000) == 90000


def test_cap_80ccd1():
    # old regime private 10% of (net sal - perq); new regime 14% [AN121/AN137]
    assert via.cap_80ccd1(200000, 1000000, 0.10, C.REGIME_OLD, 1000000) == 100000
    assert via.cap_80ccd1(200000, 1000000, 0.10, C.REGIME_NEW, 1000000) == 140000
    assert via.cap_80ccd1(50000, 1000000, 0.10, C.REGIME_OLD, 1000000) == 50000


def test_cap_80tta_ttb():
    assert via.cap_80tta(15000, 1000000, 8000) == 8000    # capped by actual interest
    assert via.cap_80tta(15000, 1000000, 20000) == 10000  # capped at 10k
    assert via.cap_80ttb(60000, 1000000, 70000) == 50000


def test_cap_80gg_zeroed_AY26_27():
    # [AN152] literal MIN(...,0) zeroes 80GG this AY
    assert via.cap_80gg(50000, 0, 1000000, 0) == 0


def test_new_regime_blocks_old_only_deductions():
    assert via.cap_80dd(75000, C.REGIME_NEW) == 0
    assert via.cap_80dd(75000, C.REGIME_OLD) == 75000
    assert via.cap_80u(125000, C.REGIME_NEW) == 0
    assert via.cap_80gga(10000, C.REGIME_NEW) == 0


def test_total_income_rounded_to_10():
    assert via.round_total_income(825003) == 825000
    assert via.round_total_income(825005) == 825010
    assert via.round_total_income(-5) == 0


# ---------------------------------------------------------------- house property
def test_hp_let_out():
    hp = HouseProperty(property_type="Let Out", gross_annual_value=300000,
                       tax_paid_local_authorities=10000, interest_24b=50000)
    assert hp.annual_value() == 290000
    assert hp.standard_deduction_30pct() == 87000
    assert hp.income() == 153000


def test_hp_loss_cap_old_regime():
    hp = HouseProperty(property_type="Self Occupied", interest_24b=300000)
    assert total_hp_income([hp]) == -200000     # [HP!I76] MAX(-200000, ...)


# ---------------------------------------------------------------- interest & fees
def test_fee_234f():
    assert fee_234f(400000, True) == 1000
    assert fee_234f(600000, True) == 5000
    assert fee_234f(600000, False) == 0


def test_interest_234b():
    # NTL 50k, nothing paid, verification Sep-2026 -> 6 months
    assert interest_234b(50000, 0, 0, 0, date(2026, 9, 15), 40) == 3000
    # below 10k threshold -> no interest
    assert interest_234b(9000, 0, 0, 0, date(2026, 9, 15), 40) == 0
    # advance tax >= 90% of (NTL - TDS - TCS) -> no interest
    assert interest_234b(100000, 90000, 0, 0, date(2026, 9, 15), 40) == 0
    # senior exempt
    assert interest_234b(50000, 0, 0, 0, date(2026, 9, 15), 65) == 0


def test_interest_234a():
    # due 31/07/2026, filed 15/09/2026 -> 2 months (Aug part + Sept)
    assert interest_234a(100000, date(2026, 7, 31), date(2026, 9, 15)) == 2000
    assert interest_234a(100000, date(2026, 7, 31), date(2026, 7, 31)) == 0


def test_age_calc():
    assert calculate_age(date(1986, 8, 14), date(2026, 7, 31)) == 39
    assert calculate_age("14/08/1956", date(2026, 7, 31)) == 69
