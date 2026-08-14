"""Tests: AMT (115JC) + AMT credit (115JD), CG exemptions (54 family), DTAA."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from itr_filing.amt import (amt_adjusted_total_income, amt_payable, amt_cess,
                            AmtcTracker, AmtcEntry)
from itr_filing.cg_exemptions import (exemption_54ec, exemption_54f, exemption_54,
                                      two_house_option_allowed)
from itr_filing.dtaa import relief_90, relief_91
from itr_filing import constants as C


# ------------------------------------------------------------------ AMT
def test_amt_not_applicable_new_regime():
    assert amt_adjusted_total_income(3000000, 150000, 0, 0, C.REGIME_NEW) == 0
    assert amt_payable(5000000, True, C.REGIME_NEW) == 0


def test_amt_adjusted_income_and_tax():
    adj = amt_adjusted_total_income(3000000, 150000, 0, 0, C.REGIME_OLD)
    assert adj == 3150000
    assert amt_payable(adj, True, C.REGIME_OLD) == 582750   # 18.5%, no surcharge


def test_amt_threshold_and_no_via():
    assert amt_payable(2000000, True, C.REGIME_OLD) == 0    # not > 20L
    assert amt_payable(3150000, False, C.REGIME_OLD) == 0   # no VI-A claimed


def test_amt_surcharge_band():
    adj = 6000000
    tax = amt_payable(adj, True, C.REGIME_OLD)
    assert tax == 1110000 + 111000                          # 10% surcharge > 50L
    assert amt_cess(tax) == 48840                           # 4%


# ------------------------------------------------------------------ AMTC
def test_amtc_credit_arising():
    t = AmtcTracker()
    r = t.current_year(amt_tax=500000, regular_tax=300000,
                       current_ay="2026-27", regime=C.REGIME_OLD)
    assert r["credit_arising"] == 200000
    assert r["utilised"] == 0


def test_amtc_credit_utilisation():
    t = AmtcTracker([AmtcEntry(ay="2024-25", gross_credit=200000)])
    r = t.current_year(amt_tax=100000, regular_tax=400000,
                       current_ay="2026-27", regime=C.REGIME_OLD)
    assert r["brought_forward"] == 200000
    assert r["utilised"] == 200000
    assert r["carried_forward"] == 0


def test_amtc_expiry_15_years():
    t = AmtcTracker([AmtcEntry(ay="2010-11", gross_credit=500000)])
    assert t.brought_forward_balance("2026-27") == 0        # 16 AYs old


def test_amtc_new_regime_zeroed():
    t = AmtcTracker([AmtcEntry(ay="2024-25", gross_credit=200000)])
    r = t.current_year(0, 400000, "2026-27", C.REGIME_NEW)
    assert r["utilised"] == 0 and r["carried_forward"] == 0


# ------------------------------------------------------------------ 54 family
def test_54ec_cap_50_lakh():
    assert exemption_54ec(6000000) == 5000000
    assert exemption_54ec(4000000) == 4000000


def test_54f_pro_rata():
    # gain 20L, net consideration 30L, invested 15L -> half the gain
    assert exemption_54f(2000000, 3000000, 1500000) == 1000000
    # full investment of net consideration -> full gain
    assert exemption_54f(2000000, 3000000, 3000000) == 2000000
    # investment capped at 10 crore
    assert exemption_54f(200000000, 2000000000, 5000000000) == 100000000


def test_54_min_gain_cost_and_two_house_rule():
    assert exemption_54(1500000, 1200000) == 1200000
    assert exemption_54(1500000, 2000000) == 1500000
    assert two_house_option_allowed(20000000) is True
    assert two_house_option_allowed(20000001) is False
    # second house claim invalid when gain > 2Cr
    assert exemption_54(25000000, 10000000, second_house_used=True) == 0


# ------------------------------------------------------------------ DTAA
def test_relief_90_lower_of_two():
    # Indian incremental tax 60k, foreign tax 100k -> relief 60k
    assert relief_90(100000, 300000, 240000) == 60000
    # foreign tax lower
    assert relief_90(40000, 300000, 240000) == 40000


def test_relief_91_lower_rate():
    # Indian rate 30%, foreign 20%: relief at 20% on the doubly-taxed income
    assert relief_91(45000, 300000, 240000, 0.30, 0.20) == 40000
    # foreign tax itself lower than computed credit
    assert relief_91(30000, 300000, 240000, 0.30, 0.20) == 30000
