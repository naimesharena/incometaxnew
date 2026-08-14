"""Part B TI/TTI final computation tests (ITR-2/3 engine capstone)."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from itr_filing.ti_tti import TITTIInput, compute_ti_tti, special_rate_tax, surcharge
from itr_filing import constants as C


def test_new_regime_12l_zero_after_rebate():
    r = compute_ti_tti(TITTIInput(total_income_normal=1200000))
    assert r["TaxNormal"] == 60000
    assert r["Rebate87A"] == 60000
    assert r["GrossTaxLiability"] == 0


def test_new_regime_marginal_relief_and_cess():
    r = compute_ti_tti(TITTIInput(total_income_normal=1230000))
    assert r["TaxNormal"] == 64500
    assert r["Rebate87A"] == 34500
    assert r["Cess"] == 1200
    assert r["GrossTaxLiability"] == 31200


def test_special_rate_bands():
    sp = special_rate_tax({"STCG20": 100000, "LTCG125": 200000})
    assert sp["STCG20"] == 20000                    # 20%
    assert sp["LTCG125"] == 9375                    # (2L - 1.25L) x 12.5%


def test_cg_bands_inside_full_computation():
    r = compute_ti_tti(TITTIInput(total_income_normal=1000000,
                                  cg_net_by_band={"STCG20": 100000,
                                                  "LTCG125": 200000}))
    assert r["TaxNormal"] == 40000                  # 10L new regime: 20000+20000
    assert r["TaxSpecialTotal"] == 29375
    # rebate tested on TI excl. 112A (11L <= 12L) per utility AO177 logic
    assert r["Rebate87A"] == 40000
    assert r["GrossTaxLiability"] == 30550          # special 29375 + cess 1175


def test_surcharge_new_regime_capped_25pct():
    assert surcharge(60000000, 1000000, C.REGIME_NEW) == 250000
    assert surcharge(60000000, 1000000, C.REGIME_OLD) == 370000


def test_full_surcharge_computation():
    r = compute_ti_tti(TITTIInput(total_income_normal=60000000,
                                  regime=C.REGIME_OLD, age=40))
    assert r["TaxNormal"] == 17812500
    assert r["Surcharge"] == 6590625                # 37%
    assert r["Cess"] == 976125                      # 4% of 24403125


def test_amt_max_rule_and_relief():
    inp = TITTIInput(total_income_normal=1200000, amt_payable=500000)
    r = compute_ti_tti(inp)
    assert r["AMTApplied"] is True
    assert r["GrossTaxLiability"] == 520000         # AMT + 4% cess wins
    inp.relief_89 = 10000
    r2 = compute_ti_tti(inp)
    assert r2["NetTaxLiability"] == 510000
