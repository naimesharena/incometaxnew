"""CFL carry-forward expiry + BFLA set-off tests."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from itr_filing.carry_forward import (CFLTracker, LossEntry, bfla, ay_start,
                                      total_income_after_bfla)

CUR = "2026-27"


def test_cfl_expiry_hp_8_years():
    t = CFLTracker([
        LossEntry(ay="2018-19", hp=100000),   # 8 AYs old -> expires in 26-27
        LossEntry(ay="2019-20", hp=50000),    # still within 8 AYs
    ])
    avail = t.available(CUR)
    assert avail["hp"] == 50000


def test_cfl_speculation_4_years():
    t = CFLTracker([
        LossEntry(ay="2022-23", speculation=80000),  # 4 AYs -> expired
        LossEntry(ay="2023-24", speculation=30000),  # alive
    ])
    assert t.available(CUR)["speculation"] == 30000


def test_cfl_specified_business_no_limit():
    t = CFLTracker([LossEntry(ay="2010-11", specified_bus=500000)])
    assert t.available(CUR)["specified_bus"] == 500000


def test_cfl_depreciation_indefinite():
    t = CFLTracker([LossEntry(ay="2005-06", depreciation=200000)])
    assert t.available(CUR)["depreciation"] == 200000


def test_bfla_hp_loss_only_hp_income():
    r = bfla({"SAL": 500000, "HP": 100000}, {"hp": 150000})
    assert r.setoff_by_category["hp"] == 100000
    assert r.remaining["hp"] == 50000
    assert r.income_after_setoff["HP"] == 0
    assert r.income_after_setoff["SAL"] == 500000     # untouched


def test_bfla_bp_loss_cascade_business_to_speculation():
    # business BF loss absorbs BP first, then flows to SPEC [W36 chain]
    r = bfla({"BP": 100000, "SPEC": 60000}, {"bp": 140000})
    assert r.setoff_by_category["bp"] == 140000
    assert r.income_after_setoff["BP"] == 0
    assert r.income_after_setoff["SPEC"] == 20000
    assert r.remaining["bp"] == 0


def test_bfla_speculation_loss_strict():
    r = bfla({"BP": 100000, "SPEC": 40000}, {"speculation": 60000})
    assert r.setoff_by_category["speculation"] == 40000
    assert r.remaining["speculation"] == 20000
    assert r.income_after_setoff["BP"] == 100000      # speculation cannot touch BP


def test_bfla_stcg_loss_flows_into_ltcg():
    # brought-forward STCL can absorb LTCG too [AK links]
    r = bfla({"STCG20": 50000, "LTCG125": 100000}, {"stcg": 120000})
    assert r.setoff_by_category["stcg"] == 120000
    assert r.income_after_setoff["STCG20"] == 0
    assert r.income_after_setoff["LTCG125"] == 30000


def test_bfla_ltcg_loss_only_ltcg():
    r = bfla({"STCG20": 50000, "LTCG125": 100000}, {"ltcg": 120000})
    assert r.setoff_by_category["ltcg"] == 100000
    assert r.remaining["ltcg"] == 20000
    assert r.income_after_setoff["STCG20"] == 50000   # LTCG loss cannot touch STCG


def test_bfla_depreciation_not_salary():
    r = bfla({"SAL": 300000, "BP": 100000, "OS": 50000}, {}, depreciation=200000)
    assert r.depreciation_setoff == 150000            # BP + OS only
    assert r.income_after_setoff["SAL"] == 300000
    assert r.remaining["depreciation"] == 50000


def test_total_income_after_bfla():
    r = bfla({"SAL": 300000, "HP": 100000}, {"hp": 150000})
    assert total_income_after_bfla(r) == 300000
