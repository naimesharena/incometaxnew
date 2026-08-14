"""CYLA set-off engine tests (ITR-2/ITR-3 CYLA-BFLA sheet port)."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from itr_filing.setoff import cyla, total_income_after_cyla, HP_LOSS_CAP
from itr_filing import constants as C


def test_hp_loss_cap_2_lakh():
    r = cyla({"SAL": 500000}, hp_loss=300000, bp_loss=0, os_loss=0)
    assert r.hp_loss_setoff == HP_LOSS_CAP          # capped at 2L
    assert r.income_after_setoff["SAL"] == 300000
    assert r.hp_loss_remaining == 0


def test_hp_loss_setoff_order():
    # SAL first, then OS (capital gains between are empty)
    r = cyla({"SAL": 100000, "OS": 150000}, hp_loss=200000, bp_loss=0, os_loss=0)
    assert r.hp_loss_setoff == 200000
    assert r.income_after_setoff["SAL"] == 0
    assert r.income_after_setoff["OS"] == 50000


def test_bp_loss_not_against_salary():
    # [H8 absent in matrix] business loss cannot absorb salary
    r = cyla({"SAL": 500000}, hp_loss=0, bp_loss=100000, os_loss=0)
    assert r.bp_loss_setoff == 0
    assert r.bp_loss_remaining == 100000
    assert r.income_after_setoff["SAL"] == 500000


def test_bp_loss_against_stcg_and_hp():
    r = cyla({"HP": 60000, "STCG20": 400000}, hp_loss=0, bp_loss=100000, os_loss=0)
    assert r.bp_loss_setoff == 100000
    assert r.income_after_setoff["HP"] == 0        # HP first in BP order
    assert r.income_after_setoff["STCG20"] == 360000


def test_os_loss_against_salary_and_heads():
    r = cyla({"SAL": 40000, "HP": 80000}, hp_loss=0, bp_loss=0, os_loss=100000)
    assert r.os_loss_setoff == 100000
    assert r.income_after_setoff["SAL"] == 0
    assert r.income_after_setoff["HP"] == 20000
    assert r.os_loss_remaining == 0


def test_hp_loss_remaining_new_regime_zeroed():
    # [G26] IF(bacValue=1, 0, ...) - new regime disallows the carry
    r_old = cyla({}, hp_loss=150000, bp_loss=0, os_loss=0, regime=C.REGIME_OLD)
    assert r_old.hp_loss_remaining == 150000
    r_new = cyla({}, hp_loss=150000, bp_loss=0, os_loss=0, regime=C.REGIME_NEW)
    assert r_new.hp_loss_remaining == 0


def test_special_rate_income_not_touched_by_bp_order_gap():
    # BP loss order starts at STCG20 (H13/STCG15 absent)
    r = cyla({"STCG15": 100000, "LTCG125": 200000}, hp_loss=0,
             bp_loss=250000, os_loss=0)
    assert r.income_after_setoff["STCG15"] == 100000     # untouched
    assert r.income_after_setoff["LTCG125"] == 0         # absorbed first (order)
    assert r.bp_loss_setoff == 250000 - 50000


def test_total_income_after_cyla():
    r = cyla({"SAL": 500000, "OS": 100000}, hp_loss=200000, bp_loss=0, os_loss=0)
    assert total_income_after_cyla(r) == 400000
