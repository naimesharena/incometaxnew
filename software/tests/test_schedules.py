"""Tests for schedule ports: HRA (EA 10(13A)), 80D AY 26-27 model, 80DD/80U,
80DDB, 80G, entertainment allowance cap, interest 234C."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from itr_filing.schedules import (hra_exemption_10_13a, Schedule80D, amount_80dd,
                                  amount_80u, amount_80ddb, Donation80G,
                                  qualifying_limit_80g, eligible_donations_80g,
                                  cap_entertainment_allowance_16ii)
from itr_filing.interest import interest_234c
from itr_filing import constants as C


# ------------------------------------------------------------------ HRA
def test_hra_old_regime_metro():
    # [Schedule EA 10(13A)!G12] MIN(HRA, rent-10%salary, 50%salary)
    e = hra_exemption_10_13a(hra_received=120000, rent_paid=140000,
                             basic_salary=300000, dearness_allowance=0,
                             metro=True, regime=C.REGIME_OLD)
    assert e == min(120000, 140000 - 30000, 150000)   # 110000


def test_hra_new_regime_zero():
    assert hra_exemption_10_13a(120000, 140000, 300000, 0, True,
                                C.REGIME_NEW) == 0


def test_hra_non_metro_40pct():
    e = hra_exemption_10_13a(200000, 200000, 300000, 0, False, C.REGIME_OLD)
    assert e == 120000    # 40% of salary binds


# ------------------------------------------------------------------ 80D
def test_80d_selection_caps_AY26_27():
    # [BK18] selection caps
    assert Schedule80D(selection=1, premium_paid=99999).premium_cap() == 25000
    assert Schedule80D(selection=2, premium_paid=99999).premium_cap() == 50000
    assert Schedule80D(selection=6, premium_paid=99999).premium_cap() == 75000
    assert Schedule80D(selection=7, premium_paid=200000,
                       taxpayer_age=65).premium_cap() == 100000
    assert Schedule80D(selection=7, premium_paid=200000,
                       taxpayer_age=55).premium_cap() == 75000


def test_80d_eligible_amount():
    # [80D!L52] MIN(100000, components, TI)
    d = Schedule80D(selection=2, premium_paid=60000,
                    medical_selection=3, medical_expenditure=120000,
                    checkup_selection=1, checkup_amount=8000)
    # premium 50000 + medical 100000 + checkup 5000 -> capped at 100000
    assert d.eligible_amount(1000000) == 100000
    assert d.eligible_amount(80000) == 80000        # capped by TI


def test_80dd_80u_80ddb():
    assert amount_80dd(1) == 75000
    assert amount_80dd(2) == 125000
    assert amount_80u(1) == 75000
    assert amount_80u(2) == 125000
    assert amount_80ddb(1) == 40000     # [BM46]
    assert amount_80ddb(2) == 100000


# ------------------------------------------------------------------ 80G
def test_80g_qualifying_limit():
    # [80G!C33] 10% of adjusted TI
    ql = qualifying_limit_80g(total_income=1000000, via_before_80g=150000,
                              via_after_80g=10000, s80gg=0,
                              s80eea_calc=0, s80eeb_calc=0)
    assert ql == 84000


def test_80g_eligible_donations():
    dons = [
        Donation80G(amount=50000, percent=100, with_qualifying_limit=False, cash=False),
        Donation80G(amount=20000, percent=50, with_qualifying_limit=False, cash=False),
        Donation80G(amount=100000, percent=100, with_qualifying_limit=True, cash=False),
        Donation80G(amount=3000, percent=100, with_qualifying_limit=False, cash=True),
    ]
    # cash 3000 > 2000 excluded; ql capped at 84000
    assert eligible_donations_80g(dons, 84000) == 50000 + 10000 + 84000


# ------------------------------------------------------------------ 16(ii)
def test_entertainment_allowance_cap():
    # [VBA ~1866] least of actual, 5000, salary/5
    assert cap_entertainment_allowance_16ii(10000, 100000) == 5000
    assert cap_entertainment_allowance_16ii(3000, 100000) == 3000
    assert cap_entertainment_allowance_16ii(10000, 20000) == 4000


# ------------------------------------------------------------------ 234C
def test_234c_no_instalments():
    # [VBA calcIntrst234C] base 1,00,000 -> *1.04 = 104000 per quarter
    got = interest_234c([100000] * 5, 0, [0, 0, 0, 0, 0], 0, 0, 0, 0, 100000, 40)
    # i: 15600*.03=468; ii: 46800*.03=1404; iii: 78000*.03=2340;
    # iv: 104000*.01=1040; v: 104000*.01=1040  => 6292
    assert got == 6292


def test_234c_threshold_and_senior():
    assert interest_234c([100000] * 5, 0, [0, 0, 0, 0, 0], 0, 0, 0, 0, 9000, 40) == 0
    assert interest_234c([100000] * 5, 0, [0, 0, 0, 0, 0], 0, 0, 0, 0, 100000, 65) == 0


def test_234c_full_advance_tax_paid():
    # Q1..Q4 are cumulative quarterly taxes; Q5 = annual minus Q4 (usually 0)
    q_taxes = [25000, 50000, 75000, 100000, 0]
    paid = [15600, 31200, 31200, 26000, 0]      # 15/45/75/100% of 104000
    got = interest_234c(q_taxes, 0, paid, 0, 0, 0, 0, 100000, 40)
    assert got == 0


# ------------------------------------------------------------------ pipeline wiring
def test_pipeline_with_80d_selection_and_hra():
    from datetime import date
    from itr_filing.itr1 import ITR1Return, PersonalInfo, FilingStatus, Salary, ChapterVIA
    r = ITR1Return()
    r.personal = PersonalInfo(first_name="A", last_name="B", pan="ABCPS1234K",
                              date_of_birth=date(1980, 1, 1), nature_of_employment="PE")
    r.filing = FilingStatus(opt_out_new_regime="Yes", filing_sec_code=11)  # OLD regime
    r.salary = Salary(salary_17_1=900000, basic_salary=600000,
                      hra_received=150000, rent_paid=180000, metro=True)
    r.chapter_via = ChapterVIA(
        schedule_80d=Schedule80D(selection=1, premium_paid=40000),
        disability_80dd=1, severity_80u=0, selection_80ddb=1,
        donations_80g=[Donation80G(amount=20000, percent=100,
                                   with_qualifying_limit=False, cash=False)],
        user_total=1000000)
    s = r.compute_summary()
    # HRA auto: MIN(150000, 180000-60000, 300000) = 120000
    assert r.salary.hra_exemption(C.REGIME_OLD) == 120000
    comps = s["ChapterVIAComponents"]
    assert comps["80D"] == 25000            # selection 1 cap
    assert comps["80DD"] == 75000           # old regime allowed
    assert comps["80DDB"] == 0              # no amount entered
    assert comps["80G"] == 20000            # 100% no qualifying limit
