"""Schedule-level computations ported from hidden ITR-1 sheets.

Sources:
  Schedule EA 10(13A)!G7/G10/G12      HRA exemption u/s 10(13A)
  80D!L4..L52 + Income Details!BI18/BI19/BK18/BH143   80D AY 2026-27 selection model
  Income Details!BM46                 80DDB caps (40,000 / 1,00,000)
  80U-80DD sheet + IncD!BG160-161     80DD / 80U amounts
  80G sheet (Q13/Q27/R35/R40/R48) + 'Income Details'!AN151   80G eligible donations
"""
from dataclasses import dataclass, field
from typing import List
from .tax_engine import excel_round
from . import constants as C


# ---------------------------------------------------------------------------
# Schedule EA - HRA exemption u/s 10(13A)
# ---------------------------------------------------------------------------
def hra_exemption_10_13a(hra_received: float, rent_paid: float, basic_salary: float,
                         dearness_allowance: float, metro: bool, regime: int,
                         employed: bool = True) -> int:
    """[Schedule EA 10(13A)!G12] exemption = MIN(A, B, C), OLD regime only.

    A = actual HRA received
    B = MAX(0, rent paid - ROUND(10% of salary))          [G10]
    C = 50% of salary (metro) / 40% (non-metro)
    salary = MAX(0, basic + DA)                            [G7]
    New regime (BacValue=1) or 'Not employed' -> 0.
    """
    if regime == C.REGIME_NEW or not employed:
        return 0
    salary = max(0.0, basic_salary + dearness_allowance)
    a = float(hra_received)
    b = max(0.0, rent_paid - excel_round(salary * 0.10))
    c = salary * (0.50 if metro else 0.40)
    return excel_round(min(a, b, c))


# ---------------------------------------------------------------------------
# Schedule 80D - AY 2026-27 single-selection model [DataBase!Selection80D list]
# ---------------------------------------------------------------------------
#: option -> cap (option 7 is age dependent - see Schedule80D.premium_cap)
CAPS_80D = {1: 25000, 2: 50000, 3: 25000, 4: 50000, 5: 50000, 6: 75000, 7: 100000}
CAPS_80D_MEDICAL = {1: 50000, 2: 50000, 3: 100000}   # [BI21]
CAP_80D_CHECKUP = 5000                              # [BH143]
CAP_80D_TOTAL = 100000                              # [80D!L52]


@dataclass
class Schedule80D:
    """Hidden '80D' sheet model (AY 2026-27)."""
    selection: int = 0            # 1..7 per Selection80D list
    premium_paid: float = 0
    medical_selection: int = 0    # Selection80DB: 1/2/3
    medical_expenditure: float = 0
    checkup_selection: int = 0    # Selection80DC: 1/2/3
    checkup_amount: float = 0
    taxpayer_age: int = 0

    def premium_cap(self) -> int:
        """[BK18] cap table; option 7 -> 75k when age<=59 else 1L [BK9]."""
        if self.selection == 7:
            return 75000 if self.taxpayer_age <= 59 else 100000
        return CAPS_80D.get(self.selection, 0)

    def eligible_amount(self, total_income: float) -> int:
        """[80D!L52] MIN(100000, SUM(components), TOTAL_INCOME)."""
        prem = min(self.premium_cap(), float(self.premium_paid or 0))
        med = min(CAPS_80D_MEDICAL.get(self.medical_selection, 0),
                  float(self.medical_expenditure or 0))
        chk = min(CAP_80D_CHECKUP, float(self.checkup_amount or 0)) \
            if self.checkup_selection else 0
        return excel_round(min(CAP_80D_TOTAL, prem + med + chk, total_income))


# ---------------------------------------------------------------------------
# 80DD / 80U / 80DDB
# ---------------------------------------------------------------------------
def amount_80dd(disability_type: int) -> int:
    """1-dependent with disability -> 75k; 2-severe disability -> 125k
    [80U-80DD sheet; IncD!BG160-161 pattern]."""
    return {1: C.CAP_80DD_NORMAL, 2: C.CAP_80DD_SEVERE}.get(disability_type, 0)


def amount_80u(severity: int) -> int:
    """1-self with disability -> 75k; 2-severe -> 125k."""
    return {1: C.CAP_80U_NORMAL, 2: C.CAP_80U_SEVERE}.get(severity, 0)


def amount_80ddb(selection: int) -> int:
    """[BM46] 1-self or dependent -> 40k; 2-senior citizen -> 1L."""
    return {1: 40000, 2: 100000}.get(selection, 0)


# ---------------------------------------------------------------------------
# Schedule 80G
# ---------------------------------------------------------------------------
@dataclass
class Donation80G:
    amount: float = 0
    percent: int = 100            # 100 or 50
    with_qualifying_limit: bool = False
    cash: bool = True             # cash > 2000 not eligible [80G!R48 IF(L48>2000,0,..)]


def qualifying_limit_80g(total_income: float, via_before_80g: float,
                         via_after_80g: float, s80gg: float,
                         s80eea_calc: float, s80eeb_calc: float) -> int:
    """[80G!C33] 10% of (TI - (VI-A other than 80G/80GGA/GGC + 80GG + 80EEA + 80EEB))."""
    adjusted = total_income - (via_before_80g + via_after_80g
                               + min(60000, s80gg) + s80eea_calc + s80eeb_calc)
    return excel_round(0.10 * max(0.0, adjusted))


def eligible_donations_80g(donations: List[Donation80G], qualifying_limit: int) -> int:
    """Port of the 80G sheet category sums:
    100%/50% without qualifying limit [Q13/Q27], qualifying-limit categories
    aggregated then capped at the 10% limit [R35/R40/R48];
    cash donations > Rs 2,000 are not eligible [R48 IF(L48>2000,0,..)]."""
    nl100 = nl50 = ql = 0.0
    for d in donations:
        amt = 0.0 if (d.cash and d.amount > 2000) else float(d.amount)
        eligible = amt if d.percent == 100 else amt / 2
        if d.with_qualifying_limit:
            ql += eligible
        elif d.percent == 100:
            nl100 += eligible
        else:
            nl50 += eligible
    return excel_round(nl100 + nl50 + min(ql, qualifying_limit))


# ---------------------------------------------------------------------------
# Entertainment allowance u/s 16(ii) cap
# ---------------------------------------------------------------------------
def cap_entertainment_allowance_16ii(amount: float, salary_171: float) -> int:
    """[VBA ~1866] 16(ii) allowed = least of actual, Rs 5,000, 1/5 of salary
    (only for government employees in law; utility enforces the numeric caps)."""
    return excel_round(min(float(amount or 0), 5000, float(salary_171 or 0) / 5))
