"""Chapter VI-A deduction caps - exact port of Income Details!AN115..AN163 formulas.

Each function cites the Excel formula it reproduces. 'ti' = TOTAL_INCOME
[hidden helper TOTAL_INCOME = Income Details!BI17], 'gti' = Gross Total Income.
"""
from .constants import *
from .tax_engine import excel_round


def cap_80c(input_amt: float, ti: float) -> int:
    """[AN115] =MIN(MIN(VALUE(IncD.Section80C),150000),TOTAL_INCOME)"""
    return excel_round(min(min(float(input_amt or 0), CAP_80C_80CCC), ti))


def cap_80ccc(input_amt: float, gti: float, s80c_allowed: float, ti: float) -> int:
    """[AN116] 80CCC shares the 1.5L aggregate cap with 80C."""
    amt = float(input_amt or 0)
    allowed = min(amt, min(CAP_80C_80CCC, abs(gti)),
                  min(CAP_80C_80CCC, abs(gti)) - s80c_allowed, ti)
    return excel_round(min(allowed, CAP_80C_80CCC))


def cap_80ccd1(input_amt: float, net_salary_less_perq: float, employer_pct: float,
               regime: int, ti: float) -> int:
    """[AN121/AN137] 80CCD(1): % of (Net salary - perquisites).

    OLD regime: private 10% / government 14% (BA126=1/2).
    NEW regime: 14% for both employer categories.
    """
    base = max(0.0, net_salary_less_perq)
    pct = 0.14 if regime == REGIME_NEW else employer_pct
    return excel_round(min(float(input_amt or 0), excel_round(pct * base), ti))


def cap_80ccd1b(input_amt: float, ti: float) -> int:
    """[AN126] =MIN(VALUE(IncD.Section80CCD1B_SE),50000,TOTAL_INCOME)"""
    return excel_round(min(float(input_amt or 0), CAP_80CCD_1B, ti))


def cap_80ccd2(input_amt: float, net_salary_less_perq: float, employer_pct: float,
               regime: int, ti: float) -> int:
    """[AN137] 80CCD(2) employer contribution: old 10%/14%, new 14%."""
    return cap_80ccd1(input_amt, net_salary_less_perq, employer_pct, regime, ti)


def cap_80ccg(input_amt: float, ti: float) -> int:
    """[AN138] =MIN(IF(TOTAL_INCOME>1200000,0,MIN(VALUE(80CCG),25000)),TOTAL_INCOME)"""
    if ti > CAP_80CCG_TI_LIMIT:
        return 0
    return excel_round(min(min(float(input_amt or 0), CAP_80CCG), ti))


def cap_80d(schedule_amount: float, gti: float) -> int:
    """[AN140] =IF(IncD.GrossTotIncome<0,0,IncD.Section80DValue)
    [AN141] overall cap MIN(100000, ...) - schedule_amount is the value
    computed by the hidden '80D' sheet (premium caps by age + 5k check-up)."""
    if gti < 0:
        return 0
    return excel_round(min(float(schedule_amount or 0), CAP_80D_OVERALL))


def cap_80dd(schedule_amount: float, regime: int) -> int:
    """[AN144] =IF(BacValue=1,0,IncD.Section80DD) - not available under new regime."""
    if regime == REGIME_NEW:
        return 0
    return excel_round(float(schedule_amount or 0))


def cap_80ddb(schedule_amount: float) -> int:
    """[AN145] =IF(ISERROR(IncdSection80DD),0,IncdSection80DD)"""
    return excel_round(float(schedule_amount or 0))


def cap_80e(input_amt: float, ti: float) -> int:
    """[AN147] =MIN(VALUE(IncD.Section80E),TOTAL_INCOME) - no monetary cap."""
    return excel_round(min(float(input_amt or 0), ti))


def cap_80ee(input_amt: float, ti: float) -> int:
    """[AN148] cap 50,000."""
    return excel_round(min(min(float(input_amt or 0), ti), CAP_80EE))


def cap_80eea(input_amt: float, ti: float) -> int:
    """[AN149] cap 1,50,000."""
    return excel_round(min(min(float(input_amt or 0), ti), CAP_80EEA))


def cap_80eeb(input_amt: float, ti: float) -> int:
    """[AN150] cap 1,50,000."""
    return excel_round(min(min(float(input_amt or 0), ti), CAP_80EEB))


def cap_80g(schedule_amount: float) -> int:
    """[AN151] =IncD.Section80G - value computed by hidden '80G' sheet
    (donee categories 100%/50%, qualifying limit 10% of adjusted GTI)."""
    return excel_round(float(schedule_amount or 0))


def cap_80gg(input_amt: float, hra_exemption: float, ti: float,
             other_via_deductions: float) -> int:
    """[AN152] =MIN(MIN(ROUND(MIN(IF(HRA>0,55000,60000),
    0.25*(MAX(0,TOTAL_INCOME-(SUM(AN115:AW151)+SUM(AN154:AW161)+0)))),0),
    VALUE(IncD.Section80GG)),TOTAL_INCOME)"""
    cap = CAP_80GG_WITH_HRA if (hra_exemption or 0) > 0 else CAP_80GG_WITHOUT_HRA
    quarter = 0.25 * max(0.0, ti - other_via_deductions)
    # NB: the AY 2026-27 utility literally takes MIN(ROUND(...), 0), which
    # zeroes 80GG (deduction discontinued for this AY). Faithful port:
    inner = min(excel_round(min(cap, quarter)), 0)
    return excel_round(min(min(inner, float(input_amt or 0)), ti))


def cap_80gga(eligible_donation: float, regime: int) -> int:
    """[AN154] =IF(BacValue=1,0,Total_Donation_Eligible_80GGA)"""
    if regime == REGIME_NEW:
        return 0
    return excel_round(float(eligible_donation or 0))


def cap_80ggc(eligible_donation: float, regime: int) -> int:
    """[AN155] =IF(BacValue=1,0,Total_Donation_Eligible_80GGC)"""
    if regime == REGIME_NEW:
        return 0
    return excel_round(float(eligible_donation or 0))


def cap_80tta(input_amt: float, ti: float, savings_interest: float) -> int:
    """[AN158] =MIN(MIN(VALUE(...),10000),MAX(0,MIN(TOTAL_INCOME,MAX(0,tta))))"""
    return excel_round(min(min(float(input_amt or 0), CAP_80TTA),
                           max(0.0, min(ti, max(0.0, savings_interest)))))


def cap_80ttb(input_amt: float, ti: float, senior_deposit_interest: float) -> int:
    """[AN159] cap 50,000 for resident senior citizens."""
    return excel_round(min(min(float(input_amt or 0), CAP_80TTB),
                           max(0.0, min(ti, max(0.0, senior_deposit_interest)))))


def cap_80u(schedule_amount: float, regime: int) -> int:
    """[AN160] =IF(BacValue=1,0,IncD.Section80U); schedule value is
    75,000 (40%+ disability) or 1,25,000 (80%+/severe) [hidden 80U-80DD sheet]."""
    if regime == REGIME_NEW:
        return 0
    return excel_round(float(schedule_amount or 0))


def cap_80cch(input_amt: float, allowances_exempt_u10: float) -> int:
    """[AN161] 80CCH Agnipath:
    =IF(AnyOther>0,MIN(AnyOther,MIN(ROUND(Allowances*46.2/100),288000)),0)"""
    amt = float(input_amt or 0)
    if amt > 0:
        return excel_round(min(amt, min(excel_round(float(allowances_exempt_u10 or 0) * AGNIPATH_80CCH_PCT / 100),
                                        AGNIPATH_80CCH_CAP)))
    return 0


def total_chapter_via(components: dict, user_total: float, gti: float) -> int:
    """[AO163] Total Deductions =
    ROUND(MAX(MIN(MIN(input_total, SUM(component calc values)), GTI),0),0)"""
    calc_sum = sum(components.values())
    return excel_round(max(min(min(float(user_total or 0), calc_sum), gti), 0))


def round_total_income(value: float) -> int:
    """[AO164/165] Total Income rounded to nearest 10: ROUND(MAX(0, x), -1)."""
    return excel_round(max(0.0, value), -1)
