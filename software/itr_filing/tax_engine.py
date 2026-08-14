"""Tax computation engine - faithful port of the ITR-1 VBA slab functions.

Sources (ITR1_AY_26-27_V1.2.xlsm, vba/ITR1.vba):
  - calcTaxPayableOnTI      : OLD regime slabs, age-based (senior/super-senior)
  - calcTaxPayableOnTINTR   : NEW regime slabs AY 2026-27 (Finance Act 2025)
  - calcTaxPayableOnTINTRQ1 : same new-regime slabs used for the quarterly
                              advance-tax estimator (Q1..Q5 series)
  - Income Details!AO177    : Rebate u/s 87A (both regimes + marginal relief)
  - Income Details!AO180/181: Health & Education Cess 4%
All VBA rounding uses Excel-style Round(x, 0) (round half away from zero).
"""
from .constants import (
    REGIME_NEW, REGIME_OLD,
    REBATE_87A_NEW_LIMIT, REBATE_87A_NEW_TI,
    REBATE_87A_OLD_LIMIT, REBATE_87A_OLD_TI,
    CESS_RATE,
)


def excel_round(x: float, ndigits: int = 0) -> int:
    """Excel ROUND: half away from zero."""
    import decimal
    d = decimal.Decimal(str(x))
    q = decimal.Decimal(1).scaleb(-ndigits)
    return int(d.quantize(q, rounding=decimal.ROUND_HALF_UP))


def calculate_age(dob, on=None) -> int:
    """[VBA:EfilingCommon.calculateAge] age in completed years."""
    from datetime import date, datetime
    on = on or date.today()
    if isinstance(dob, str):
        s = dob.strip()
        if "/" in s:
            d, m, y = s.split("/")          # DD/MM/YYYY  (utility format)
            dob = date(int(y), int(m), int(d))
        else:
            dob = date.fromisoformat(s)     # YYYY-MM-DD
    elif isinstance(dob, datetime):
        dob = dob.date()
    if not isinstance(dob, date):
        raise ValueError("DOB required for age-based slabs")
    years = on.year - dob.year - ((on.month, on.day) < (dob.month, dob.day))
    return years


def tax_old_regime(total_income: float, age: int) -> int:
    """[VBA:calcTaxPayableOnTI] OLD regime tax on Total Income (excl. 112A LTCG).

    Basic exemption: <60y = 2.5L, 60-79y = 3L, >=80y = 5L.
    """
    ti = total_income
    if age > 79:                                   # super senior citizen
        if ti <= 500000:
            return 0
        if ti <= 1000000:
            return excel_round((ti - 500000) * 0.2)
        return excel_round((ti - 1000000) * 0.3 + 100000)
    if age > 59:                                   # senior citizen
        if ti <= 300000:
            return 0
        if ti <= 500000:
            return excel_round((ti - 300000) * 0.05)
        if ti <= 1000000:
            return excel_round((ti - 500000) * 0.2 + 10000)
        return excel_round((ti - 1000000) * 0.3 + 110000)
    # below 60
    if ti <= 250000:
        return 0
    if ti <= 500000:
        return excel_round((ti - 250000) * 0.05)
    if ti <= 1000000:
        return excel_round((ti - 500000) * 0.2 + 12500)
    return excel_round((ti - 1000000) * 0.3 + 112500)


def tax_new_regime(total_income: float) -> int:
    """[VBA:calcTaxPayableOnTINTR] NEW regime slabs AY 2026-27 (FY 2025-26).

    0-4L nil; 4-8L 5%; 8-12L 10%; 12-16L 15%; 16-20L 20%; 20-24L 25%; >24L 30%.
    No age-based variation under 115BAC.
    """
    ti = total_income
    if ti <= 400000:
        return 0
    if ti <= 800000:
        return excel_round((ti - 400000) * 0.05)
    if ti <= 1200000:
        return excel_round((ti - 800000) * 0.10 + 20000)
    if ti <= 1600000:
        return excel_round((ti - 1200000) * 0.15 + 60000)
    if ti <= 2000000:
        return excel_round((ti - 1600000) * 0.20 + 120000)
    if ti <= 2400000:
        return excel_round((ti - 2000000) * 0.25 + 200000)
    return excel_round((ti - 2400000) * 0.30 + 300000)


def tax_on_total_income(total_income: float, regime: int, age: int) -> int:
    """D1 'Tax Payable on Total Income' [IncD!AO176]."""
    if regime == REGIME_NEW:
        return tax_new_regime(total_income)
    return tax_old_regime(total_income, age)


def rebate_87a(total_tax_payable: int, total_income: float, total_income_incl_112a: float,
               regime: int) -> int:
    """D2 Rebate u/s 87A - exact port of Income Details!AO177.

    NEW regime (BacValue=1): tested on Total Income EXCLUDING 112A LTCG
      TI <= 12,00,000 -> MIN(tax, 60,000)
      else marginal relief: if tax > (TI - 12,00,000) -> tax - (TI - 12,00,000)
    OLD regime (BacValue=2): tested on Total Income INCLUDING 112A LTCG
      TI_new <= 5,00,000 -> MIN(tax, 12,500)
    """
    if regime == REGIME_NEW:
        if total_income <= REBATE_87A_NEW_TI:
            return min(total_tax_payable, REBATE_87A_NEW_LIMIT)
        if total_tax_payable > (total_income - REBATE_87A_NEW_TI):
            return excel_round(total_tax_payable - (total_income - REBATE_87A_NEW_TI))
        return 0
    # old regime
    if total_income_incl_112a <= REBATE_87A_OLD_TI:
        return min(total_tax_payable, REBATE_87A_OLD_LIMIT)
    return 0


def education_cess(tax_after_rebate: int) -> int:
    """D4 Health & Education Cess @4% [IncD!AO180]."""
    return excel_round(tax_after_rebate * CESS_RATE)


def ltcg_112a_not_chargeable(sale_consideration: float, cost_of_acquisition: float) -> int:
    """C3a(iii) LTCG u/s 112A not chargeable [IncD!AO175].

    =IF(MAX(0,Sale-Cost)>125000, 0, MAX(0,Sale-Cost))
    (ITR-1 is only eligible when the gain itself is <= 1,25,000.)
    """
    gain = max(0.0, sale_consideration - cost_of_acquisition)
    return excel_round(0 if gain > 125000 else gain)
