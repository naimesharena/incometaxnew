"""Interest & fee engines u/s 234A / 234B / 234C / 234F / 234-I.

Sources (ITR1_AY_26-27_V1.2.xlsm):
  VBA mdCalInterst234B.ComputeInterest / Calculate_InterestPayable234B
  hidden sheet 'Taxes Paid and Verification' monthly grid (GH columns)
  VBA ~line 18470 fee 234F assignment
"""
from datetime import date
from .tax_engine import excel_round

AY_START = date(2026, 4, 1)     # AY 2026-27 starts April 2026
ASSESSMENT_YEAR_INT = 2026       # [VBA] AssYear = 2026


def floor100(x: float) -> float:
    """[VBA] WorksheetFunction.Floor(shortFall/100, 1)*100 when shortFall>100."""
    return (int(x) // 100) * 100


def interest_234b(net_tax_liability: float, advance_tax_paid: float,
                  tds: float, tcs: float, verification_date: date,
                  age: int, filing_sec: str = "") -> int:
    """Port of Calculate_InterestPayable234B.

    Conditions [VBA]:
      - base = NTL - (TDS + TCS) must be >= 10,000  (CONST_NET_Limit)
      - advance tax paid must be < 90% of base       (CONST_ATP_Limit)
      - shortFall = MAX(0, NTL - (adv + tds + tcs)), floored to 100
      - months from April of AY up to verification month (inclusive):
            mthdop - 4 + (yrdop - AssYear)*12 + 1
      - 1% simple interest per month
      - senior citizens (age > 59) exempt in ITR-1 context
      - 139(8A) late filing: NTL increased by any refund/amount due from
        previous processing (AdditionalTax) - pass it inside NTL if applicable.
    """
    if age > 59:                                    # [VBA] If (bacage > 59) Then 234B = 0
        return 0
    base = net_tax_liability - (tds + tcs)
    if base < 10000:
        return 0
    if advance_tax_paid >= 0.90 * base:
        return 0
    shortfall = max(0.0, net_tax_liability - (advance_tax_paid + tds + tcs))
    if shortfall > 100:
        shortfall = floor100(shortfall)
    if shortfall <= 0:
        return 0
    vd = verification_date
    months = vd.month - 4 + (vd.year - ASSESSMENT_YEAR_INT) * 12 + 1
    if months <= 0:
        return 0
    return excel_round(shortfall * 0.01 * months)


def interest_234a(net_tax_liability: float, due_date: date, filing_date: date) -> int:
    """1% per month (or part of month) of net tax liability from the due date
    to the date of filing. The utility accumulates month-wise values in the
    hidden grid (Sheet5!GH columns); equivalent closed-form below."""
    if filing_date <= due_date or net_tax_liability <= 0:
        return 0
    # months inclusive of both end months, part month = full month
    months = (filing_date.year - due_date.year) * 12 + (filing_date.month - due_date.month)
    if filing_date.day > due_date.day:
        months += 1
    months = max(months, 1)
    return excel_round(net_tax_liability * 0.01 * months)


def fee_234f(total_income: float, filed_after_due_date: bool) -> int:
    """[VBA ~18470-18497] late-filing fee: Rs 1,000 if TI <= 5L else Rs 5,000."""
    if not filed_after_due_date:
        return 0
    return 1000 if total_income <= 500000 else 5000


def fee_234i(revised_return: bool, base_234f: int) -> int:
    """D10a fee for furnishing revised return u/s 234-I (label at Income
    Details r189). The utility applies it when a revised return is filed;
    equals the 234F fee component."""
    return base_234f if revised_return else 0


def interest_234c(q_taxes: list, rebate_87a: float, instalments: list,
                  tds: float, tcs: float, sec89: float, sec89a: float,
                  base_net_tax: float, age: int) -> int:
    """Exact port of mIncmDtls.calcIntrst234C.

    q_taxes     : [Q1Tax..Q5Tax] tax on cumulative quarterly income estimates
                  (utility computes each via the slab engine)
    instalments : advance tax paid by [15-Jun, 15-Sep, 15-Dec, 15-Mar, later]
                  (slab0..slab4)
    Rules:
      base_k  = QkTax - Rebate87A, then * 1.04 (cess added)
      TDS'    = TDS + Section89 + Section89A
      i  (15% by 15-Jun): tolerance floor(12%*(base-TDS'-TCS),100);
             shortfall if paid < 15%*base: 15%*base - paid (0 if paid>=tolerance),
             round down to 100, interest = shortfall * 1% * 3 months
      ii (45% by 15-Sep): tolerance floor(36%*...); shortfall vs 45% cumulative
      iii(75% by 15-Dec): shortfall vs 75% cumulative (no tolerance)
      iv (100% by 15-Mar): shortfall vs 100%, 1 month interest
      v  : excess paid carried forward; balance of Q5Tax, 1% for 1 month
      No interest if (base_net_tax - TDS - TCS) < 10,000 or age > 59.
    """
    if age > 59:
        return 0
    if (base_net_tax - tds - tcs) < 10000:
        return 0
    slabs = list(instalments) + [0.0] * (5 - len(instalments))
    qts = list(q_taxes) + [0.0] * (5 - len(q_taxes))
    tds_eff = tds + sec89 + sec89a

    def base(k):
        return (qts[k] - rebate_87a) * 1.04

    total = 0.0
    # i - 15%
    b = base(0)
    if b - tds_eff - tcs >= 0:
        tol = (int((0.12 * (b - tds_eff - tcs)) // 100)) * 100
        if slabs[0] < (b - tds_eff - tcs) * 0.15:
            short = 0 if slabs[0] >= tol else (b - tds_eff - tcs) * 0.15 - slabs[0]
            if short > 100:
                short = (int(short) // 100) * 100
            total += short * 0.01 * 3
        # ii - 45%
        b = base(1)
        tol = (int((0.36 * (b - tds_eff - tcs)) // 100)) * 100
        if slabs[0] + slabs[1] < (b - tds_eff - tcs) * 0.45:
            short = 0 if (slabs[0] + slabs[1]) >= tol else \
                (b - tds_eff - tcs) * 0.45 - slabs[0] - slabs[1]
            if short > 100:
                short = (int(short) // 100) * 100
            total += short * 0.01 * 3
        # iii - 75%
        b = base(2)
        if slabs[0] + slabs[1] + slabs[2] < (b - tds_eff - tcs) * 0.75:
            short = (b - tds_eff - tcs) * 0.75 - slabs[0] - slabs[1] - slabs[2]
            if short > 100:
                short = (int(short) // 100) * 100
            total += short * 0.01 * 3
        # iv - 100%
        b = base(3)
        if sum(slabs[:4]) < (b - tds_eff - tcs):
            short = b - tds_eff - tcs - sum(slabs[:4])
            if short > 100:
                short = (int(short) // 100) * 100
            total += short * 0.01
        # v - final balancing month
        excess = sum(slabs[:4]) - b - tds_eff - tcs
        if excess < 0:
            excess = 0
        slabs[4] += excess
        b = base(4)
        bal = b - slabs[4]
        if bal > 100:
            bal = (int(bal) // 100) * 100
        total += max(0.0, bal * 0.01)
    return excel_round(total)
