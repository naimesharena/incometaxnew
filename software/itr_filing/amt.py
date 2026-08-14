"""Alternate Minimum Tax (115JC/115JD) + AMT credit (AMTC).

Sources (ITR2_AY_26-27_V1.3.xlsm):
  AMT sheet rows 3-14   adjusted total income & 115JC tax
  AMTC sheet rows 4-22  115JD credit grid (AY 2013-14 .. current)

AY 2026-27 rules decoded:
  - AMT not applicable under the new regime [AMT!J10 IF(bacValue=1,0,...)]
  - Adjusted TI = Total Income + VI-A deductions (excl. 80QQB/80RRB)
                  + 10AA + 35AD (net of depreciation)         [AMT!H9]
  - 115JC tax only if adjusted TI > Rs 20,00,000 AND a VI-A
    deduction was claimed                                       [AMT!J13]
  - rate 18.5% (+ surcharge + cess)                             [AMT!N8 0.185]
  - AMT credit: arises when AMT > regular tax [AMTC!G22];
    utilised when regular tax > AMT [AMTC!I6]; carried forward
    15 AYs; credit zeroed under the new regime [AMTC!K9 IF(bacValue=1,0,...)]
"""
from dataclasses import dataclass, field
from typing import Dict, List
from . import constants as C
from .tax_engine import excel_round

AMT_RATE = 0.185
AMT_THRESHOLD = 2000000          # [AMT!J13] adjusted TI > 20L
AMTC_CARRY_YEARS = 15            # s.115JD


def amt_adjusted_total_income(total_income: float, via_deductions: float,
                              deduction_10aa: float, deduction_35ad_net: float,
                              regime: int) -> int:
    """[AMT!J10] IF(bacValue=1, 0, MAX(0, TI + adjustments))."""
    if regime == C.REGIME_NEW:
        return 0
    return excel_round(max(0.0, total_income + via_deductions
                           + deduction_10aa + deduction_35ad_net))


def amt_payable(adjusted_ti: float, via_claimed: bool, regime: int) -> int:
    """[AMT!J13] 18.5% of adjusted TI (surcharge/cess added downstream),
    only when adjusted TI > 20L and a Chapter VI-A deduction was claimed."""
    if regime == C.REGIME_NEW:
        return 0
    if adjusted_ti <= AMT_THRESHOLD or not via_claimed:
        return 0
    tax = excel_round(adjusted_ti * AMT_RATE)
    # surcharge on AMT (individual/HUF rates) [AMT!BTax bands]
    if adjusted_ti > 10000000:
        tax += excel_round(tax * 0.15)
    elif adjusted_ti > 5000000:
        tax += excel_round(tax * 0.10)
    return tax


def amt_cess(amt_tax: float) -> int:
    return excel_round(amt_tax * C.CESS_RATE)


# ---------------------------------------------------------------------------
# AMT credit tracker (115JD)
# ---------------------------------------------------------------------------
@dataclass
class AmtcEntry:
    ay: str
    gross_credit: float = 0          # B1
    setoff_earlier: float = 0        # B2


class AmtcTracker:
    """Grid AMTC rows 9-22: balance = MAX(gross - earlier set-off, 0),
    utilised against current-year headroom, remainder carried forward."""

    def __init__(self, entries: List[AmtcEntry] = None):
        self.entries: List[AmtcEntry] = list(entries or [])

    def add(self, entry: AmtcEntry):
        self.entries.append(entry)

    def brought_forward_balance(self, current_ay: str) -> float:
        """Sum of balances still within 15 AYs [AMTC!I9 formula]."""
        from .carry_forward import ay_start
        cur = ay_start(current_ay)
        tot = 0.0
        for e in self.entries:
            if cur - ay_start(e.ay) < AMTC_CARRY_YEARS:
                tot += max(0.0, e.gross_credit - e.setoff_earlier)
        return tot

    def current_year(self, amt_tax: float, regular_tax: float,
                     current_ay: str, regime: int) -> Dict[str, float]:
        """Returns utilisation, credit arising and carry-forward."""
        if regime == C.REGIME_NEW:
            # [AMTC!K9] no credit carry under the new regime
            return {"brought_forward": 0.0, "utilised": 0.0,
                    "credit_arising": 0.0, "carried_forward": 0.0}
        bf = self.brought_forward_balance(current_ay)
        headroom = max(0.0, regular_tax - amt_tax)      # [AMTC!I6]
        utilised = min(bf, headroom)
        arising = max(0.0, amt_tax - regular_tax - utilised)  # [AMTC!G22]
        return {"brought_forward": bf, "utilised": excel_round(utilised),
                "credit_arising": excel_round(arising),
                "carried_forward": excel_round(bf - utilised + arising)}
