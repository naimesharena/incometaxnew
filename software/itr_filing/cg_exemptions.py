"""Capital-gains exemption sections (54 family) for AY 2026-27.

Sources (ITR2_AY_26-27_V1.3.xlsm):
  VBA line 31622   54EC: investment in notified bonds cannot exceed Rs 50 Lakh
  VBA line 31637   54F : investment cannot exceed Rs 10 crore
  CGDeductions.bas 54F grid: DateOfTransfer/CostOfHouse/DateOfPurc/AmtdepCG/
                   AmtDedClaimed -> pro-rata deduction
  CG sheet LTCG items B06d/B07d exemption grand totals feed the band matrix
"""
from dataclasses import dataclass
from .tax_engine import excel_round

CAP_54EC = 5000000           # [VBA 31622] Rs 50 Lakh
CAP_54F_INVESTMENT = 1000000000  # [VBA 31637] Rs 10 crore
CAP_54_TWO_HOUSE_GAIN = 20000000  # two-house option limit (Rs 2 crore gain)


def exemption_54ec(amount_invested: float) -> int:
    """Deduction u/s 54EC = investment in notified bonds, capped at 50L."""
    return excel_round(min(max(0.0, amount_invested), CAP_54EC))


def exemption_54f(capital_gain: float, net_consideration: float,
                  cost_of_new_house: float) -> int:
    """Pro-rata exemption u/s 54F:
    gain x MIN(cost of new house, Rs 10 crore cap) / net consideration.
    Full exemption when the whole net consideration (capped) is invested."""
    if net_consideration <= 0 or capital_gain <= 0:
        return 0
    invested = min(max(0.0, cost_of_new_house), CAP_54F_INVESTMENT)
    return excel_round(min(capital_gain, capital_gain * invested / net_consideration))


def exemption_54(capital_gain: float, cost_of_new_house: float,
                 second_house_used: bool = False,
                 already_owned_more_than_one: bool = False) -> int:
    """Deduction u/s 54 (residential house): MIN(gain, cost).
    AY 2026-27: a second house may be claimed only when the capital gain
    does not exceed Rs 2 crore (one-time option); the taxpayer must not
    own more than one other house on the transfer date."""
    if second_house_used and (capital_gain > CAP_54_TWO_HOUSE_GAIN
                              or already_owned_more_than_one):
        # second house not permitted -> only first-house cost qualifies;
        # caller should pass the eligible cost; we cap conservatively here
        return 0
    return excel_round(min(max(0.0, capital_gain), max(0.0, cost_of_new_house)))


def two_house_option_allowed(capital_gain: float) -> bool:
    return capital_gain <= CAP_54_TWO_HOUSE_GAIN


@dataclass
class ExemptionClaim:
    section: str            # "54", "54B", "54D", "54EC", "54EE", "54F", "54G", "54GB"
    amount: float = 0

    def allowed(self) -> int:
        if self.section == "54EC":
            return exemption_54ec(self.amount)
        return excel_round(max(0.0, self.amount))
