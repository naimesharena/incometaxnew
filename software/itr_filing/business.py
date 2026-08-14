"""Business & Profession engine for ITR-3 AY 2026-27.

Sources (ITR3_AY_26-27_V1.2.xlsm):
  PARTA_PL sheet + Profit_Loss.bas   net profit & adjustments
  Schedule DPM / DPM-DOA sheet       block-wise depreciation (SchDPM_DOA.bas,
                                     311 functions) - standard rules:
                                       * full rate on opening WDV + additions
                                         used >= 180 days
                                       * 50% of rate on additions used < 180 days
                                       * reduced by sale proceeds of assets
  BP sheet                           head-wise business income
  mdNOBBP.bas                        presumptive taxation 44AD/44ADA/44AE
"""
from dataclasses import dataclass, field
from typing import List
from .tax_engine import excel_round
from .itr4 import Business44AD, Profession44ADA, Business44AE


@dataclass
class DepreciationBlock:
    """One depreciation block [Schedule DPM rows]."""
    name: str = ""
    rate: float = 0.15                 # e.g. plant & machinery 15%
    opening_wdv: float = 0
    additions_full: float = 0          # used >= 180 days
    additions_half: float = 0          # used < 180 days -> 50% rate
    deletions: float = 0               # sale value of assets sold

    def depreciation(self) -> int:
        base = max(0.0, self.opening_wdv + self.additions_full - self.deletions)
        dep = base * self.rate + self.additions_half * self.rate * 0.5
        return excel_round(min(dep, base + self.additions_half))

    def closing_wdv(self) -> float:
        return max(0.0, self.opening_wdv + self.additions_full
                   + self.additions_half - self.deletions - self.depreciation())


@dataclass
class BusinessPL:
    """Profit & loss account of one business activity."""
    net_profit_as_per_pl: float = 0
    # add-backs (disallowed / inadmissible debits) [PARTA_PL DebitsToPL]
    inadmissible_expenses: float = 0     # s.40/43B etc.
    depreciation_as_per_pl: float = 0    # add back, recompute under IT rules
    other_additions: float = 0
    # deductions (admissible but not debited to P&L)
    admissible_not_debited: float = 0
    depreciation_blocks: List[DepreciationBlock] = field(default_factory=list)

    def income_before_depreciation(self) -> int:
        """Net profit + inadmissible debits + PL depreciation - admissible."""
        return excel_round(self.net_profit_as_per_pl
                           + self.inadmissible_expenses
                           + self.depreciation_as_per_pl
                           + self.other_additions
                           - self.admissible_not_debited)

    def total_depreciation(self) -> int:
        return excel_round(sum(b.depreciation() for b in self.depreciation_blocks))

    def income(self) -> int:
        """Business income after IT-rules depreciation [BP sheet net]."""
        return excel_round(self.income_before_depreciation() - self.total_depreciation())


@dataclass
class SpeculationBusiness:
    """Speculation business - losses set off only against speculation profit."""
    turnover: float = 0
    profit: float = 0          # net (may be negative)

    @property
    def income(self) -> float:
        return self.profit


@dataclass
class SpecifiedBusiness:
    """Specified business u/s 35AD."""
    turnover: float = 0
    profit: float = 0
    deduction_35ad: float = 0

    @property
    def income(self) -> float:
        return max(0.0, self.profit - self.deduction_35ad)


@dataclass
class BusinessIncomeSummary:
    """Feeds Schedule BP + CYLA buckets."""
    normal: int = 0            # BusinessIncOthThanSpec
    speculation: int = 0
    specified: int = 0
    presumptive_44ad: int = 0
    presumptive_44ada: int = 0
    presumptive_44ae: int = 0

    @property
    def total_bp(self) -> int:
        return (self.normal + max(0, self.speculation) + self.specified
                + self.presumptive_44ad + self.presumptive_44ada
                + self.presumptive_44ae)


def compute_business_income(
        businesses: List[BusinessPL],
        speculation: List[SpeculationBusiness] = None,
        specified: List[SpecifiedBusiness] = None,
        b44ad: Business44AD = None,
        p44ada: Profession44ADA = None,
        b44ae: Business44AE = None) -> BusinessIncomeSummary:
    """Aggregate all business tracks [ITR3ScheduleBP]."""
    normal = sum(b.income() for b in (businesses or []))
    spec = sum(s.income for s in (speculation or []))
    specified_inc = sum(s.income for s in (specified or []))
    return BusinessIncomeSummary(
        normal=excel_round(normal),
        speculation=excel_round(spec),
        specified=excel_round(specified_inc),
        presumptive_44ad=excel_round(b44ad.income) if b44ad else 0,
        presumptive_44ada=excel_round(p44ada.income) if p44ada else 0,
        presumptive_44ae=excel_round(b44ae.income) if b44ae else 0,
    )
