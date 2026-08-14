"""Schedule CYLA - Current Year Loss Adjustment (set-off engine).

Faithful port of the ITR-2/ITR-3 'CYLA - BFLA' sheet matrix
(cells G8:G24, H9:H24, I8:I24 and helper columns S/T/U) plus
CYLACalculations.bas ordering.

AY 2026-27 income buckets (sheet rows ii..xii):
  SAL, HP, BP, SPEC (speculation), SPECIFIED (35AD),
  STCG15, STCG20, STCG30, STCG_RATE, STCG_DTAA,
  LTCG10 (112A legacy), LTCG125 (112A AY26-27), LTCG20, LTCG_DTAA,
  OS (normal rates), RACEHORSE, OS_DTAA

Loss sources (row i): HP loss (capped 2,00,000), BP loss (non-speculation),
OS loss (no race-horse).
"""
from dataclasses import dataclass, field
from typing import Dict, List
from . import constants as C
from .tax_engine import excel_round

ALL_BUCKETS = ["SAL", "HP", "BP", "SPEC", "SPECIFIED",
               "STCG15", "STCG20", "STCG30", "STCG_RATE", "STCG_DTAA",
               "LTCG10", "LTCG125", "LTCG20", "LTCG_DTAA",
               "OS", "RACEHORSE", "OS_DTAA"]

# HP loss set-off order [G8, G12:G17, G19, G21:G24] - no BP row (G10 absent)
HP_LOSS_ORDER = ["SAL", "SPECIFIED",
                 "STCG15", "STCG20", "STCG30", "STCG_RATE", "STCG_DTAA",
                 "LTCG10", "LTCG125", "LTCG20", "LTCG_DTAA",
                 "OS", "RACEHORSE"]

# BP loss set-off order [H9, H14:H17, H19, H21:H24] - not against SAL (H8 absent)
BP_LOSS_ORDER = ["HP",
                 "STCG20", "STCG30", "STCG_RATE", "STCG_DTAA",
                 "LTCG125", "LTCG20", "LTCG_DTAA",
                 "OS", "RACEHORSE"]

# OS (normal) loss set-off order [I8:I12, I14:I17, I19, I21:I24]
OS_LOSS_ORDER = ["SAL", "HP", "BP", "SPEC", "SPECIFIED",
                 "STCG20", "STCG30", "STCG_RATE", "STCG_DTAA",
                 "LTCG125", "LTCG20", "LTCG_DTAA",
                 "OS", "RACEHORSE"]

HP_LOSS_CAP = 200000     # [CYLA!S6 cyla.TotHPlossCurYr1 via HP.REStwolakh; G25 cap]


@dataclass
class CYLAResult:
    income_after_setoff: Dict[str, float] = field(default_factory=dict)
    hp_loss_setoff: float = 0
    bp_loss_setoff: float = 0
    os_loss_setoff: float = 0
    hp_loss_remaining: float = 0      # carried to CFL [G26 BalHPlossCurYrAftSetoff]
    bp_loss_remaining: float = 0
    os_loss_remaining: float = 0


def cyla(incomes: Dict[str, float], hp_loss: float, bp_loss: float,
         os_loss: float, regime: int = C.REGIME_OLD) -> CYLAResult:
    """Sequential set-off exactly as the utility matrix allocates it.

    incomes : {bucket: positive income} (losses within buckets must be 0;
              HP head loss comes via hp_loss, OS normal loss via os_loss)
    regime  : NEW regime -> HP loss cannot be carried forward unabsorbed
              inter-head beyond flooring [G26 IF(bacValue=1,0,...)] and GTI
              floors HP at 0 upstream.
    """
    avail = {b: max(0.0, float(incomes.get(b, 0))) for b in ALL_BUCKETS}
    res = CYLAResult()

    def allocate(loss: float, order: List[str]) -> float:
        """Set off `loss` across buckets in order; returns amount absorbed."""
        absorbed = 0.0
        for b in order:
            if loss <= 0:
                break
            take = min(loss, avail[b])
            if take > 0:
                avail[b] -= take
                loss -= take
                absorbed += take
        return absorbed

    # 1) HP loss (cap 2L) [cyla.TotHPlossCurYr1 + G25 MIN(...,200000)]
    hp_loss = min(max(0.0, hp_loss), HP_LOSS_CAP)
    res.hp_loss_setoff = allocate(hp_loss, HP_LOSS_ORDER)
    res.hp_loss_remaining = hp_loss - res.hp_loss_setoff
    if regime == C.REGIME_NEW:
        res.hp_loss_remaining = 0          # [G26] IF(bacValue=1,0,...)

    # 2) BP loss (non-speculation)
    bp_loss = max(0.0, bp_loss)
    res.bp_loss_setoff = allocate(bp_loss, BP_LOSS_ORDER)
    res.bp_loss_remaining = bp_loss - res.bp_loss_setoff

    # 3) OS loss (no race-horse)
    os_loss = max(0.0, os_loss)
    res.os_loss_setoff = allocate(os_loss, OS_LOSS_ORDER)
    res.os_loss_remaining = os_loss - res.os_loss_setoff

    res.income_after_setoff = {b: excel_round(v) for b, v in avail.items()}
    return res


def total_income_after_cyla(result: CYLAResult) -> int:
    """Column 4 (income remaining after set-off) summed across heads."""
    return excel_round(sum(result.income_after_setoff.values()))
