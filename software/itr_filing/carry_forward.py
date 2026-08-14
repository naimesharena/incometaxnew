"""Schedule CFL (carry-forward of losses) + Schedule BFLA (brought-forward
loss adjustment) - port of the ITR-2/ITR-3 'CFL' grid (rows 6-25) and the
BFLA matrix (rows 30-52 of 'CYLA - BFLA').

Loss categories [CFL!R4 columns]:
  hp           House property loss                 carry 8 AYs
  bp           Business loss (non-speculation)     carry 8 AYs (115BAB adj.)
  speculation  Speculative business loss           carry 4 AYs
  specified_bus Specified business loss (35AD)     carry forward (no limit)
  stcg         Short-term capital loss             carry 8 AYs
  ltcg         Long-term capital loss              carry 8 AYs
  racehorse    Race-horse activity loss            carry 4 AYs
  depreciation Unabsorbed depreciation (UD sheet)  indefinite

Expiry rule decoded from CFL!R25 (e.g. CFL_HP_Normal_2018 expires the AY-2018
losses in AY 2026-27): a loss of AY start year Y is available while
(current_ay_start - Y) < limit.
"""
from dataclasses import dataclass, field
from typing import Dict, List
from .setoff import ALL_BUCKETS
from .tax_engine import excel_round

LOSS_CATEGORIES = ["hp", "bp", "speculation", "specified_bus",
                   "stcg", "ltcg", "racehorse"]

CARRY_LIMITS = {"hp": 8, "bp": 8, "speculation": 4, "specified_bus": None,
                "stcg": 8, "ltcg": 8, "racehorse": 4}


def ay_start(ay: str) -> int:
    """'2023-24' -> 2023."""
    return int(str(ay).split("-")[0])


@dataclass
class LossEntry:
    ay: str
    hp: float = 0
    bp: float = 0
    speculation: float = 0
    specified_bus: float = 0
    stcg: float = 0
    ltcg: float = 0
    racehorse: float = 0
    depreciation: float = 0     # tracked separately, indefinite


class CFLTracker:
    """Holds prior-year losses and applies statutory expiry."""

    def __init__(self, entries: List[LossEntry] = None):
        self.entries: List[LossEntry] = list(entries or [])

    def add(self, entry: LossEntry):
        self.entries.append(entry)

    def available(self, current_ay: str) -> Dict[str, float]:
        """Totals per category still eligible in the current AY [CFL!R22
        totofbfloss.*CF8 sums]."""
        cur = ay_start(current_ay)
        tot = {c: 0.0 for c in LOSS_CATEGORIES}
        tot["depreciation"] = 0.0
        for e in self.entries:
            age = cur - ay_start(e.ay)
            for c in LOSS_CATEGORIES:
                limit = CARRY_LIMITS[c]
                if limit is None or age < limit:
                    tot[c] += max(0.0, getattr(e, c))
            tot["depreciation"] += max(0.0, e.depreciation)   # indefinite
        return tot


# BFLA set-off orders [matrix rows 34-50]
# BF business loss cascades BP -> SPEC -> SPECIFIED [W36/W37 chain];
# BF STCG loss may absorb LTCG bands too (s.70) [AK41/AK42 links];
# BF LTCG loss only against LTCG bands.
BFLA_ORDERS = {
    "hp": ["HP"],                                   # only against HP income
    "bp": ["BP", "SPEC", "SPECIFIED"],              # business-type heads only
    "speculation": ["SPEC"],                        # s.73 speculation profit only
    "specified_bus": ["SPECIFIED"],                 # s.35AD same head only
    "stcg": ["STCG15", "STCG20", "STCG30", "STCG_RATE", "STCG_DTAA",
             "LTCG10", "LTCG125", "LTCG20", "LTCG_DTAA"],             # S41->AK
    "ltcg": ["LTCG10", "LTCG125", "LTCG20", "LTCG_DTAA"],              # S44 cascade
    "racehorse": ["RACEHORSE"],
}
# unabsorbed depreciation: any head except salary [H35:H38,H40:H43,H45,H47:H50]
DEPRECIATION_ORDER = [b for b in ALL_BUCKETS if b != "SAL"]


@dataclass
class BFLAResult:
    income_after_setoff: Dict[str, float] = field(default_factory=dict)
    setoff_by_category: Dict[str, float] = field(default_factory=dict)
    depreciation_setoff: float = 0
    remaining: Dict[str, float] = field(default_factory=dict)  # -> CFL current row


def bfla(incomes_after_cyla: Dict[str, float], bf_losses: Dict[str, float],
         depreciation: float = 0.0) -> BFLAResult:
    """Brought-forward loss adjustment. Losses of each category absorb income
    in the category's order; unabsorbed amounts remain to carry forward
    [BFLA!R51-52]."""
    avail = {b: max(0.0, float(incomes_after_cyla.get(b, 0))) for b in ALL_BUCKETS}
    res = BFLAResult()

    def allocate(loss: float, order: List[str]) -> float:
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

    # capital losses first (utility cascades STCG then LTCG) [S41/S44 columns]
    for cat in ["stcg", "ltcg"]:
        loss = max(0.0, float(bf_losses.get(cat, 0)))
        res.setoff_by_category[cat] = allocate(loss, BFLA_ORDERS[cat])
        res.remaining[cat] = loss - res.setoff_by_category[cat]

    for cat in ["hp", "bp", "speculation", "specified_bus", "racehorse"]:
        loss = max(0.0, float(bf_losses.get(cat, 0)))
        res.setoff_by_category[cat] = allocate(loss, BFLA_ORDERS[cat])
        res.remaining[cat] = loss - res.setoff_by_category[cat]

    # unabsorbed depreciation last (any head except salary)
    dep = max(0.0, float(depreciation))
    res.depreciation_setoff = allocate(dep, DEPRECIATION_ORDER)
    res.remaining["depreciation"] = dep - res.depreciation_setoff

    res.income_after_setoff = {b: excel_round(v) for b, v in avail.items()}
    return res


def total_income_after_bfla(result: BFLAResult) -> int:
    """sheet16.IncomeOfCurrYrAftCYLABFLA = SUM(I34:I50)."""
    return excel_round(sum(result.income_after_setoff.values()))
