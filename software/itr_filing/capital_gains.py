"""Capital Gains engine for ITR-2/ITR-3 AY 2026-27.

Sources (ITR2_AY_26-27_V1.3.xlsm):
  CG sheet rows 498-523          IHLA band matrix: intra-CG waterfall set-off
                                 (STCG bands 20/30/applicable/DTAA, LTCG 12.5/DTAA,
                                 plus legacy 15%/10% bands of grid 1)
  Schedule 112A + LTCG items B1a-B1g  grandfathering for listed STT-paid assets
  VBA SchCG.bas / CG_Calc.bas    holding-period classification & computations

AY 2026-27 notes (Finance (No.2) Act 2024 + Finance Act 2025):
  - holding periods unified: 12 months (securities), 24 months (immovable &
    other) for assets acquired on/after 23/07/2024
  - indexation removed; LTCG taxed @ 12.5% (112A) / 20% bands
  - STCG @ 20% (111A STT-paid) / other bands per asset class
"""
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List
from .tax_engine import excel_round

# band names used across CYLA/setoff modules
STCG_BANDS = ["STCG15", "STCG20", "STCG30", "STCG_RATE", "STCG_DTAA"]
LTCG_BANDS = ["LTCG10", "LTCG125", "LTCG20", "LTCG_DTAA"]
ALL_CG_BANDS = STCG_BANDS + LTCG_BANDS

HOLDING_SHORT_SEC_MONTHS = 12      # listed securities
HOLDING_SHORT_OTHER_MONTHS = 24    # immovable & other assets (post 23/07/2024)
UNIFIED_HOLDING_CUTOFF = date(2024, 7, 23)


# ---------------------------------------------------------------------------
# 112A grandfathering
# ---------------------------------------------------------------------------
def grandfathered_cost_112a(cost: float, fmv_31jan2018: float,
                            sale_consideration: float) -> float:
    """Cost basis for listed STT-paid equity acquired before 01/02/2018:
    MAX(actual cost, MIN(FMV as on 31/01/2018, sale consideration)).
    (If FMV exceeds sale, the gain is nil - cost equals sale.)"""
    if fmv_31jan2018 <= 0:
        return cost
    return max(cost, min(fmv_31jan2018, sale_consideration))


# ---------------------------------------------------------------------------
# Asset model
# ---------------------------------------------------------------------------
@dataclass
class CapitalAsset:
    asset_class: str = "immovable"      # securities | immovable | other
    listed_stt_paid: bool = False       # 112A category
    acquisition_date: date = None
    transfer_date: date = None
    sale_consideration: float = 0
    cost_of_acquisition: float = 0
    improvement_cost: float = 0
    transfer_expenses: float = 0
    fmv_31jan2018: float = 0            # only for pre-01/02/2018 listed assets
    exemptions: Dict[str, float] = field(default_factory=dict)  # s.54/54EC/...

    def months_held(self) -> int:
        if not self.acquisition_date or not self.transfer_date:
            return 0
        return (self.transfer_date.year - self.acquisition_date.year) * 12 \
            + (self.transfer_date.month - self.acquisition_date.month)

    def is_long_term(self) -> bool:
        """[SchCG holding-period rules] unified periods for acquisitions
        on/after 23/07/2024."""
        limit = HOLDING_SHORT_SEC_MONTHS if self.asset_class == "securities" \
            else HOLDING_SHORT_OTHER_MONTHS
        return self.months_held() > limit

    def cost_basis(self) -> float:
        if self.listed_stt_paid and self.fmv_31jan2018 > 0 \
                and self.acquisition_date and self.acquisition_date < date(2018, 2, 1):
            return grandfathered_cost_112a(self.cost_of_acquisition,
                                           self.fmv_31jan2018, self.sale_consideration)
        return self.cost_of_acquisition

    def gain(self) -> float:
        """Full value of gain before exemptions."""
        return (self.sale_consideration - self.cost_basis()
                - self.improvement_cost - self.transfer_expenses)

    def total_exemptions(self) -> float:
        return sum(self.exemptions.values())

    def gain_after_exemptions(self) -> float:
        g = self.gain()
        ex = min(self.total_exemptions(), max(0.0, g))
        return g - ex

    def band(self) -> str:
        """Rate band per asset category (resident normal rates)."""
        if not self.is_long_term():
            if self.asset_class == "securities" and self.listed_stt_paid:
                return "STCG20"          # 111A STT-paid -> 20% AY 26-27
            if self.asset_class == "securities":
                return "STCG30"          # equity-oriented fund units etc.
            return "STCG_RATE"           # normal slab rates
        # long term
        if self.listed_stt_paid:
            return "LTCG125"             # 112A -> 12.5%
        if self.asset_class == "immovable":
            return "LTCG125"             # 12.5% without indexation AY 26-27
        return "LTCG20"


# ---------------------------------------------------------------------------
# Intra-CG waterfall set-off [CG!rows 498-523]
# ---------------------------------------------------------------------------
def cg_current_year_setoff(net_by_band: Dict[str, float]) -> Dict[str, float]:
    """Cross-set-off of current-year capital gains/losses across bands.

    net_by_band: {band: net gain (+) or net loss (-)} computed after
    intra-band netting of the schedule items (grid rows take MAX(0,..) gains
    and ABS(MIN(0,..)) losses separately).

    Waterfall rules decoded from the IHLA matrix:
      * STCL of each band absorbs other STCG bands in band order, then
        LTCG bands (112A 12.5 first, then 20, then DTAA); own band already
        netted upstream.
      * LTCL absorbs only LTCG bands.
    Returns {band: net amount remaining} - positives are the CYLA buckets,
    negatives are carried to CFL (STCGLossCF / LTCGLossCF).
    """
    avail = {b: max(0.0, float(net_by_band.get(b, 0))) for b in ALL_CG_BANDS}
    losses = [(b, -min(0.0, float(net_by_band.get(b, 0)))) for b in ALL_CG_BANDS]

    def absorb(loss_band: str, loss: float, targets: List[str]) -> float:
        absorbed = 0.0
        for t in targets:
            if t == loss_band or loss <= 0:
                continue
            take = min(loss, avail[t])
            if take > 0:
                avail[t] -= take
                loss -= take
                absorbed += take
        return absorbed

    remaining = {b: 0.0 for b in ALL_CG_BANDS}
    # short-term losses first [grid J..N columns]
    for b in STCG_BANDS:
        loss = dict(losses).get(b, 0.0)
        if loss > 0:
            targets = [x for x in STCG_BANDS + LTCG_BANDS]
            absorbed = absorb(b, loss, targets)
            remaining[b] = loss - absorbed
    # long-term losses [grid O..R columns] - LTCG targets only
    for b in LTCG_BANDS:
        loss = dict(losses).get(b, 0.0)
        if loss > 0:
            absorbed = absorb(b, loss, LTCG_BANDS)
            remaining[b] = loss - absorbed

    out = {b: avail[b] - remaining[b] for b in ALL_CG_BANDS}
    # express as net per band: positive income, negative carry-forward loss
    return {b: excel_round(avail[b]) if avail[b] > 0 else -excel_round(remaining[b])
            for b in ALL_CG_BANDS}


def compute_cg_summary(assets: List[CapitalAsset]) -> Dict[str, float]:
    """Aggregate assets per band (intra-band netting) then waterfall."""
    net = {b: 0.0 for b in ALL_CG_BANDS}
    for a in assets:
        net[a.band()] += a.gain_after_exemptions()
    return cg_current_year_setoff(net)
