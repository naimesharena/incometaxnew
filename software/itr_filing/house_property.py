"""House Property computation - port of HP sheet + Income Details!B2.

Sources (ITR1_AY_26-27_V1.2.xlsm):
  HP!I25,I26,I27,I36-I39,I61-I76   per-property computation (up to 2 properties)
  Income Details!AO80-AO84          annual value / 30% / interest / arrears
  hidden sheet 'Schedule 24(b)'     interest on borrowed capital detail
"""
from dataclasses import dataclass, field
from .constants import REGIME_NEW, REGIME_OLD, HP_LOSS_CAP, HP_SELF_OCCUPIED_INTEREST_CAP
from .tax_engine import excel_round


@dataclass
class HouseProperty:
    """One house property (ITR-1 allows max 2, all in India)."""
    property_type: str = "(Select)"        # Self Occupied / Let Out / Deemed Let Out [DV AO77]
    gross_annual_value: float = 0          # gross rent received / lettable value
    unrealised_rent: float = 0
    tax_paid_local_authorities: float = 0
    interest_24b: float = 0                # from Schedule 24(b) detail
    arrears_unrealised_received: float = 0 # arrears/unrealised rent received, less 30%
    ownership_share_pct: float = 100
    address: str = ""
    city: str = ""
    state: str = ""
    pincode: int = 0

    @property
    def is_self_occupied(self) -> bool:
        return self.property_type[:1].upper() == "S"   # [AO82] MID(TypeOfHP,1,1)="S"

    def annual_value(self) -> int:
        """[HP!I25] =MAX((GAV - unrealised - tax),0)"""
        return excel_round(max((self.gross_annual_value
                                - self.unrealised_rent
                                - self.tax_paid_local_authorities), 0))

    def share_of_annual_value(self) -> int:
        """[HP!I26] =MAX(ROUND((share/100)*AV,0),0)"""
        return excel_round(max(excel_round((self.ownership_share_pct / 100) * self.annual_value()), 0))

    def standard_deduction_30pct(self) -> int:
        """[HP!I27] =MAX(ROUND(IF(30%*K26<0,0,30%*K26),0),0)"""
        av = self.share_of_annual_value()
        return excel_round(max(excel_round(0 if av < 0 else 0.30 * av), 0))

    def income(self) -> int:
        """[HP!I39] = K26 - (30% + interest) + arrears_after_30pct"""
        return excel_round(self.share_of_annual_value()
                           - (self.standard_deduction_30pct() + self.interest_24b)
                           + self.arrears_unrealised_received)


def allowed_interest(prop: HouseProperty, regime: int) -> int:
    """Income Details!AO82 - interest allowed for the Income Details summary:
    new regime + self-occupied -> 0
    old regime + self-occupied -> MIN(2,00,000, interest)
    let out & gross rent 0     -> 0
    let out                    -> full interest
    """
    if regime == REGIME_NEW and prop.is_self_occupied:
        return 0
    if regime == REGIME_OLD and prop.is_self_occupied:
        return excel_round(min(HP_SELF_OCCUPIED_INTEREST_CAP, prop.interest_24b))
    if not prop.is_self_occupied and prop.gross_annual_value == 0:
        return 0
    return excel_round(prop.interest_24b)


def total_hp_income(properties: list) -> int:
    """[HP!I76] =MAX(-200000, HP1 + HP2) - loss capped at 2,00,000."""
    total = sum(p.income() for p in properties)
    return excel_round(max(-HP_LOSS_CAP, total))


def hp_for_gti(properties: list, regime: int) -> int:
    """[AO111] GTI uses MAX(0, HP income) under the NEW regime (no loss set-off),
    actual (possibly negative) value under the OLD regime."""
    hp = total_hp_income(properties)
    return max(0, hp) if regime == REGIME_NEW else hp
