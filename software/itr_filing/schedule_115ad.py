"""Schedule 115AD(1)(iii) proviso - NON-RESIDENTS (FII/FPI): capital gains on
sale of equity shares/units of an Indian company (listed).

Source (ITR2_AY_26-27_V1.3.xlsm):
  'Schedule 115AD(1)(iii) proviso' sheet rows 4-14 + md115AD.bas
  Columns: 1a acquisition bucket (on/before vs after 31/01/2018),
           1b transfer bucket (before vs on/after 23/07/2024),
           6  Full value of consideration = MAX(0, qty x sale price)
           7  Cost without indexation = MAX(actual cost, grandfathered)
           10 LTCG asset-acquired rule = MIN(sale, cost basis)
           14 gain = sale - deductions
  Row 12 (i)  total gains where transfer BEFORE 23/07/2024
  Row 13 (ii) total gains where transfer ON/AFTER 23/07/2024
  Row 14 (iii) LTCG u/s 112A r/w 115AD(1)(b)(iii) proviso

Effect: for eligible FPIs the gains are taxed at the special capital-gains
rates (111A/112A bands) instead of the 40% FPI rate.
"""
from dataclasses import dataclass
from datetime import date
from typing import List
from .tax_engine import excel_round
from .capital_gains import grandfathered_cost_112a

GRANDFATHER_DATE = date(2018, 1, 31)
TRANSFER_REGIME_DATE = date(2024, 7, 23)
STCG_RATE_115AD = 0.20      # 111A-equivalent for listed securities
LTCG_RATE_115AD = 0.125     # 112A-equivalent (AY 2026-27)
LTCG_EXEMPT_112A = 125000


@dataclass
class FPIEquityTransaction:
    acquired_on_or_before_31jan2018: bool = True
    shares: float = 0
    sale_price_per_share: float = 0
    cost_per_share: float = 0
    fmv_per_share_31jan2018: float = 0
    transfer_expenses: float = 0
    transfer_date: date = None

    def full_value_consideration(self) -> int:
        """[Col 6] MAX(0, qty x sale price)."""
        return excel_round(max(0.0, self.shares * self.sale_price_per_share))

    def cost_of_acquisition(self) -> int:
        """[Col 7/8] grandfathering for pre-01/02/2018 acquisitions:
        MAX(actual cost, MIN(FMV 31/01/2018, sale consideration))."""
        actual = self.shares * self.cost_per_share
        if self.acquired_on_or_before_31jan2018 and self.fmv_per_share_31jan2018 > 0:
            total_fmv = self.shares * self.fmv_per_share_31jan2018
            basis = grandfathered_cost_112a(actual, total_fmv,
                                            self.shares * self.sale_price_per_share)
        else:
            basis = actual
        return excel_round(max(0.0, basis))

    def is_long_term(self) -> bool:
        """Listed equity: > 12 months holding."""
        return True  # schedule applies to LTCG transfers; STCG handled at 111A

    def transferred_after_regime_change(self) -> bool:
        """[Col 1b] on/after 23/07/2024 -> new 12.5% bucket."""
        return bool(self.transfer_date) and self.transfer_date >= TRANSFER_REGIME_DATE

    def gain(self) -> int:
        """[Col 14] sale consideration - cost - transfer expenses (min 0)."""
        return excel_round(max(0.0, self.full_value_consideration()
                               - self.cost_of_acquisition()
                               - self.transfer_expenses))


@dataclass
class Schedule115AD:
    transactions: List[FPIEquityTransaction] = None

    def summary(self) -> dict:
        txns = self.transactions or []
        before = sum(t.gain() for t in txns if not t.transferred_after_regime_change())
        after = sum(t.gain() for t in txns if t.transferred_after_regime_change())
        total = before + after
        taxable = max(0, total - (LTCG_EXEMPT_112A if after else 0))
        return {
            "TotalSaleValue": excel_round(sum(t.full_value_consideration() for t in txns)),
            "TotalCost": excel_round(sum(t.cost_of_acquisition() for t in txns)),
            "GainsTransferBefore23Jul2024": excel_round(before),   # row 12 (i)
            "GainsTransferOnAfter23Jul2024": excel_round(after),   # row 13 (ii)
            "LTCGUs112A_115ADProviso": excel_round(total),         # row 14 (iii)
            "TaxableAfter112AExempt": excel_round(taxable),
            "TaxAtSpecialRates": excel_round(taxable * LTCG_RATE_115AD),
        }
