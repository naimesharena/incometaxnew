"""DTAA relief u/s 90 / 90A / 91 - double taxation relief.

Port of the FSI / PTI / TR schedule logic (ITR-2/ITR-3): relief for income
taxed both in India and a treaty/non-treaty country is the LOWER of
  (a) tax actually paid abroad on that income, and
  (b) the Indian tax attributable to that income (incremental tax method:
      tax(total) - tax(total - foreign income), which is what the utility's
      PTI/TR computation yields via its slab re-runs).
"""
from .tax_engine import excel_round


def relief_90(foreign_tax_paid: float, indian_tax_on_total: float,
              indian_tax_excluding_foreign: float) -> int:
    """s.90/90A relief (treaty countries): MIN(foreign tax, Indian tax on
    the doubly-taxed income)."""
    indian_attributable = max(0.0, indian_tax_on_total - indian_tax_excluding_foreign)
    return excel_round(min(max(0.0, foreign_tax_paid), indian_attributable))


def relief_91(foreign_tax_paid: float, indian_tax_on_total: float,
              indian_tax_excluding_foreign: float, indian_rate: float,
              foreign_rate: float) -> int:
    """s.91 relief (no treaty): credit at the LOWER of the Indian rate and
    the foreign rate applied to the doubly-taxed income."""
    doubly_taxed = max(0.0, indian_tax_on_total - indian_tax_excluding_foreign)
    if doubly_taxed <= 0:
        return 0
    if indian_rate + foreign_rate <= 0:
        return 0
    rate = min(indian_rate, foreign_rate)
    # income on which both taxes apply, derived from the higher side
    base_income = doubly_taxed / max(indian_rate, 1e-9) if indian_rate >= foreign_rate \
        else doubly_taxed / max(foreign_rate, 1e-9)
    return excel_round(min(base_income * rate, min(foreign_tax_paid, doubly_taxed)))
