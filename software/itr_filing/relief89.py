"""Relief u/s 89 (Form 10E) - arrears/salary received in advance.

Standard algorithm (as applied by the utility's D6 computation with Form 10E
data): relief = additional tax in the current year due to arrears MINUS the
additional tax that would have arisen in the year to which the arrears relate.

relief = [tax(CY income + arrears) - tax(CY income)]
       - [tax(PY income + arrears) - tax(PY income)]
where the applicable slabs depend on the regime/age for each year.
"""
from .tax_engine import tax_new_regime, tax_old_regime, excel_round
from . import constants as C


def relief_89(arrears: float,
              current_year_income_excl_arrears: float,
              prior_year_income_excl_arrears: float,
              prior_year: str = "2025",
              regime: int = C.REGIME_NEW,
              age_current: int = 40,
              age_prior: int = None) -> int:
    """Compute relief u/s 89. Prior year defaults use the same slab engine
    (AY-specific slabs for earlier years should override via the callbacks)."""
    if arrears <= 0:
        return 0
    age_prior = age_prior if age_prior is not None else max(0, age_current - 1)

    def tax(income: float, age: int) -> int:
        if regime == C.REGIME_NEW:
            return tax_new_regime(max(0.0, income))
        return tax_old_regime(max(0.0, income), age)

    cy_delta = (tax(current_year_income_excl_arrears + arrears, age_current)
                - tax(current_year_income_excl_arrears, age_current))
    py_delta = (tax(prior_year_income_excl_arrears + arrears, age_prior)
                - tax(prior_year_income_excl_arrears, age_prior))
    relief = max(0, cy_delta - py_delta)
    return excel_round(relief)
