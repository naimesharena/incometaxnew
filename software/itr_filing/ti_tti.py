"""Part B - TI & TTI: total-income aggregation + final tax computation for
ITR-2 / ITR-3 (AY 2026-27).

Ties together: head-wise incomes, CYLA/BFLA results, CG waterfall bands,
Chapter VI-A, slabs, special rates (111A/112A bands), rebate 87A, regime-aware
surcharge tiers, cess, AMT comparison (115JC) and interest/fees.

Sources (ITR2_AY_26-27_V1.3.xlsm):
  Part B - TI TTI sheet      aggregation rows
  'Tax Calculated' sheet     slab + special-rate engine (1,981 formulas)
  SPI - SI sheet             special-rate income split
  AMT/AMTC sheets            115JC/115JD (see amt.py)
"""
from dataclasses import dataclass, field
from typing import Dict
from . import constants as C
from .tax_engine import (tax_new_regime, tax_old_regime, rebate_87a,
                         education_cess, excel_round)
from .capital_gains import STCG_BANDS, LTCG_BANDS

# AY 2026-27 special rates per CG band
BAND_RATES = {
    "STCG15": 0.15,      # legacy band (grid 1)
    "STCG20": 0.20,      # 111A STT-paid
    "STCG30": 0.30,      # equity-oriented fund units etc.
    "LTCG10": 0.10,      # legacy 112A band
    "LTCG125": 0.125,    # 112A AY 26-27 / immovable without indexation
    "LTCG20": 0.20,
}
LTCG_112A_EXEMPT = 125000    # first 1.25L of 112A gains exempt

# surcharge tiers [Tax Calculated sheet surcharge blocks]
SURCHARGE_OLD = [(50000000, 0.37), (20000000, 0.25), (10000000, 0.15), (5000000, 0.10)]
SURCHARGE_NEW = [(50000000, 0.25), (20000000, 0.25), (10000000, 0.15), (5000000, 0.10)]


@dataclass
class TITTIInput:
    income_heads: Dict[str, float] = field(default_factory=dict)  # after CYLA/BFLA
    cg_net_by_band: Dict[str, float] = field(default_factory=dict)
    total_income_normal: float = 0        # slab-taxable (excl. special-rate CG)
    ti_incl_112a: float = 0               # for old-regime 87A test
    regime: int = C.REGIME_NEW
    age: int = 40
    relief_89: float = 0
    relief_90_91: float = 0
    amt_payable: float = 0                # from amt.py (0 if N/A)
    tds: float = 0
    tcs: float = 0
    advance_tax: float = 0


def special_rate_tax(cg_net_by_band: Dict[str, float]) -> Dict[str, int]:
    """Tax on each special-rate CG band. 112A buckets get the 1.25L exempt
    threshold before the 12.5% rate."""
    out = {}
    for band, rate in BAND_RATES.items():
        amt = max(0.0, float(cg_net_by_band.get(band, 0)))
        if amt <= 0:
            continue
        if band in ("LTCG125", "LTCG10"):
            taxable = max(0.0, amt - LTCG_112A_EXEMPT)
        else:
            taxable = amt
        out[band] = excel_round(taxable * rate)
    return out


def surcharge(total_income: float, base_tax: float, regime: int) -> int:
    """Regime-aware surcharge tiers; new regime capped at 25%."""
    tiers = SURCHARGE_NEW if regime == C.REGIME_NEW else SURCHARGE_OLD
    for threshold, rate in tiers:
        if total_income > threshold:
            return excel_round(base_tax * rate)
    return 0


def compute_ti_tti(inp: TITTIInput) -> Dict[str, object]:
    """Final computation: normal slabs + special rates - 87A + surcharge +
    cess, then MAX(regular, AMT) and reliefs."""
    # 1) normal slab tax
    ti_normal = max(0.0, inp.total_income_normal)
    if inp.regime == C.REGIME_NEW:
        tax_normal = tax_new_regime(ti_normal)
    else:
        tax_normal = tax_old_regime(ti_normal, inp.age)

    # 2) special-rate taxes on CG bands
    sp = special_rate_tax(inp.cg_net_by_band)
    tax_special = sum(sp.values())

    # 3) rebate 87A
    total_ti = ti_normal + sum(max(0.0, v) for v in inp.cg_net_by_band.values())
    ti_for_rebate_excl_112a = ti_normal + sum(
        max(0.0, inp.cg_net_by_band.get(b, 0)) for b in
        ["STCG15", "STCG20", "STCG30", "STCG_RATE", "STCG_DTAA", "LTCG20", "LTCG_DTAA"])
    rebate = rebate_87a(tax_normal, ti_for_rebate_excl_112a,
                        inp.ti_incl_112a or total_ti, inp.regime)

    # 4) surcharge (new regime capped 25%) + cess
    base = max(0, tax_normal - rebate) + tax_special
    sur = surcharge(total_ti, base, inp.regime)
    cess = education_cess(base + sur)

    regular_total = base + sur + cess

    # 5) AMT comparison [Tax Calculated / AMT sheets]
    amt_total = inp.amt_payable + education_cess(inp.amt_payable) if inp.amt_payable else 0
    gross_tax = max(regular_total, amt_total)
    amt_applied = amt_total > regular_total

    # 6) reliefs u/s 89 / 90 / 91
    net_tax = excel_round(max(0.0, gross_tax - inp.relief_89 - inp.relief_90_91))

    return {
        "TaxNormal": excel_round(tax_normal),
        "TaxSpecialRates": sp,
        "TaxSpecialTotal": excel_round(tax_special),
        "Rebate87A": rebate,
        "Surcharge": sur,
        "Cess": cess,
        "RegularTotal": excel_round(regular_total),
        "AMTTotal": excel_round(amt_total),
        "AMTApplied": amt_applied,
        "GrossTaxLiability": excel_round(gross_tax),
        "NetTaxLiability": net_tax,
    }
