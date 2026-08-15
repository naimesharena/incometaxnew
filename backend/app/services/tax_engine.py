"""
Tax Engine for AY 2026-27 (FY 2025-26) - New & Old Regime
Based on CBDT slabs extracted from web search & analysis of Excel hidden SUMMARY + mdTaxCalc

New Regime (115BAC(1A) default):
0-4L nil, 4-8L 5%, 8-12L 10%, 12-16L 15%, 16-20L 20%, 20-24L 25%, >24L 30%
Old Regime:
0-2.5L nil (senior 3L, super senior 5L), 2.5-5L 5%, 5-10L 20%, >10L 30%
Standard deduction: New 75k (salaried), Old 50k
Rebate 87A: New up to 12L => rebate makes tax zero (max 60k), Old up to 5L => 12.5k
Surcharge: up to 50L nil, 50L-1Cr 10%, 1-2Cr 15%, 2-5Cr 25%, >5Cr new capped 25% old 37%
Health cess 4%

Interest 234A/B/C implemented per mdCalInterst234B logic
"""
from typing import Dict, Any, Tuple
import math
from datetime import date

# Constants AY26-27
NEW_SLABS = [
    (400000, 0.0),
    (800000, 0.05),
    (1200000, 0.10),
    (1600000, 0.15),
    (2000000, 0.20),
    (2400000, 0.25),
    (float('inf'), 0.30),
]

OLD_SLABS = [
    (250000, 0.0),
    (500000, 0.05),
    (1000000, 0.20),
    (float('inf'), 0.30),
]

OLD_SLABS_SENIOR = [
    (300000, 0.0),
    (500000, 0.05),
    (1000000, 0.20),
    (float('inf'), 0.30),
]

OLD_SLABS_SUPER_SENIOR = [
    (500000, 0.0),
    (1000000, 0.20),
    (float('inf'), 0.30),
]

SURCHARGE_SLABS_NEW = [
    (5000000, 0.0),
    (10000000, 0.10),
    (20000000, 0.15),
    (50000000, 0.25),
    (float('inf'), 0.25),  # capped
]

SURCHARGE_SLABS_OLD = [
    (5000000, 0.0),
    (10000000, 0.10),
    (20000000, 0.15),
    (50000000, 0.25),
    (float('inf'), 0.37),
]

def calculate_slab_tax(income: float, slabs) -> float:
    tax = 0.0
    prev_limit = 0
    for limit, rate in slabs:
        if income <= prev_limit:
            break
        taxable_in_bracket = min(income, limit) - prev_limit
        if taxable_in_bracket > 0:
            tax += taxable_in_bracket * rate
        prev_limit = limit
    return tax

def get_surcharge_rate(income: float, slabs) -> float:
    for limit, rate in slabs:
        if income <= limit:
            return rate
    return 0.0

def calculate_tax_old_regime(taxable_income: float, age: int, is_resident: bool = True) -> Dict[str, Any]:
    # Select slab based on age
    if age >= 80:
        slabs = OLD_SLABS_SUPER_SENIOR
    elif age >= 60:
        slabs = OLD_SLABS_SENIOR
    else:
        slabs = OLD_SLABS
    base_tax = calculate_slab_tax(taxable_income, slabs)
    # Rebate 87A old: if taxable_income <= 5L, rebate up to 12500
    rebate = 0
    if is_resident and taxable_income <= 500000:
        rebate = min(base_tax, 12500)
    tax_after_rebate = max(0, base_tax - rebate)
    # Surcharge
    surcharge_rate = get_surcharge_rate(taxable_income, SURCHARGE_SLABS_OLD)
    surcharge = tax_after_rebate * surcharge_rate
    # Cess 4%
    tax_plus_surcharge = tax_after_rebate + surcharge
    cess = tax_plus_surcharge * 0.04
    total = tax_plus_surcharge + cess

    return {
        "regime": "OLD",
        "taxable_income": taxable_income,
        "base_tax": round(base_tax, 2),
        "rebate87A": round(rebate, 2),
        "tax_after_rebate": round(tax_after_rebate, 2),
        "surcharge_rate": surcharge_rate,
        "surcharge": round(surcharge, 2),
        "cess": round(cess, 2),
        "total_tax": round(total, 2),
        "slabs_used": "OLD_SENIOR" if age>=60 else "OLD",
    }

def calculate_tax_new_regime(taxable_income: float, age: int = 30, is_resident: bool = True) -> Dict[str, Any]:
    base_tax = calculate_slab_tax(taxable_income, NEW_SLABS)
    # Rebate 87A new: if taxable_income <= 12L, total tax free (rebate = base_tax, max 60k but actually 60k is base tax at 12L)
    # At 12L, tax = 0*0 + 4L*5% =20k + 4L*10%=40k total 60k => rebate 60k => zero
    rebate = 0
    if is_resident and taxable_income <= 1200000:
        rebate = min(base_tax, 60000)  # actually rebate caps at base_tax, but max needed is 60k
        # if income <= 12L, rebate = base_tax => tax zero
        if taxable_income <= 1200000:
            rebate = base_tax
    tax_after_rebate = max(0, base_tax - rebate)
    surcharge_rate = get_surcharge_rate(taxable_income, SURCHARGE_SLABS_NEW)
    surcharge = tax_after_rebate * surcharge_rate
    tax_plus_surcharge = tax_after_rebate + surcharge
    cess = tax_plus_surcharge * 0.04
    total = tax_plus_surcharge + cess

    return {
        "regime": "NEW",
        "taxable_income": taxable_income,
        "base_tax": round(base_tax, 2),
        "rebate87A": round(rebate, 2),
        "tax_after_rebate": round(tax_after_rebate, 2),
        "surcharge_rate": surcharge_rate,
        "surcharge": round(surcharge, 2),
        "cess": round(cess, 2),
        "total_tax": round(total, 2),
        "slabs_used": "NEW_AY26-27",
    }

def compute_interest_234A(tax_due: float, filing_date: date, due_date: date, rate=0.01) -> float:
    """234A: delay filing - 1% per month from due date to filing date on tax due"""
    if filing_date <= due_date or tax_due <= 0:
        return 0.0
    months = (filing_date.year - due_date.year)*12 + (filing_date.month - due_date.month)
    if filing_date.day > due_date.day:
        months += 1
    if months < 1:
        months = 1
    return round(tax_due * rate * months, 0)

def compute_interest_234B(tax_payable: float, advance_paid: float, assessment_year_end: date = date(2026,3,31)) -> float:
    """234B: 1% per month if advance <90%"""
    # Simplified: if advance <90% of assessed tax, interest from April 1 AY to date of determination
    # Here we use simplified months=12 if shortfall
    # Real logic from mdCalInterst234B uses Balance_Interest table
    assessed = tax_payable
    advance = advance_paid
    if advance >= assessed * 0.9:
        return 0.0
    shortfall = assessed - advance
    # Assume interest for 12 months till 31st March next? Excel uses up to filing date
    # We'll compute 1% * months (assume 12 if not paid)
    months = 12
    return round(shortfall * 0.01 * months, 0)

def compute_interest_234C(advance_shortfall_quarters: Dict[str, float]) -> float:
    """
    advance_shortfall_quarters: dict like {"Q1": shortfall, "Q2":...}
    Rates: Q1 15% expected, Q2 45%, Q3 75%, Q4 100%
    Interest 1%*3 months per quarter on shortfall
    """
    # Simplified: if not enough advance timeline
    # Real CBDT: Q1 short 15% -> 1%*3, Q2 45% etc
    total_interest = 0.0
    # quarter rates
    rates = {"Q1": 3, "Q2": 3, "Q3": 3, "Q4": 1}  # months for interest
    for q, shortfall in advance_shortfall_quarters.items():
        months = rates.get(q, 1)
        total_interest += shortfall * 0.01 * months
    return round(total_interest, 0)

def compute_hra_exemption(basic_salary: float, da: float, hra_received: float, rent_paid: float, is_metro: bool) -> float:
    """
    HRA exemption u/s 10(13A) = min of:
    1. Actual HRA received
    2. Rent paid - 10% of salary (basic+DA)
    3. 50% of salary if metro, 40% non-metro
    Formula from hidden sheet Schedule EA 10(13A): G7, G10, G12
    """
    salary = basic_salary + da
    if salary <=0:
        return 0.0
    option1 = hra_received
    option2 = max(0, rent_paid - 0.1*salary)
    option3 = 0.5*salary if is_metro else 0.4*salary
    return round(min(option1, option2, option3), 0)

def compute_total_income(salary_income: float, hp_income: float, other_sources: float, 
                         exemption_hra: float = 0, standard_deduction: float = 50000,
                         deductions_80C: float = 0, deductions_other: float = 0,
                         regime: str = "OLD",
                         capital_gains: float = 0,
                         business_income: float = 0) -> Dict[str, float]:
    # Gross total income - includes CG and business for ITR-2/3/4
    gross = salary_income + hp_income + other_sources + capital_gains + business_income
    # Salary after HRA exemption and standard deduction
    if regime=="NEW":
        std = 75000
    else:
        std = standard_deduction
    taxable_salary = max(0, salary_income - exemption_hra - std)
    # Recalc gross with taxable salary + CG + business
    gross_recalc = taxable_salary + hp_income + other_sources + capital_gains + business_income
    # Deductions: only old regime allows 80C etc full, new regime limited
    if regime=="NEW":
        # new regime allows 80CCD(2) employer NPS, but for simplicity we allow 0 other deductions
        total_deductions = 0  # actually 80CCD2 allowed, but ignore for MVP
        # For AY26-27 new regime, still some deductions not allowed - we will keep as param but flag
        # In MVP we still apply if user provides but we will calculate both
        total_deductions_new = deductions_80C + deductions_other
        # Actually we should apply only limited, but for compatibility we will have two totals
        total_income_old = max(0, gross_recalc - (deductions_80C + deductions_other))
        total_income_new = max(0, gross_recalc - 0)  # new regime no deductions except std already
    else:
        total_deductions = deductions_80C + deductions_other
        total_income_old = max(0, gross_recalc - total_deductions)
        total_income_new = total_income_old  # placeholder
    
    return {
        "gross_total_income": round(gross,2),
        "taxable_salary_after_exemptions": round(taxable_salary,2),
        "gross_after_salary_adjustment": round(gross_recalc,2),
        "total_deductions_old": round(deductions_80C + deductions_other,2),
        "total_income_old": round(total_income_old,2) if regime=="OLD" else round(gross_recalc - (deductions_80C+deductions_other),2),
        "total_income_new": round(gross_recalc,2),  # new regime no 80 deductions
        "standard_deduction_applied": std,
        "hra_exemption": exemption_hra,
        "capital_gains": round(capital_gains,2),
        "business_income": round(business_income,2),
    }

def full_tax_computation(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    payload example:
    {
      "age": 30,
      "is_resident": True,
      "is_salaried": True,
      "regime": "NEW" or "OLD",
      "salary": 1200000,
      "basic": 600000,
      "da": 100000,
      "hra_received": 200000,
      "rent_paid": 250000,
      "is_metro": True,
      "hp_income": 0,
      "other_sources": 50000,
      "deductions_80C": 150000,
      "deductions_80D": 25000,
      "deductions_other": 25000, # 80G etc
      "tds": 50000,
      "advance_tax": 20000,
      "self_assessment": 0
    }
    Returns full breakdown matching SUMMARY sheet logic
    """
    age = payload.get("age", 30)
    is_resident = payload.get("is_resident", True)
    regime = payload.get("regime", "NEW")
    salary = payload.get("salary", 0)
    basic = payload.get("basic", salary*0.5)
    da = payload.get("da", 0)
    hra_received = payload.get("hra_received", 0)
    rent_paid = payload.get("rent_paid", 0)
    is_metro = payload.get("is_metro", False)
    hp_income = payload.get("hp_income", 0)
    other_sources = payload.get("other_sources", 0)
    capital_gains = payload.get("capital_gains", 0)
    business_income = payload.get("business_income", 0)
    ded_80C = payload.get("deductions_80C", 0)
    ded_other = payload.get("deductions_other", 0) + payload.get("deductions_80D", 0)

    # HRA only allowed in OLD regime, but compute anyway
    hra_exempt = 0
    if regime=="OLD" and hra_received>0:
        hra_exempt = compute_hra_exemption(basic, da, hra_received, rent_paid, is_metro)

    income_break = compute_total_income(salary, hp_income, other_sources, hra_exempt, 
                                        standard_deduction=50000, deductions_80C=ded_80C,
                                        deductions_other=ded_other, regime=regime,
                                        capital_gains=capital_gains,
                                        business_income=business_income)

    taxable_old = income_break["total_income_old"]
    taxable_new = income_break["total_income_new"]

    tax_old = calculate_tax_old_regime(taxable_old, age, is_resident)
    tax_new = calculate_tax_new_regime(taxable_new, age, is_resident)

    # Choose tax based on regime
    chosen_tax = tax_new if regime=="NEW" else tax_old
    taxable_income_chosen = taxable_new if regime=="NEW" else taxable_old

    # Interest 234A/B/C simplified
    tds = payload.get("tds",0)
    advance = payload.get("advance_tax",0)
    self_ass = payload.get("self_assessment",0)
    # Tax payable after TDS etc before interest
    tax_payable_before_interest = max(0, chosen_tax["total_tax"] - tds - advance - self_ass)

    # 234A: assume filing date 31/07/2027 vs due 31/07/2026, if delayed
    from datetime import date
    filing_date = payload.get("filing_date", date(2026,7,31))
    due_date = payload.get("due_date", date(2026,7,31))
    if isinstance(filing_date, str):
        filing_date = date.fromisoformat(filing_date)
    if isinstance(due_date, str):
        due_date = date.fromisoformat(due_date)

    interest_234A = compute_interest_234A(tax_payable_before_interest, filing_date, due_date)
    interest_234B = compute_interest_234B(chosen_tax["total_tax"], advance)
    # 234C simplified 0
    interest_234C = payload.get("interest_234C",0)

    total_interest = interest_234A + interest_234B + interest_234C
    total_payable = tax_payable_before_interest + total_interest
    # Refund if negative
    refund = 0
    payable = 0
    if total_payable <0:
        refund = abs(total_payable)
    else:
        payable = total_payable

    return {
        "income_breakdown": income_break,
        "tax_old_regime": tax_old,
        "tax_new_regime": tax_new,
        "chosen_regime": regime,
        "taxable_income": taxable_income_chosen,
        "chosen_tax": chosen_tax,
        "prepaid_taxes": {"tds": tds, "advance": advance, "self_assessment": self_ass, "total_prepaid": tds+advance+self_ass},
        "interest": {"234A": interest_234A, "234B": interest_234B, "234C": interest_234C, "total": total_interest},
        "final": {"payable": round(payable,2), "refund": round(refund,2), "total_tax_with_interest": round(chosen_tax["total_tax"]+total_interest,2)},
    }
