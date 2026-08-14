"""ITR-4 (SUGAM) AY 2026-27 - Individuals/HUF/Firms (excl. LLP), presumptive
taxation u/s 44AD/44ADA/44AE.

Sources (ITR4_AY_26-27_V1.1.xlsm):
  hidden TaxCalc sheet          slab engine per entity (GrpA/B/C, HUF, Firm),
                                surcharge + marginal relief, cess, 44AD tax split
  VBA mdNOBBP.bas (lines 319-323) 44AD rates: 6% (a/c payee & electronic) and
                                8% (other modes); 44AD cap Rs 3,00,00,000
  VBA validation messages       44AE minimums: Rs 1,000/MT/month (heavy goods
                                vehicles), Rs 7,500/month (others);
                                eligibility: total income <= Rs 50,00,000
"""
from dataclasses import dataclass, field
from typing import List, Optional
from datetime import date

from . import constants as C
from .tax_engine import excel_round, tax_new_regime, calculate_age
from .house_property import HouseProperty, total_hp_income, hp_for_gti
from . import chapter_via as via
from .interest import interest_234a, interest_234b, interest_234c, fee_234f

ENTITY_INDIVIDUAL = "I"
ENTITY_HUF = "H"
ENTITY_FIRM = "F"

CAP_44AD = 30000000          # [VBA 11125] presumptive income u/s 44AD <= 3 Cr
CAP_44ADA_GROSS = 7500000    # professional gross receipts limit for 44ADA
RATE_44AD_DIGITAL = 0.06     # [VBA 319] a/c payee cheque / electronic clearing
RATE_44AD_OTHER = 0.08       # [VBA 320/323] any other mode
RATE_44ADA = 0.50            # professionals: 50% of gross receipts
ITR4_TI_CAP = 5000000        # [VBA 15467] total income upto Rs 50 Lakh


# ---------------------------------------------------------------------------
# Presumptive business models
# ---------------------------------------------------------------------------
@dataclass
class Business44AD:
    """BP sheet E1/E2 - business under 44AD."""
    turnover_digital: float = 0       # E1(a) a/c payee / electronic clearing
    turnover_other: float = 0         # E1(b+c) other modes
    income_digital: float = 0         # E2(a) declared (>= 6%)
    income_other: float = 0           # E2(b) declared (>= 8% of b+c)

    @property
    def income(self) -> float:
        return self.income_digital + self.income_other

    def min_income(self) -> float:
        return (RATE_44AD_DIGITAL * self.turnover_digital
                + RATE_44AD_OTHER * self.turnover_other)

    def validation_errors(self) -> list:
        errs = []
        if self.income > CAP_44AD:
            errs.append("[VBA 11125] Presumptive income under section 44AD "
                        "cannot exceed 30000000")
        if self.income_digital and self.income_digital < RATE_44AD_DIGITAL * self.turnover_digital:
            errs.append("[VBA 319] 44AD E2(a) must be at least 6% of gross receipts "
                        "through a/c payee or electronic clearing system "
                        "(else tax audit u/s 44AB -> use ITR-3/5)")
        if self.income_other and self.income_other < RATE_44AD_OTHER * self.turnover_other:
            errs.append("[VBA 323] 44AD income must be at least 8% of E1(b+c) "
                        "(else tax audit -> use ITR-3/ITR-5)")
        return errs


@dataclass
class Profession44ADA:
    """BP sheet E3/E4 - profession under 44ADA."""
    gross_receipts: float = 0
    income: float = 0                 # >= 50% of gross receipts

    def validation_errors(self) -> list:
        errs = []
        if self.gross_receipts > CAP_44ADA_GROSS:
            errs.append("[VBA mdNOBBP] Gross receipts > Rs 75 Lakh require tax "
                        "audit u/s 44AB - use regular ITR-3/5")
        if self.income and self.income < RATE_44ADA * self.gross_receipts:
            errs.append("[44ADA] Presumptive income must be at least 50% of "
                        "gross receipts")
        if self.income > 0 and not self.gross_receipts:
            errs.append("[VBA 343] Invalid presumptive income under 44ADA, as "
                        "gross turnover is zero")
        return errs


@dataclass
class Vehicle44AE:
    """44AE sheet - goods carriages."""
    heavy: bool = False          # heavy goods vehicle
    units_or_tons: float = 0     # no. of vehicles (or MT for heavy)
    months: int = 12

    def min_income(self) -> float:
        """[VBA 11244/11252] Rs 1,000/MT/month heavy; Rs 7,500/month otherwise."""
        if self.heavy:
            return 1000 * self.units_or_tons * self.months
        return 7500 * self.units_or_tons * self.months


@dataclass
class Business44AE:
    vehicles: List[Vehicle44AE] = field(default_factory=list)
    declared_income: float = 0

    @property
    def income(self) -> float:
        return self.declared_income

    def min_income(self) -> float:
        return sum(v.min_income() for v in self.vehicles)

    def validation_errors(self) -> list:
        if self.declared_income and self.declared_income < self.min_income():
            return ["[VBA 11244/11252] 44AE income cannot be lower than "
                    "Rs 1,000/MT/month (heavy) or Rs 7,500/month (other)"]
        return []


# ---------------------------------------------------------------------------
# The return
# ---------------------------------------------------------------------------
@dataclass
class ITR4Return:
    entity: str = ENTITY_INDIVIDUAL
    dob: Optional[date] = None
    regime: int = C.REGIME_NEW
    salary: float = 0
    house_properties: List[HouseProperty] = field(default_factory=list)
    other_sources: float = 0
    b44ad: Business44AD = field(default_factory=Business44AD)
    p44ada: Profession44ADA = field(default_factory=Profession44ADA)
    b44ae: Business44AE = field(default_factory=Business44AE)
    chapter_via_total: float = 0        # computed Chapter VI-A (same caps as ITR-1)
    relief_89: float = 0
    tds: float = 0
    tcs: float = 0
    advance_tax: float = 0
    instalments: List[float] = field(default_factory=list)
    filing_sec_code: int = 11
    due_date: date = C.ITR_FILING_DUE_DATE
    verification_date: Optional[date] = None

    @property
    def age(self) -> int:
        return calculate_age(self.dob) if self.dob else 0

    def presumptive_income(self) -> float:
        """[BP!I72] MAX(0, E2 + E4 + E7)."""
        return max(0.0, self.b44ad.income + self.p44ada.income + self.b44ae.income)

    def gti(self) -> float:
        hp = hp_for_gti(self.house_properties, self.regime)
        return max(0.0, self.salary) + hp + max(0.0, self.other_sources) \
            + self.presumptive_income()

    def total_income(self) -> int:
        """Rounded to Rs 10 like the utility."""
        return via.round_total_income(self.gti() - self.chapter_via_total)

    # ---- tax on total income [TaxCalc D26] ----------------------------
    def tax_old_regime(self, ti: float) -> int:
        """TaxCalc columns D (individual GrpA/B/C), F (HUF), H (Firm)."""
        if self.entity == ENTITY_FIRM:
            return excel_round(ti * 0.30)                       # [H14]
        if self.entity == ENTITY_HUF:                           # [F10-F13]
            if ti <= 250000:
                return 0
            if ti <= 500000:
                return excel_round((ti - 250000) * 0.05)
            if ti <= 1000000:
                return excel_round((ti - 500000) * 0.20 + 12500)
            return excel_round((ti - 1000000) * 0.30 + 112500)
        age = self.age
        if age > 79:
            if ti <= 500000:
                return 0
            if ti <= 1000000:
                return excel_round((ti - 500000) * 0.20)
            return excel_round((ti - 1000000) * 0.30 + 100000)
        if age > 59:
            if ti <= 300000:
                return 0
            if ti <= 500000:
                return excel_round((ti - 300000) * 0.05)
            if ti <= 1000000:
                return excel_round((ti - 500000) * 0.20 + 10000)
            return excel_round((ti - 1000000) * 0.30 + 110000)
        if ti <= 250000:
            return 0
        if ti <= 500000:
            return excel_round((ti - 250000) * 0.05)
        if ti <= 1000000:
            return excel_round((ti - 500000) * 0.20 + 12500)
        return excel_round((ti - 1000000) * 0.30 + 112500)

    def d1_total_tax_payable(self) -> int:
        """[TaxCalc!D26] IF(bacValue=1, new-regime tax, old-regime sum)."""
        ti = self.total_income()
        if self.entity == ENTITY_FIRM:
            return excel_round(ti * 0.30)
        if self.regime == C.REGIME_NEW:
            return tax_new_regime(ti)
        return self.tax_old_regime(ti)

    # ---- surcharge + marginal relief [TaxCalc rows 29-43] -------------
    def surcharge(self) -> int:
        ti = self.total_income()
        tax = self.d1_total_tax_payable()
        if self.entity == ENTITY_FIRM:
            if ti > 10000000:
                return self._surcharge_with_relief(tax, ti, 0.12, 10000000)
            return 0
        # individual / HUF
        if 5000000 < ti <= 10000000:
            return self._surcharge_with_relief(tax, ti, 0.10, 5000000)
        if ti > 10000000:
            return self._surcharge_with_relief(tax, ti, 0.15, 10000000)
        return 0

    def _surcharge_with_relief(self, tax: int, ti: float, rate: float,
                               threshold: float) -> int:
        """[TaxCalc!B32-B36/C32-C36] surcharge with marginal relief:
        relief = max(0, (tax+surcharge - tax_at_threshold) - (ti - threshold));
        final surcharge = tax*rate - relief."""
        tax_at_threshold = self.tax_old_regime(threshold) if self.regime == C.REGIME_OLD \
            else tax_new_regime(threshold)
        if self.entity == ENTITY_FIRM:
            tax_at_threshold = excel_round(threshold * 0.30)
        gross = tax + tax * rate
        relief = max(0.0, (gross - tax_at_threshold) - (ti - threshold))
        return excel_round(max(0.0, tax * rate - relief))

    def cess(self) -> int:
        """[TaxCalc!N19] 4% on (tax + surcharge - rebate) [N18 = post-rebate]."""
        base = self.d1_total_tax_payable() + self.surcharge() - self.rebate_87a()
        return excel_round(max(0, base) * 0.04)

    def rebate_87a(self) -> int:
        """Individuals only (not HUF/Firm) [TaxCalc!N17 guard Status<>F/H]."""
        if self.entity != ENTITY_INDIVIDUAL:
            return 0
        from .tax_engine import rebate_87a
        return rebate_87a(self.d1_total_tax_payable(), self.total_income(),
                          self.total_income(), self.regime)

    # ---- net liability / interest --------------------------------------
    def gross_tax_liability(self) -> int:
        return excel_round(self.d1_total_tax_payable() + self.surcharge()
                           - self.rebate_87a() + self.cess())

    def net_tax_liability(self) -> int:
        return excel_round(max(0, self.gross_tax_liability() - self.relief_89))

    def interest_234b(self) -> int:
        vdate = self.verification_date or date.today()
        age = self.age if self.entity == ENTITY_INDIVIDUAL else 0
        return interest_234b(self.net_tax_liability(), self.advance_tax,
                             self.tds, self.tcs, vdate, age)

    def interest_234a(self) -> int:
        vdate = self.verification_date or self.due_date
        ntl = self.net_tax_liability() - (self.tds + self.tcs + self.advance_tax)
        return interest_234a(max(0.0, ntl), self.due_date, vdate)

    def fee_234f(self) -> int:
        return fee_234f(self.total_income(), self.filing_sec_code not in (11,))

    def refund_or_payable(self) -> float:
        total_paid = self.tds + self.tcs + self.advance_tax
        liability = (self.net_tax_liability() + self.interest_234a()
                     + self.interest_234b() + self.fee_234f())
        return total_paid - liability

    # ------------------------------------------------------------------
    def validation_errors(self) -> list:
        errs = self.b44ad.validation_errors() + self.p44ada.validation_errors() \
            + self.b44ae.validation_errors()
        if self.total_income() > ITR4_TI_CAP:
            errs.append("[VBA 15467] ITR-4 total income must not exceed Rs 50 Lakh "
                        "- please file ITR-3 or 5")
        if self.entity not in (ENTITY_INDIVIDUAL, ENTITY_HUF, ENTITY_FIRM):
            errs.append("[Eligibility] ITR-4 is for Individuals, HUFs and Firms "
                        "(other than LLP) only")
        return errs

    def compute_summary(self) -> dict:
        return {
            "PresumptiveIncome": excel_round(self.presumptive_income()),
            "GrossTotIncome": excel_round(self.gti()),
            "TotalIncome": self.total_income(),
            "TotalTaxPayable": self.d1_total_tax_payable(),
            "Surcharge": self.surcharge(),
            "Rebate87A": self.rebate_87a(),
            "EducationCess": self.cess(),
            "GrossTaxLiability": self.gross_tax_liability(),
            "NetTaxLiability": self.net_tax_liability(),
            "IntrstPay234A": self.interest_234a(),
            "IntrstPay234B": self.interest_234b(),
            "FeeIncUS234F": self.fee_234f(),
            "RefundOrPayable": self.refund_or_payable(),
        }
