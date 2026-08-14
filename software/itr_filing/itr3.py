"""ITR-3 AY 2026-27 return model - ITR-2 pipeline plus business/profession
income (P&L businesses, depreciation, speculation, specified business,
presumptive taxation)."""
from dataclasses import dataclass, field
from datetime import date
from typing import List, Dict

from . import constants as C
from .tax_engine import excel_round
from .itr1 import PersonalInfo, BankAccount, TaxesPaid
from .itr2 import Employer
from .house_property import HouseProperty, total_hp_income
from .capital_gains import CapitalAsset, compute_cg_summary, ALL_CG_BANDS
from .business import (BusinessPL, SpeculationBusiness, SpecifiedBusiness,
                       BusinessIncomeSummary, compute_business_income)
from .itr4 import Business44AD, Profession44ADA, Business44AE
from .setoff import cyla
from .carry_forward import bfla, CFLTracker
from . import chapter_via as via
from .ti_tti import TITTIInput, compute_ti_tti
from .interest import interest_234a, interest_234b, fee_234f


@dataclass
class ITR3Return:
    personal: PersonalInfo = field(default_factory=PersonalInfo)
    regime: int = C.REGIME_NEW
    residential_status: str = "RES"
    filing_sec_code: int = 11
    due_date: date = C.ITR_FILING_DUE_DATE
    verification_date: date = None

    # salary / HP / OS heads (same as ITR-2)
    employer: Employer = field(default_factory=Employer)
    gross_salary: float = 0
    exempt_allowances: float = 0
    deduction_16ii: float = 0
    deduction_16iii: float = 0
    house_properties: List[HouseProperty] = field(default_factory=list)
    other_sources_normal: float = 0
    other_sources_racehorse: float = 0
    cg_assets: List[CapitalAsset] = field(default_factory=list)

    # business & profession
    businesses: List[BusinessPL] = field(default_factory=list)
    speculation: List[SpeculationBusiness] = field(default_factory=list)
    specified: List[SpecifiedBusiness] = field(default_factory=list)
    b44ad: Business44AD = None
    p44ada: Profession44ADA = None
    b44ae: Business44AE = None

    # losses & deductions
    cfl: CFLTracker = field(default_factory=CFLTracker)
    depreciation_bf: float = 0
    via_total: float = 0

    # taxes paid / credits
    taxes_paid: TaxesPaid = field(default_factory=TaxesPaid)
    bank: BankAccount = field(default_factory=BankAccount)
    relief_89: float = 0
    relief_90: float = 0
    relief_91: float = 0
    advance_tax_instalments: List[float] = field(default_factory=list)

    @property
    def age(self) -> int:
        from .tax_engine import calculate_age
        return calculate_age(self.personal.date_of_birth) if self.personal.date_of_birth else 40

    def std_deduction_16ia(self) -> int:
        cap = C.STD_DEDUCTION_NEW if self.regime == C.REGIME_NEW else C.STD_DEDUCTION_OLD
        return excel_round(min(max(0.0, self.gross_salary - self.exempt_allowances), cap))

    def income_from_salaries(self) -> int:
        return excel_round(max(0.0, self.gross_salary - self.exempt_allowances
                               - self.std_deduction_16ia()
                               - self.deduction_16ii - self.deduction_16iii))

    def business_summary(self) -> BusinessIncomeSummary:
        return compute_business_income(self.businesses, self.speculation,
                                       self.specified, self.b44ad, self.p44ada,
                                       self.b44ae)

    def compute(self) -> dict:
        cg = compute_cg_summary(self.cg_assets)
        sal = self.income_from_salaries()
        hp = total_hp_income(self.house_properties)
        os_n = excel_round(max(0.0, self.other_sources_normal))
        bp = self.business_summary()

        incomes = {"SAL": sal, "HP": max(0, hp), "OS": os_n,
                   "RACEHORSE": excel_round(max(0.0, self.other_sources_racehorse)),
                   "BP": max(0, bp.normal + bp.presumptive_44ad
                             + bp.presumptive_44ada + bp.presumptive_44ae),
                   "SPEC": max(0, bp.speculation),
                   "SPECIFIED": max(0, bp.specified)}
        for b in ALL_CG_BANDS:
            incomes[b] = max(0.0, cg.get(b, 0))

        cyla_res = cyla(incomes, hp_loss=-min(0, hp), bp_loss=-min(0, bp.normal),
                        os_loss=0, regime=self.regime)

        bf = self.cfl.available("2026-27")
        bfla_res = bfla(cyla_res.income_after_setoff, bf, self.depreciation_bf)

        a = bfla_res.income_after_setoff
        cg_special_111a_112a = excel_round(a.get("STCG20", 0) + a.get("LTCG125", 0)
                                           + a.get("LTCG10", 0))
        cg_special_other = excel_round(a.get("STCG15", 0) + a.get("STCG30", 0)
                                       + a.get("LTCG20", 0) + a.get("STCG_DTAA", 0)
                                       + a.get("LTCG_DTAA", 0))
        normal_ti = excel_round(a.get("SAL", 0) + a.get("HP", 0) + a.get("OS", 0)
                                + a.get("BP", 0) + a.get("SPECIFIED", 0)
                                + a.get("STCG_RATE", 0) + a.get("RACEHORSE", 0))
        speculation_ti = excel_round(a.get("SPEC", 0))
        gti = normal_ti + speculation_ti + cg_special_111a_112a + cg_special_other
        via_ded = min(float(self.via_total or 0), gti)
        total_income = via.round_total_income(gti - via_ded)

        tti = compute_ti_tti(TITTIInput(
            total_income_normal=normal_ti,
            cg_net_by_band={b: a.get(b, 0) for b in ALL_CG_BANDS},
            ti_incl_112a=total_income,
            regime=self.regime, age=self.age,
            relief_89=self.relief_89,
            relief_90_91=self.relief_90 + self.relief_91,
            amt_payable=0,
        ))

        tp = self.taxes_paid
        vdate = self.verification_date or date.today()
        ntl = tti["NetTaxLiability"]
        i234b = interest_234b(ntl, tp.advance_tax, tp.total_tds, tp.total_tcs,
                              vdate, self.age)
        i234a = interest_234a(max(0.0, ntl - tp.total_tds - tp.total_tcs
                                  - tp.total_it), self.due_date, vdate)
        f234f = fee_234f(total_income, self.filing_sec_code not in (11,))
        total_liability = ntl + i234a + i234b + f234f
        paid = tp.total_tds + tp.total_tcs + tp.total_it
        refund = max(0.0, paid - total_liability)

        return {
            "bp": bp, "cg_summary": cg, "cyla": cyla_res, "bfla": bfla_res,
            "Salaries": sal, "IncomeFromHP": a.get("HP", 0),
            "BusinessIncome": excel_round(a.get("BP", 0)),
            "SpeculationIncome": speculation_ti,
            "IncFromOS": excel_round(a.get("OS", 0) + a.get("RACEHORSE", 0)),
            "CapGainTotal": excel_round(cg_special_111a_112a + cg_special_other),
            "IncChargeTaxSplRate111A112": cg_special_111a_112a,
            "IncChargeableTaxSplRates": cg_special_other,
            "GrossTotalIncome": gti,
            "VIADeductions": excel_round(via_ded),
            "TotalIncome": total_income,
            "tti": tti,
            "IntrstPay234A": i234a, "IntrstPay234B": i234b, "Fee234F": f234f,
            "TotalLiability": excel_round(total_liability),
            "TaxesPaid": excel_round(paid),
            "BalTaxPayable": excel_round(max(0.0, total_liability - paid)),
            "RefundDue": excel_round(refund),
        }
