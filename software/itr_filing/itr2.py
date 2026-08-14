"""ITR-2 AY 2026-27 return model - ties together all engine modules:
Salary/HP/OS heads, CG waterfall, CYLA, BFLA, CFL, Chapter VI-A, TI/TTI
computation, interest, refund."""
from dataclasses import dataclass, field
from datetime import date
from typing import List, Dict

from . import constants as C
from .tax_engine import excel_round, calculate_age
from .itr1 import PersonalInfo, BankAccount, TaxesPaid
from .house_property import HouseProperty, total_hp_income, hp_for_gti
from .capital_gains import CapitalAsset, compute_cg_summary, ALL_CG_BANDS
from .setoff import cyla, CYLAResult
from .carry_forward import bfla, BFLAResult, CFLTracker, LossEntry, ay_start
from . import chapter_via as via
from .ti_tti import TITTIInput, compute_ti_tti
from .interest import interest_234a, interest_234b, fee_234f


@dataclass
class Employer:
    name: str = ""
    tan: str = ""
    nature_of_employment: str = "Salaried"


@dataclass
class ITR2Return:
    personal: PersonalInfo = field(default_factory=PersonalInfo)
    regime: int = C.REGIME_NEW
    residential_status: str = "RES"
    filing_sec_code: int = 11
    due_date: date = C.ITR_FILING_DUE_DATE
    verification_date: date = None

    # income heads
    employer: Employer = field(default_factory=Employer)
    gross_salary: float = 0
    exempt_allowances: float = 0
    deduction_16ia: float = None      # auto: 75k/50k by regime
    deduction_16ii: float = 0
    deduction_16iii: float = 0
    house_properties: List[HouseProperty] = field(default_factory=list)
    other_sources_normal: float = 0
    other_sources_racehorse: float = 0
    cg_assets: List[CapitalAsset] = field(default_factory=list)

    # losses
    cfl: CFLTracker = field(default_factory=CFLTracker)
    depreciation_bf: float = 0

    # chapter VI-A
    via_user_inputs: dict = field(default_factory=dict)
    via_total: float = 0               # computed total (or override)

    # taxes paid
    taxes_paid: TaxesPaid = field(default_factory=TaxesPaid)
    bank: BankAccount = field(default_factory=BankAccount)
    relief_89: float = 0
    relief_90: float = 0
    relief_91: float = 0

    # ------------------------------------------------------------------
    @property
    def age(self) -> int:
        return calculate_age(self.personal.date_of_birth) if self.personal.date_of_birth else 40

    def std_deduction_16ia(self) -> int:
        cap = C.STD_DEDUCTION_NEW if self.regime == C.REGIME_NEW else C.STD_DEDUCTION_OLD
        return excel_round(min(self.net_salary_gross(), cap))

    def net_salary_gross(self) -> float:
        return max(0.0, self.gross_salary - self.exempt_allowances)

    def income_from_salaries(self) -> int:
        return excel_round(max(0.0, self.net_salary_gross()
                               - self.std_deduction_16ia()
                               - self.deduction_16ii - self.deduction_16iii))

    def hp_income(self) -> int:
        return total_hp_income(self.house_properties)

    def cg_summary(self) -> Dict[str, float]:
        return compute_cg_summary(self.cg_assets)

    def compute(self) -> dict:
        cg = self.cg_summary()
        hp = self.hp_income()
        sal = self.income_from_salaries()
        os_n = excel_round(max(0.0, self.other_sources_normal))

        # ---- CYLA (current-year losses: HP / OS normal) -----------------
        incomes = {"SAL": sal, "HP": max(0, hp), "OS": os_n,
                   "RACEHORSE": excel_round(max(0.0, self.other_sources_racehorse))}
        for b in ALL_CG_BANDS:
            incomes[b] = max(0.0, cg.get(b, 0))
        hp_loss = -min(0, hp)
        cyla_res = cyla(incomes, hp_loss=hp_loss, bp_loss=0, os_loss=0,
                        regime=self.regime)

        # ---- BFLA --------------------------------------------------------
        bf = self.cfl.available("2026-27")
        bfla_res = bfla(cyla_res.income_after_setoff, bf, self.depreciation_bf)

        # ---- TI aggregation ----------------------------------------------
        a = bfla_res.income_after_setoff
        cg_special_111a_112a = excel_round(a.get("STCG20", 0) + a.get("LTCG125", 0)
                                           + a.get("LTCG10", 0))
        cg_special_other = excel_round(a.get("STCG15", 0) + a.get("STCG30", 0)
                                       + a.get("LTCG20", 0) + a.get("STCG_DTAA", 0)
                                       + a.get("LTCG_DTAA", 0))
        normal_ti = excel_round(a.get("SAL", 0) + a.get("HP", 0) + a.get("OS", 0)
                                + a.get("STCG_RATE", 0) + a.get("RACEHORSE", 0))
        gti = normal_ti + cg_special_111a_112a + cg_special_other
        via_ded = min(float(self.via_total or 0), gti)
        total_income = via.round_total_income(gti - via_ded)
        ti_incl_112a = via.round_total_income(gti - via_ded)   # ITR-2: 112A in TI

        # ---- TI/TTI computation ------------------------------------------
        tti = compute_ti_tti(TITTIInput(
            total_income_normal=normal_ti,
            cg_net_by_band={b: a.get(b, 0) for b in ALL_CG_BANDS},
            ti_incl_112a=ti_incl_112a,
            regime=self.regime, age=self.age,
            relief_89=self.relief_89,
            relief_90_91=self.relief_90 + self.relief_91,
            amt_payable=0,
        ))

        # ---- interest / fees / refund -------------------------------------
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
            "cg_summary": cg, "cyla": cyla_res, "bfla": bfla_res,
            "Salaries": sal, "IncomeFromHP": a.get("HP", 0),
            "IncFromOS": excel_round(a.get("OS", 0) + a.get("RACEHORSE", 0)),
            "CapGainTotal": excel_round(cg_special_111a_112a + cg_special_other),
            "IncChargeTaxSplRate111A112": cg_special_111a_112a,
            "IncChargeableTaxSplRates": cg_special_other,
            "GrossTotalIncome": gti,
            "VIADeductions": excel_round(via_ded),
            "TotalIncome": total_income,
            "tti": tti,
            "IntrstPay234A": i234a, "IntrstPay234B": i234b,
            "Fee234F": f234f,
            "TotalLiability": excel_round(total_liability),
            "TaxesPaid": excel_round(paid),
            "BalTaxPayable": excel_round(max(0.0, total_liability - paid)),
            "RefundDue": excel_round(refund),
        }
