"""ITR-1 (SAHAJ) AY 2026-27 - complete return model & computation pipeline.

Field-for-field port of ITR1_AY_26-27_V1.2.xlsm (visible 'Income Details',
'HP', 'TDS', 'TCS', 'Taxes Paid and Verification' sheets + hidden schedules +
VBA engine). Every computed property names the Excel cell/VBA procedure it
replicates, so the implementation can be audited against the extraction
registry (registry/ITR1.*.csv).
"""
from dataclasses import dataclass, field
from datetime import date
from typing import List, Optional

from . import constants as C
from .tax_engine import (tax_on_total_income, rebate_87a, education_cess,
                         ltcg_112a_not_chargeable, calculate_age, excel_round)
from .house_property import HouseProperty, total_hp_income, hp_for_gti
from . import chapter_via as via
from .interest import interest_234a, interest_234b, fee_234f

# Filing-section codes [VBA ReturnSecList strings]
RETURN_FILE_SEC = {
    11: "139(1)-On or before due date",
    12: "139(4)-Belated",
    13: "142(1)",
    14: "148",
    15: "153A",
    16: "153C",
    17: "139(5)-Revised",
    18: "139(9)",
    20: "119(2)(b)-Condonation of delay",
    21: "139(8A)-Updated return",
}


# ---------------------------------------------------------------------------
# Part A - General Information  [Income Details rows 6..53]
# ---------------------------------------------------------------------------
@dataclass
class Address:
    flat_door_block: str = ""
    name_of_premises: str = ""
    road_street_post_office: str = ""
    area_locality: str = ""
    city: str = ""
    state: str = "(Select)"
    country: str = C.COUNTRY_DEFAULT
    pin_code: str = ""


@dataclass
class PersonalInfo:
    first_name: str = ""                       # E7
    middle_name: str = ""                      # O7
    last_name: str = ""                        # Y7
    pan: str = ""                              # Part A PAN
    father_name: str = ""                      # Verification declaration
    aadhaar_number: str = ""                   # AN8 (12 digits)
    aadhaar_enrolment_id: str = ""             # AN9 (28 chars)
    date_of_birth: Optional[date] = None       # E10
    nature_of_employment: str = "(Select)"     # Z11 -> EmpCatList
    primary_email: str = ""                    # E13
    secondary_email: str = ""                  # K13
    primary_mobile: str = ""                   # Y13
    secondary_mobile: str = ""                 # AB13
    primary_address: Address = field(default_factory=Address)
    secondary_address_same_as_primary: str = "(Select)"   # AN23
    secondary_address: Address = field(default_factory=Address)

    @property
    def age(self) -> int:
        return calculate_age(self.date_of_birth)


@dataclass
class RepresentativeAssessee:                    # rows 42-45
    applicable: str = "(Select)"               # AD42 (Select/Yes/No)
    name: str = ""
    email: str = ""
    contact_no: str = ""


@dataclass
class FilingStatus:                              # rows 31..53
    # row 31 - option u/s 115BAC -> BacValue  [DV AN31]
    opt_out_new_regime: str = "(Select)"       # "Yes" = opt out -> old regime
    seventh_provisio_139: str = "(Select)"     # row 33 [AN33]
    amount_seventh_provisio_ii: int = 0        # row 35 travel/foreign > 2L
    flag_electricity_1lk: str = "(Select)"     # row 36
    amount_seventh_provisio_iii: int = 0
    clause_iv_seventh_provisio: str = "(Select)"   # row 37
    clause_iv_details: List[tuple] = field(default_factory=list)  # rows 38-41 (flag, amount)
    sales_turnover_gt_60cr: str = "(Select)"   # row 38 flag context
    filing_sec_code: int = 11                  # AC47 -> ReturnSecList
    receipt_number: str = ""                   # P50 (revised/defective)
    original_return_filed_date: Optional[date] = None   # AD50
    notice_unique_number: str = ""             # P52
    notice_date: Optional[date] = None         # AD52
    due_date: date = C.ITR_FILING_DUE_DATE     # P53 -> 31/07/2026
    representative: RepresentativeAssessee = field(default_factory=RepresentativeAssessee)

    @property
    def regime(self) -> int:
        """BacValue semantics: 'Yes' (opt out of 115BAC) -> OLD regime (2)."""
        v = str(self.opt_out_new_regime).replace("(Select)", "").strip().lower()
        return C.REGIME_OLD if v.startswith("yes") else C.REGIME_NEW

    @property
    def filed_after_due_date(self) -> bool:
        return self.filing_sec_code not in (11,)


# ---------------------------------------------------------------------------
# B1 - Salary  [Income Details rows 54..76]
# ---------------------------------------------------------------------------
@dataclass
class RetirementBenefitAccount:                # rows 58-63 / 92-97
    us_amount: float = 0                       # IncomeNotified89A_AmountUS
    uk_amount: float = 0
    canada_amount: float = 0
    other_country_amount: float = 0            # IncomeNotifiedOther89A

    @property
    def notified_total(self) -> float:         # [AO58] US+UK+Canada
        return self.us_amount + self.uk_amount + self.canada_amount


@dataclass
class Salary:
    salary_17_1: float = 0                     # IncD.Allowances (r55)
    perquisites_17_2: float = 0                # IncD.Perquisites
    profits_in_lieu_17_3: float = 0            # IncD.Profits
    retirement_benefit: RetirementBenefitAccount = field(default_factory=RetirementBenefitAccount)
    exempt_allowances: List[tuple] = field(default_factory=list)  # (nature, amount) rows 66-68
    hra_exempt_u10_13a: float = 0              # Sheet1.HRA from Schedule EA 10(13A)
    # Schedule EA 10(13A) inputs -> auto-compute HRA exemption when set
    basic_salary: float = 0
    dearness_allowance: float = 0
    hra_received: float = 0
    rent_paid: float = 0
    metro: bool = False
    relief_89a_on_allowances: float = 0        # iia row 70 (Increliefus89A)
    entertainment_allowance_16_ii: float = 0   # r74
    professional_tax_16_iii: float = 0         # r75

    def hra_exemption(self, regime: int) -> int:
        """[Schedule EA 10(13A)!G12] auto-computed unless manually entered."""
        if self.hra_exempt_u10_13a:
            return int(self.hra_exempt_u10_13a)
        if self.hra_received or self.rent_paid:
            from .schedules import hra_exemption_10_13a
            return hra_exemption_10_13a(self.hra_received, self.rent_paid,
                                        self.basic_salary, self.dearness_allowance,
                                        self.metro, regime)
        return 0

    @property
    def gross_salary(self) -> int:             # [AO54] MAX(0, 17(1)+17(2)+17(3))
        return excel_round(max(0.0, self.salary_17_1 + self.perquisites_17_2
                               + self.profits_in_lieu_17_3))

    def exempt_u10_total(self, regime: int = None) -> float:
        """[AO64] SUM(Others.Amount_1, HRA). Regime needed for auto-HRA."""
        hra = self.hra_exemption(regime) if regime is not None else self.hra_exempt_u10_13a
        return sum(a for _, a in self.exempt_allowances) + hra

    def net_salary_for(self, regime: int) -> int:
        """[AO71] MAX(0, gross - MIN(exempt, gross))"""
        return excel_round(max(0.0, self.gross_salary
                               - min(self.exempt_u10_total(regime), self.gross_salary)))

    @property
    def net_salary(self) -> int:
        return self.net_salary_for(C.REGIME_NEW)

    def deduction_16ia(self, regime: int) -> int:
        """[AO73] IF(BacValue=1,MIN(net,75000),IF(BacValue=2,MIN(net,50000),0))"""
        cap = C.STD_DEDUCTION_NEW if regime == C.REGIME_NEW else C.STD_DEDUCTION_OLD
        return excel_round(min(self.net_salary, cap))

    @property
    def deductions_16(self) -> int:
        return 0  # replaced by compute path needing regime - see Return

    def entertainment_allowance_allowed(self) -> int:
        """[VBA ~1866] least of actual, Rs 5,000, 1/5 of salary u/s 16(ii)."""
        from .schedules import cap_entertainment_allowance_16ii
        return cap_entertainment_allowance_16ii(self.entertainment_allowance_16_ii,
                                                self.salary_17_1)

    def income_from_salaries(self, regime: int) -> int:
        """[AO76] MAX(0, net_salary - deductions_u16)"""
        d16 = (self.deduction_16ia(regime) + self.entertainment_allowance_allowed()
               + self.professional_tax_16_iii)
        return excel_round(max(0.0, self.net_salary_for(regime) - d16))


# ---------------------------------------------------------------------------
# B3 - Other Sources  [rows 85..110]
# ---------------------------------------------------------------------------
@dataclass
class OtherSources:
    items: List[tuple] = field(default_factory=list)      # (nature, description, amount)
    dividend_by_quarter: List[float] = field(default_factory=lambda: [0.0] * 5)  # r104-108
    retirement_benefit_os: RetirementBenefitAccount = field(default_factory=RetirementBenefitAccount)
    deduction_57_iia: float = 0                # family pension u/s 57(iia) [r110]
    relief_89a: float = 0                      # OSIncreliefus89A [r109]

    @property
    def dividend(self) -> float:               # [IncD.Div] sum of i..v
        return sum(self.dividend_by_quarter)

    @property
    def other_income_grid_total(self) -> float:
        return sum(a for _, _, a in self.items)

    def income_from_os(self) -> int:
        """[AO85] MAX(0, Div + grid + notified89A(OS) + notifiedOther89A(OS)
        - deduction57 - relief89A)"""
        gross = (self.dividend + self.other_income_grid_total
                 + self.retirement_benefit_os.notified_total
                 + self.retirement_benefit_os.other_country_amount)
        return excel_round(max(0.0, gross - self.deduction_57_iia - self.relief_89a))

    @property
    def savings_interest(self) -> float:
        """'tta' helper - interest from savings accounts [Others.Amount_2_3]."""
        return sum(a for n, _, a in self.items if "saving" in n.lower())

    @property
    def senior_deposit_interest(self) -> float:
        """'ttb' helper - bank/post-office/co-op deposit interest [Amount_2_4/_5]."""
        keys = ("deposit", "post office", "co-operative", "cooperative")
        return sum(a for n, _, a in self.items if any(k in n.lower() for k in keys))


# ---------------------------------------------------------------------------
# C3a - LTCG 112A  [rows 172..175]
# ---------------------------------------------------------------------------
@dataclass
class LTCG112A:
    total_sale_consideration: float = 0        # r173
    total_cost_of_acquisition: float = 0       # r174

    @property
    def gain(self) -> float:
        return max(0.0, self.total_sale_consideration - self.total_cost_of_acquisition)

    @property
    def not_chargeable(self) -> int:           # [AO175]
        return ltcg_112a_not_chargeable(self.total_sale_consideration,
                                        self.total_cost_of_acquisition)


# ---------------------------------------------------------------------------
# Part C - Chapter VI-A inputs  [rows 113..163]
# ---------------------------------------------------------------------------
@dataclass
class ChapterVIA:
    # schedule objects (hidden sheets) - when provided they drive the amounts
    schedule_80d: object = None          # schedules.Schedule80D
    disability_80dd: int = 0             # 1 -> 75k, 2 -> 125k (hidden 80U-80DD sheet)
    severity_80u: int = 0                # 1 -> 75k, 2 -> 125k
    selection_80ddb: int = 0             # 1 -> 40k, 2 -> 1L [BM46]
    donations_80g: list = field(default_factory=list)   # schedules.Donation80G
    s80c: float = 0
    s80ccc: float = 0
    s80ccd1: float = 0            # _SE (state/central govt employee path)
    s80ccd1_employer_category_pct: float = 0.10   # BA126: 1->10% private, 2->14% govt
    s80ccd1b: float = 0
    s80ccd2: float = 0
    s80ccg: float = 0
    s80d_schedule_amount: float = 0      # computed by hidden 80D sheet
    s80dd_schedule_amount: float = 0     # hidden 80U-80DD sheet (75k/125k)
    s80ddb_schedule_amount: float = 0
    s80e: float = 0
    s80ee: float = 0
    s80eea: float = 0
    s80eeb: float = 0
    s80g_schedule_amount: float = 0      # hidden 80G sheet
    s80gg: float = 0
    s80gg_ack_form10ba: str = ""
    s80gga_eligible_donation: float = 0
    s80ggc_eligible_donation: float = 0
    s80qqb: float = 0
    s80rrb: float = 0
    s80tta: float = 0
    s80ttb: float = 0
    s80u_schedule_amount: float = 0
    s80cch_agnipath: float = 0
    user_total: float = 0               # TotalChapVIADeductions_Input (user-entered)


# ---------------------------------------------------------------------------
# Taxes paid schedules (TDS / TCS / IT sheets)
# ---------------------------------------------------------------------------
@dataclass
class TDSonSalary:            # TDS sheet section 1
    employer_name: str = ""
    tan: str = ""
    employer_pan: str = ""
    tds_amount: float = 0


@dataclass
class TDSonOtherThanSalary:   # TDS sheet section 2/3
    deductor_name: str = ""
    tan: str = ""
    deductor_pan: str = ""
    section: str = "(Select)"
    tds_amount: float = 0


@dataclass
class TCSEntry:
    collector_name: str = ""
    tan: str = ""
    pan: str = ""
    tcs_amount: float = 0


@dataclass
class ITChallan:              # Taxes Paid / Schedule IT
    bsr_code: str = ""
    date_of_deposit: Optional[date] = None
    challan_serial_no: str = ""
    amount: float = 0
    mode: str = "(Select)"    # Advance Tax / Self Assessment Tax


@dataclass
class TaxesPaid:
    tds_salary: List[TDSonSalary] = field(default_factory=list)
    tds_other: List[TDSonOtherThanSalary] = field(default_factory=list)
    tcs: List[TCSEntry] = field(default_factory=list)
    it_challans: List[ITChallan] = field(default_factory=list)
    # advance-tax instalments paid by [15-Jun, 15-Sep, 15-Dec, 15-Mar, later]
    instalments: List[float] = field(default_factory=list)

    @property
    def total_tds(self) -> float:
        return sum(t.tds_amount for t in self.tds_salary) + sum(t.tds_amount for t in self.tds_other)

    @property
    def total_tcs(self) -> float:
        return sum(t.tcs_amount for t in self.tcs)

    @property
    def advance_tax(self) -> float:
        return sum(c.amount for c in self.it_challans if "advance" in c.mode.lower())

    @property
    def self_assessment_tax(self) -> float:
        return sum(c.amount for c in self.it_challans if "self" in c.mode.lower())

    @property
    def total_it(self) -> float:
        return sum(c.amount for c in self.it_challans)


# ---------------------------------------------------------------------------
# Verification & refund
# ---------------------------------------------------------------------------
@dataclass
class Verification:
    verification_date: Optional[date] = None   # Ver.Date (Ver.Date drives 234 periods)
    capacity: str = "Self"


@dataclass
class BankAccount:
    ifsc: str = ""
    account_no: str = ""
    bank_name: str = ""
    is_for_refund: str = "No"


# ---------------------------------------------------------------------------
# The complete return
# ---------------------------------------------------------------------------
@dataclass
class ITR1Return:
    personal: PersonalInfo = field(default_factory=PersonalInfo)
    filing: FilingStatus = field(default_factory=FilingStatus)
    salary: Salary = field(default_factory=Salary)
    house_properties: List[HouseProperty] = field(default_factory=list)
    other_sources: OtherSources = field(default_factory=OtherSources)
    ltcg_112a: LTCG112A = field(default_factory=LTCG112A)
    chapter_via: ChapterVIA = field(default_factory=ChapterVIA)
    relief_89: float = 0                 # D6 [r182 input]
    relief_89a: float = 0                # 13a computed [AO183]
    taxes_paid: TaxesPaid = field(default_factory=TaxesPaid)
    bank: BankAccount = field(default_factory=BankAccount)
    verification: Verification = field(default_factory=Verification)

    # ------------------------------------------------------------------
    # Pipeline - mirrors Income Details!AO54..AO191 exactly
    # ------------------------------------------------------------------
    @property
    def regime(self) -> int:
        return self.filing.regime

    def income_from_salaries(self) -> int:
        return self.salary.income_from_salaries(self.regime)

    def hp_income(self) -> int:
        return total_hp_income(self.house_properties)

    def gti(self) -> int:
        """B4 [AO111] SUM(Salaries, HP[ floored if new regime ], OS)."""
        return excel_round(self.income_from_salaries()
                           + hp_for_gti(self.house_properties, self.regime)
                           + self.other_sources.income_from_os())

    def gti_incl_112a(self) -> int:
        """B4 incl. 112A [AO112]."""
        return excel_round(self.gti() + self.ltcg_112a.gain)

    def chapter_via_components(self) -> dict:
        """Computed deduction per section [AN115..AN161]."""
        ti_probe = max(0.0, self.gti())   # TOTAL_INCOME helper for caps
        cv = self.chapter_via
        net_sal_less_perq = self.salary.net_salary_for(self.regime) - self.salary.perquisites_17_2
        hra = self.salary.hra_exemption(self.regime)
        comps = {
            "80C": via.cap_80c(cv.s80c, ti_probe),
            "80CCC": via.cap_80ccc(cv.s80ccc, self.gti(),
                                    via.cap_80c(cv.s80c, ti_probe), ti_probe),
            "80CCD1": via.cap_80ccd1(cv.s80ccd1, net_sal_less_perq,
                                      cv.s80ccd1_employer_category_pct, self.regime, ti_probe),
            "80CCD1B": via.cap_80ccd1b(cv.s80ccd1b, ti_probe),
            "80CCD2": via.cap_80ccd2(cv.s80ccd2, net_sal_less_perq,
                                      cv.s80ccd1_employer_category_pct, self.regime, ti_probe),
            "80CCG": via.cap_80ccg(cv.s80ccg, ti_probe),
            "80D": via.cap_80d(self._amount_80d(), self.gti()),
            "80DD": via.cap_80dd(self._amount_80dd(), self.regime),
            "80DDB": via.cap_80ddb(self._amount_80ddb()),
            "80E": via.cap_80e(cv.s80e, ti_probe),
            "80EE": via.cap_80ee(cv.s80ee, ti_probe),
            "80EEA": via.cap_80eea(cv.s80eea, ti_probe),
            "80EEB": via.cap_80eeb(cv.s80eeb, ti_probe),
            "80G": 0,  # computed below to avoid recursion (needs other comps)
            "80GG": via.cap_80gg(cv.s80gg, hra, ti_probe, 0),
            "80GGA": via.cap_80gga(cv.s80gga_eligible_donation, self.regime),
            "80GGC": via.cap_80ggc(cv.s80ggc_eligible_donation, self.regime),
            "80QQB": excel_round(min(float(cv.s80qqb or 0), 300000)),
            "80RRB": excel_round(min(float(cv.s80rrb or 0), 300000)),
            "80TTA": via.cap_80tta(cv.s80tta, ti_probe, self.other_sources.savings_interest),
            "80TTB": via.cap_80ttb(cv.s80ttb, ti_probe, self.other_sources.senior_deposit_interest),
            "80U": via.cap_80u(self._amount_80u(), self.regime),
            "80CCH": via.cap_80cch(cv.s80cch_agnipath, self.salary.exempt_u10_total(self.regime)),
        }
        comps["80G"] = via.cap_80g(self._amount_80g(comps))
        return comps

    # --- schedule amount resolution --------------------------------
    def _amount_80d(self) -> float:
        cv = self.chapter_via
        if cv.schedule_80d is not None:
            return cv.schedule_80d.eligible_amount(max(0.0, self.gti()))
        return cv.s80d_schedule_amount

    def _amount_80dd(self) -> float:
        cv = self.chapter_via
        if cv.disability_80dd:
            from .schedules import amount_80dd
            return amount_80dd(cv.disability_80dd)
        return cv.s80dd_schedule_amount

    def _amount_80u(self) -> float:
        cv = self.chapter_via
        if cv.severity_80u:
            from .schedules import amount_80u
            return amount_80u(cv.severity_80u)
        return cv.s80u_schedule_amount

    def _amount_80ddb(self) -> float:
        """[BM46] selection sets the cap (40k / 1L); amount is user-entered."""
        cv = self.chapter_via
        if cv.selection_80ddb:
            from .schedules import amount_80ddb
            return min(amount_80ddb(cv.selection_80ddb), float(cv.s80ddb_schedule_amount or 0))
        return cv.s80ddb_schedule_amount

    def _amount_80g(self, comps: dict = None) -> float:
        """80G eligible amount. The qualifying limit uses adjusted TI that
        EXCLUDES 80G itself [80G!C33]; TI here is computed from all OTHER
        components in one pass (Excel resolves the circular ref by iteration;
        one pass is the fixed point)."""
        cv = self.chapter_via
        if cv.donations_80g:
            from .schedules import qualifying_limit_80g, eligible_donations_80g
            comps = comps or {}
            via_before = sum(v for k, v in comps.items()
                             if k in ("80C", "80CCC", "80CCD1", "80CCD1B", "80CCD2",
                                      "80CCG", "80D", "80DD", "80DDB", "80E", "80EE"))
            via_after = sum(v for k, v in comps.items()
                            if k in ("80GGA", "80GGC", "80TTA", "80TTB", "80U", "80CCH"))
            via_others = sum(v for k, v in comps.items() if k != "80G")
            capped_others = min(float(cv.user_total or 0), via_others,
                                max(0.0, self.gti()))
            ti_pass = via.round_total_income(self.gti() - capped_others)
            ql = qualifying_limit_80g(ti_pass, via_before, via_after,
                                      comps.get("80GG", 0), comps.get("80EEA", 0),
                                      comps.get("80EEB", 0))
            return min(eligible_donations_80g(cv.donations_80g, ql), max(0.0, ti_pass))
        return cv.s80g_schedule_amount

    def total_chapter_via(self) -> int:
        """[AO163]."""
        comps = self.chapter_via_components()
        return via.total_chapter_via(comps, self.chapter_via.user_total, self.gti())

    def total_income_excl_112a(self) -> int:
        """C1/D1 base [AO164] ROUND(MAX(0, GTI - VIA), -1)."""
        return via.round_total_income(self.gti() - self.total_chapter_via())

    def total_income_incl_112a(self) -> int:
        """C2 [AO165] (includes LTCG 112A)."""
        return via.round_total_income(self.gti_incl_112a() - self.total_chapter_via())

    # ---- Part D ------------------------------------------------------
    def d1_total_tax_payable(self) -> int:
        """[VBA calcTaxPayableOnTI / calcTaxPayableOnTINTR -> AO176]"""
        return tax_on_total_income(self.total_income_excl_112a(), self.regime,
                                   self.personal.age)

    def d2_rebate_87a(self) -> int:
        """[AO177]"""
        return rebate_87a(self.d1_total_tax_payable(),
                          self.total_income_excl_112a(),
                          self.total_income_incl_112a(), self.regime)

    def d3_tax_after_rebate(self) -> int:      # [AO178]
        return excel_round(self.d1_total_tax_payable() - self.d2_rebate_87a())

    def d4_cess(self) -> int:                  # [AO180]
        return education_cess(self.d3_tax_after_rebate())

    def d5_total_tax_and_cess(self) -> int:    # [AO181]
        return excel_round(self.d3_tax_after_rebate() + self.d4_cess())

    def relief_89a_computed(self) -> int:
        """13a [AO183] proportional apportionment of tax for 89A income."""
        ti = self.total_income_excl_112a()
        if self.d1_total_tax_payable() <= 0 or ti <= 0:
            return 0
        notified = self.salary.retirement_benefit.notified_total
        return excel_round((notified / ti) * self.d1_total_tax_payable())

    def balance_tax_after_relief(self) -> int:  # NetTaxLiability [AO184]
        return excel_round(max(self.d5_total_tax_and_cess()
                               - self.relief_89 - self.relief_89a_computed(), 0))

    def net_tax_liability_for_interest(self) -> float:
        """NetTaxLiability after subtracting relief, before credits
        [VBA IncD.NetTaxLiability]."""
        return self.balance_tax_after_relief()

    def d7_interest_234a(self) -> int:
        vdate = self.verification.verification_date or C.ITR_FILING_DUE_DATE
        ntl = self.balance_tax_after_relief() - (self.taxes_paid.total_tds
                                                 + self.taxes_paid.total_tcs
                                                 + self.taxes_paid.total_it)
        return interest_234a(max(0.0, ntl), self.filing.due_date, vdate)

    def d8_interest_234b(self) -> int:
        vdate = self.verification.verification_date or date.today()
        return interest_234b(self.balance_tax_after_relief(),
                             self.taxes_paid.advance_tax,
                             self.taxes_paid.total_tds, self.taxes_paid.total_tcs,
                             vdate, self.personal.age,
                             RETURN_FILE_SEC.get(self.filing.filing_sec_code, ""))

    def d10_fee_234f(self) -> int:
        return fee_234f(self.total_income_excl_112a(),
                        self.filing.filed_after_due_date)

    def d11_total_tax_fee_interest(self) -> int:   # [AO191]
        return excel_round(self.d5_total_tax_and_cess() + self.d7_interest_234a()
                           + self.d8_interest_234b() + self.d9_interest_234c()
                           + self.d10_fee_234f() + self.d10a_fee_234i() - self.relief_89)

    def quarterly_tax_estimates(self) -> list:
        """[VBA Q1Tax..Q5Tax] cumulative quarterly slab tax on income spread.
        Income is spread across the 5 utility periods using Other Sources
        quarterly grids + salary/HP assumed evenly across the year."""
        from .tax_engine import tax_new_regime, tax_old_regime
        os_ = self.other_sources
        q_div = list(os_.dividend_by_quarter) + [0.0] * (5 - len(os_.dividend_by_quarter))
        annual_rest = (self.income_from_salaries()
                       + hp_for_gti(self.house_properties, self.regime)
                       + os_.other_income_grid_total)
        per_q = annual_rest / 4
        cum, out = 0.0, []
        for k in range(5):
            if k < 4:
                cum += per_q + q_div[k]
            out.append(cum)
        taxes = []
        for cum_inc in out:
            t = (tax_new_regime(cum_inc) if self.regime == C.REGIME_NEW
                 else tax_old_regime(cum_inc, self.personal.age))
            taxes.append(t)
        taxes[4] = max(0, taxes[4] - taxes[3])   # [VBA] Q5Tax = tax - Q4Tax
        return taxes

    def d9_interest_234c(self) -> int:
        """[VBA:mIncmDtls.calcIntrst234C] - full port in interest.py."""
        from .interest import interest_234c
        if not self.taxes_paid.instalments:
            return 0
        return interest_234c(self.quarterly_tax_estimates(), self.d2_rebate_87a(),
                             self.taxes_paid.instalments,
                             self.taxes_paid.total_tds, self.taxes_paid.total_tcs,
                             self.relief_89, self.relief_89a_computed(),
                             self.balance_tax_after_relief(), self.personal.age)

    def d10a_fee_234i(self) -> int:
        from .interest import fee_234i
        return fee_234i(self.filing.filing_sec_code == 17, self.d10_fee_234f())

    # ---- Refund --------------------------------------------------------
    def total_taxes_paid(self) -> float:
        return (self.taxes_paid.total_tds + self.taxes_paid.total_tcs
                + self.taxes_paid.total_it)

    def refund_or_payable(self) -> float:
        """Part B ATI balance: paid - liability (positive => refund)."""
        return self.total_taxes_paid() - self.d11_total_tax_fee_interest()

    # ------------------------------------------------------------------
    def compute_summary(self) -> dict:
        """All Part-D internal fields (schema: ITR1_TaxComputation)."""
        return {
            "TotalTaxPayable": self.d1_total_tax_payable(),
            "Rebate87A": self.d2_rebate_87a(),
            "TaxPayableOnRebate": self.d3_tax_after_rebate(),
            "EducationCess": self.d4_cess(),
            "GrossTaxLiability": self.d5_total_tax_and_cess(),
            "Section89": excel_round(self.relief_89),
            "Section89A": self.relief_89a_computed(),
            "NetTaxLiability": self.balance_tax_after_relief(),
            "IntrstPay234A": self.d7_interest_234a(),
            "IntrstPay234B": self.d8_interest_234b(),
            "IntrstPay234C": self.d9_interest_234c(),
            "FeeIncUS234F": self.d10_fee_234f(),
            "FeeUS234I": self.d10a_fee_234i(),
            "TotalIntrstPay": excel_round(self.d7_interest_234a() + self.d8_interest_234b()
                                          + self.d9_interest_234c() + self.d10_fee_234f()
                                          + self.d10a_fee_234i()),
            "TotTaxPlusIntrstPay": self.d11_total_tax_fee_interest(),
            "GrossTotIncome": self.gti(),
            "GrossTotIncomeIncLTCG112A": self.gti_incl_112a(),
            "TotalIncome": self.total_income_excl_112a(),
            "TotalIncomeIncl112A": self.total_income_incl_112a(),
            "TotalChapterVIA": self.total_chapter_via(),
            "ChapterVIAComponents": self.chapter_via_components(),
            "RefundOrPayable": self.refund_or_payable(),
        }
