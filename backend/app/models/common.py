from pydantic import BaseModel, Field, field_validator
from typing import Optional, List, Dict, Any
import re

PAN_REGEX = r"^[A-Z]{5}[0-9]{4}[A-Z]$"
AADHAAR_REGEX = r"^\d{12}$"

class PersonalInfo(BaseModel):
    first_name: str = Field(..., min_length=1, max_length=25)
    middle_name: Optional[str] = ""
    last_name: str = Field(..., min_length=1, max_length=25)
    pan: str = Field(..., pattern=PAN_REGEX)
    aadhaar: Optional[str] = Field(None, pattern=AADHAAR_REGEX)
    dob: str = Field(..., description="YYYY-MM-DD")
    flat_no: Optional[str] = ""
    premises: Optional[str] = ""
    street: Optional[str] = ""
    area: Optional[str] = ""
    city: str = ""
    state_code: str = ""
    country_code: str = "91"
    pincode: str = Field(..., pattern=r"^[1-9][0-9]{5}$")
    mobile: str = Field(..., pattern=r"^[6-9][0-9]{9}$")
    email: str = Field(..., pattern=r"^[^@]+@[^@]+\.[^@]+$")
    filing_section: str = "139(1)"
    residential_status: str = "RES"
    employer_category: Optional[str] = "OTH"

    @field_validator('pan')
    def upper_pan(cls, v):
        return v.upper()

class IncomeDetails(BaseModel):
    salary: float = 0
    basic: float = 0
    da: float = 0
    hra_received: float = 0
    rent_paid: float = 0
    is_metro: bool = False
    employer_category: str = "OTH"
    hp_income: float = 0
    hp_type: Optional[str] = "self_occupied"  # self_occupied, let_out
    other_sources: float = 0
    other_sources_interest: float = 0
    other_sources_dividend: float = 0

class Deduction(BaseModel):
    deduction_80C: float = 0
    deduction_80CCC: float = 0
    deduction_80CCD1: float = 0
    deduction_80CCD1B: float = 0
    deduction_80CCD2: float = 0
    deduction_80D: float = 0
    deduction_80D_preventive: float = 0
    deduction_80DD: float = 0
    deduction_80DDB: float = 0
    deduction_80E: float = 0
    deduction_80EE: float = 0
    deduction_80EEA: float = 0
    deduction_80EEB: float = 0
    deduction_80G: float = 0
    donations_80G: List[Dict[str, Any]] = []
    deduction_80GGA: float = 0
    deduction_80GGC: float = 0
    deduction_80U: float = 0

class TDSEntry(BaseModel):
    tan: str
    employer_name: str
    income_chargeable: float = 0
    tax_deducted: float = 0
    year: str = "2025"

class TDSData(BaseModel):
    tds_salary: List[TDSEntry] = []
    tds_other: List[TDSEntry] = []
    tcs_total: float = 0

class BankAccount(BaseModel):
    ifsc: str = Field(..., pattern=r"^[A-Z]{4}0[A-Z0-9]{6}$")
    bank_name: Optional[str] = ""
    account_number: str
    account_type: str = "Saving"
    is_for_refund: bool = False

class BankData(BaseModel):
    bank_accounts: List[BankAccount] = []

class TaxesPaid(BaseModel):
    advance_tax: float = 0
    self_assessment: float = 0
    advance_tax_payments: List[Dict[str, Any]] = []

class ITRPayload(BaseModel):
    itr_form: str = Field("ITR-1", pattern=r"ITR-[1-4]")
    regime: str = Field("NEW", pattern=r"^(OLD|NEW)$")
    age: int = 30
    is_resident: bool = True
    personal_info: PersonalInfo
    income_details: IncomeDetails
    deductions: Deduction
    tds: TDSData
    bank: BankData
    taxes_paid: TaxesPaid
    filing_date: Optional[str] = "2026-07-31"
    due_date: Optional[str] = "2026-07-31"
    part_a_gen_1398A: Optional[Dict[str, Any]] = {}

class CalculationRequest(BaseModel):
    payload: ITRPayload
