"""Constants for AY 2026-27 - all values sourced from the CBDT Excel utility.

Source key:
  [IncD!<cell>]            Income Details sheet, ITR1_AY_26-27_V1.2.xlsm
  [VBA:<module>.<proc>]    decompiled VBA (itr_extraction/vba/ITR1.vba)
  [DV:<sheet>!<range>]     data-validation rule (dropdown list)
"""
from datetime import date

ASSESSMENT_YEAR = "2026"
SCHEMA_VER = "1.1"
FORM_VER = "1.2"          # matches workbook version ITR1_AY_26-27_V1.2
FORM_NAME_ITR1 = "ITR-1"
DESCRIPTION_ITR1 = "For individuals being a RESIDENT (OTHER THAN NOT ORDINARILY RESIDENT) having income from Salary/Pension, one House Property, Other Sources and Agricultural income up to Rs.5000"

# Filing due date - [IncD!P53] hardcodes 31/07/2026 for AY 2026-27
ITR_FILING_DUE_DATE = date(2026, 7, 31)
DUE_DATE_STR = "31/07/2026"

# --------------------------------------------------------------------------
# Dropdown / code lists (from DV rules + hidden DataBase/DropDownValues sheets)
# --------------------------------------------------------------------------
# [DV:Income Details!Z11 -> EmpCatList] - codes per schema EmployerCategory enum
# CGOV Central Govt / SGOV State Govt / PSU Public Sector / PE Private Employee
# PESG/PEPS/PEO variants / OTH Other / NA Not Applicable  [VBA: iEmpCat checks]
NATURE_OF_EMPLOYMENT = ["CGOV", "SGOV", "PSU", "PE", "PESG", "PEPS", "PEO", "OTH", "NA"]
# Official CBDT StateCode list (ITR-1 schema 'Address.StateCode' description).
# NOTE: the Excel StateList dropdown carries the same 2-digit codes.
STATE_CODES = {
    "01": "Andaman and Nicobar Islands", "02": "Andhra Pradesh",
    "03": "Arunachal Pradesh", "04": "Assam", "05": "Bihar",
    "06": "Chandigarh", "07": "Dadra Nagar and Haveli", "08": "Daman and Diu",
    "09": "Delhi", "10": "Goa", "11": "Gujarat", "12": "Haryana",
    "13": "Himachal Pradesh", "14": "Jammu and Kashmir", "15": "Karnataka",
    "16": "Kerala", "17": "Lakshadweep", "18": "Madhya Pradesh",
    "19": "Maharashtra", "20": "Manipur", "21": "Meghalaya", "22": "Mizoram",
    "23": "Nagaland", "24": "Odisha", "25": "Puducherry", "26": "Punjab",
    "27": "Rajasthan", "28": "Sikkim", "29": "Tamil Nadu", "30": "Tripura",
    "31": "Uttar Pradesh", "32": "West Bengal", "33": "Chhattisgarh",
    "34": "Uttarakhand", "35": "Jharkhand", "36": "Telangana",
    "37": "Ladakh", "99": "Foreign",
}
COUNTRY_INDIA = "91"       # schema CountryCode numeric list; India = 91
COUNTRY_DEFAULT = "91-INDIA"  # workbook display label [HP!S41 check]

# [DV:Income Details!AN31] 115BAC option question (row 31)
YES_NO_SELECT = ["(Select)", "Yes", "No"]

# Residential status (ITR-1 allows residents only)
RESIDENT = "RES-Resident"

# Nature of exempt allowances u/s 10 grid (Others.NOI list, trimmed)
ALLOWANCES_U_S10 = [
    "Sec 10(5) Travel Concession/assistance",
    "Sec 10(10) Death cum Retirement Gratuity",
    "Sec 10(10AA) Earned Leave",
    "Sec 10(11) Statutory Provident Fund received",
    "Sec 10(12) Recognized Provident Fund received",
    "Sec 10(13) Approved Superannuation Fund received",
    "Sec 10(13A) House Rent Allowance",
    "Sec 10(14)(i) Prescribed allowances/benefits",
    "Sec 10(14)(ii) Allowance to meet expenditure",
    "Any Other",
]

# --------------------------------------------------------------------------
# Regime semantics
# --------------------------------------------------------------------------
# [VBA: multiple procs reference Sheet5.Range("BacValue")]
# BacValue=1 -> opted NEW tax regime u/s 115BAC(6)   (default)
# BacValue=2 -> opted OUT of 115BAC -> OLD regime
REGIME_NEW = 1
REGIME_OLD = 2

# --------------------------------------------------------------------------
# Monetary caps (AY 2026-27) - each with Excel source
# --------------------------------------------------------------------------
STD_DEDUCTION_NEW = 75000      # [IncD!AO73] IF(BacValue=1,MIN(Net_salary,75000),...)
STD_DEDUCTION_OLD = 50000      # [IncD!AO73] ...IF(BacValue=2,MIN(Net_salary,50000),0)
CAP_80C_80CCC = 150000         # [IncD!AN115/AN116] combined 80C+80CCC cap
CAP_80CCD_1B = 50000           # [IncD!AN126] MIN(...,50000,...)
CAP_80CCG = 25000              # [IncD!AN138]
CAP_80CCG_TI_LIMIT = 1200000   # [IncD!AN138] IF(TOTAL_INCOME>1200000,0,...)
CAP_80EE = 50000               # [IncD!AN148]
CAP_80EEA = 150000             # [IncD!AN149]
CAP_80EEB = 150000             # [IncD!AN150]
CAP_80TTA = 10000              # [IncD!AN158]
CAP_80TTB = 50000              # [IncD!AN159]
CAP_80U_NORMAL = 75000         # [IncD BG160/161 formula] severity 1
CAP_80U_SEVERE = 125000        # severity 2
CAP_80GG_WITH_HRA = 55000      # [IncD!AN152] IF(HRA>0,55000,60000)
CAP_80GG_WITHOUT_HRA = 60000
HP_LOSS_CAP = 200000           # [HP!76:76] MAX(-200000, HP1+HP2)
HP_SELF_OCCUPIED_INTEREST_CAP = 200000   # [IncD!AO82]
CAP_80D_OVERALL = 100000       # [IncD!AN141] MIN(100000, ...)
CAP_80D_SELF = 25000           # standard rule per Schedule 80D (hidden sheet 80D)
CAP_80D_SELF_SENIOR = 50000
CAP_80D_PARENTS = 25000
CAP_80D_PARENTS_SENIOR = 50000
CAP_80D_CHECKUP = 5000         # preventive health check-up (within above limits)
CAP_80DD_NORMAL = 75000        # [VBA/Sch80D sheet values]
CAP_80DD_SEVERE = 125000
CAP_57iia_FAMILY_PENSION = 15000   # deduction u/s 57(iia)
LTCG_112A_MAX_ITR1 = 125000    # [VBA validation msg] "In ITR 1, the maximum gains as per Section 112A can be INR 1,25,000/-"
LTCG_112A_EXEMPT_LIMIT = 125000    # [IncD!AO175]
REBATE_87A_NEW_LIMIT = 60000   # [IncD!AO177]
REBATE_87A_NEW_TI = 1200000
REBATE_87A_OLD_LIMIT = 12500   # [IncD!AO177] BacValue=2 branch
REBATE_87A_OLD_TI = 500000
CESS_RATE = 0.04               # [IncD!AO180] *0.04 Health & Education Cess
FEE_234F_LOW = 1000            # [VBA ~18470] intrst234F = 1000 (TI<=5L, late)
FEE_234F_HIGH = 5000           # [VBA ~18477]
TI_234F_THRESHOLD = 500000
INTEREST_234B_NET_LIMIT = 10000   # [VBA:mdCalInterst234B] CONST_NET_Limit
INTEREST_234B_ATP_PCT = 90        # [VBA] CONST_ATP_Limit = 90 (% of advance tax paid)
AGNIPATH_80CCH_CAP = 288000    # [IncD!AN161]
AGNIPATH_80CCH_PCT = 46.2      # [IncD!AN161] 46.2% of allowances
