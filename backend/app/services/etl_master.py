"""
ETL Master Data from ITR xlsm hidden sheets into JSON files for backend usage.
Extracts:
- Bank Codes (BankCode sheet)
- Pincode -> State (DataBase CP:CQ)
- IFSC codes (IFSC sheet E:I columns concatenated)
- Dropdowns: EmployerCategory, Country, State, ReturnFileSection, etc from DataBase header areas
- Hash Key / Iteration from DataBase header
"""
import openpyxl
import json
import os
import pathlib

SRC = "ITR-1/ITR1_AY_26-27_V1.2.xlsm"
OUT_DIR = pathlib.Path(__file__).parent.parent / "data"
OUT_DIR.mkdir(exist_ok=True)

wb = openpyxl.load_workbook(SRC, data_only=True)
print(f"Loaded {SRC}")

# 1. Bank Codes
ws_bc = wb["BankCode"]
bank_codes = []
for r in range(1, ws_bc.max_row+1):
    code = ws_bc.cell(r,1).value
    name = ws_bc.cell(r,2).value
    if code:
        bank_codes.append({"code": str(code).strip(), "name": (str(name).strip() if name else "")})
print(f"Bank codes extracted {len(bank_codes)}")

with open(OUT_DIR/"bank_codes.json","w") as f:
    json.dump(bank_codes, f, indent=2)

# 2. Pincode mapping
ws_db = wb["DataBase"]
pincode_map = {}
# CP=94, CQ=95
for r in range(2, ws_db.max_row+1):
    pin = ws_db.cell(r,94).value
    state = ws_db.cell(r,95).value
    if pin is None:
        continue
    try:
        pin_str = str(int(pin)) if isinstance(pin,(int,float)) else str(pin).strip()
    except:
        pin_str = str(pin).strip()
    if pin_str and state:
        pincode_map[pin_str] = str(state).strip()
    if r>20000: # safety
        break
print(f"Pincode map {len(pincode_map)}")

with open(OUT_DIR/"pincode.json","w") as f:
    json.dump(pincode_map, f, indent=2)

# 3. IFSC - sheet has 5 columns E..I (5..9) each listing codes from row 3 onwards
ws_ifsc = wb["IFSC"]
ifsc_codes = set()
for c in range(5,10):
    for r in range(3, ws_ifsc.max_row+1):
        val = ws_ifsc.cell(r,c).value
        if val:
            ifsc_codes.add(str(val).strip().upper())
        # early break if empty long?
print(f"IFSC unique {len(ifsc_codes)}")

# To keep file size manageable, save as list but also split sample: we will save full list may be 45k, okay
with open(OUT_DIR/"ifsc.json","w") as f:
    json.dump(sorted(list(ifsc_codes)), f, indent=2)

# 4. Dropdowns from DataBase header
# DataBase row1 contains header names, row3 onwards contains values per column
# We'll parse first 50 cols, group by header
dropdowns = {}
# Map header row1 to column index
header_row = 1
for c in range(1, 70):
    header = ws_db.cell(header_row,c).value
    if header:
        header = str(header).strip()
        # collect values from row3 onwards until empty run >50?
        values=[]
        empty_streak=0
        for r in range(3, 500):
            v = ws_db.cell(r,c).value
            if v is None or str(v).strip()=="" or str(v).strip()=="(Select)":
                empty_streak+=1
                if empty_streak>20:
                    break
                continue
            empty_streak=0
            # keep distinct
            vs=str(v).strip()
            if vs not in values and vs!="(Select)":
                values.append(vs)
        if values:
            dropdowns[header]=values

print(f"Dropdowns extracted {len(dropdowns)} keys: {list(dropdowns.keys())[:20]}")

with open(OUT_DIR/"dropdowns.json","w") as f:
    json.dump(dropdowns, f, indent=2)

# 5. Hash key / iteration
hash_key = ws_db.cell(3,2).value
hash_iter = ws_db.cell(4,2).value
date_processing = ws_db.cell(1,2).value
print(f"HashKey {hash_key} Iteration {hash_iter} DateProcessing {date_processing}")

with open(OUT_DIR/"hash_meta.json","w") as f:
    json.dump({"hash_key": hash_key, "hash_iteration": hash_iter, "date_processing": str(date_processing)}, f, indent=2)

wb.close()
print(f"ETL Done, files in {OUT_DIR}")
