import json
import pathlib
from functools import lru_cache

DATA_DIR = pathlib.Path(__file__).parent.parent / "data"

@lru_cache(maxsize=1)
def load_bank_codes():
    path = DATA_DIR / "bank_codes.json"
    if not path.exists():
        return []
    with open(path) as f:
        return json.load(f)

@lru_cache(maxsize=1)
def load_pincode():
    path = DATA_DIR / "pincode.json"
    if not path.exists():
        return {}
    with open(path) as f:
        return json.load(f)

@lru_cache(maxsize=1)
def load_ifsc():
    path = DATA_DIR / "ifsc.json"
    if not path.exists():
        return []
    with open(path) as f:
        data = json.load(f)
        # return as set for quick lookup
        return set(data)

@lru_cache(maxsize=1)
def load_ifsc_list_paginated():
    # for API pagination we need list
    path = DATA_DIR / "ifsc.json"
    if not path.exists():
        return []
    with open(path) as f:
        return json.load(f)

@lru_cache(maxsize=1)
def load_dropdowns():
    path = DATA_DIR / "dropdowns.json"
    if not path.exists():
        return {}
    with open(path) as f:
        return json.load(f)

@lru_cache(maxsize=1)
def load_hash_meta():
    path = DATA_DIR / "hash_meta.json"
    if not path.exists():
        return {"hash_key": "7Z3mxclnABiXtYG", "hash_iteration": 1849}
    with open(path) as f:
        return json.load(f)

def search_bank(query: str, limit=20):
    banks = load_bank_codes()
    q = query.upper()
    results = [b for b in banks if q in b["code"] or q in b["name"].upper()]
    return results[:limit]

def search_pincode(pin: str):
    pmap = load_pincode()
    return pmap.get(pin)

def validate_ifsc(ifsc: str) -> bool:
    ifsc_set = load_ifsc()
    # pattern also: 4 alpha + 0 + 6 alphanumeric
    import re
    pattern = r"^[A-Z]{4}0[A-Z0-9]{6}$"
    if not re.match(pattern, ifsc.upper()):
        return False
    # if we have full list, check existence
    if ifsc_set and ifsc.upper() not in ifsc_set:
        # Allow even if not in list but pattern matches (new IFSC may not be in our snapshot)
        return True
    return True

def get_state_list():
    dropdowns = load_dropdowns()
    # State key may exist
    for key in ["State", "STATE", "State List"]:
        if key in dropdowns:
            return dropdowns[key]
    # fallback extract from pincode values unique
    pmap = load_pincode()
    states = sorted(list(set(pmap.values())))
    return states

def get_employer_categories():
    dropdowns = load_dropdowns()
    return dropdowns.get("EmployerCategory") or dropdowns.get("Employer Category") or ["State Government","Central Government","Public Sector Undertaking","Pensioners","Others","Not Applicable (eg. Family pension etc)"]

def get_return_file_sections():
    dropdowns = load_dropdowns()
    return dropdowns.get("ReturnFileUnderSection") or ["139(1)","139(4)","139(5)-Revised","119(2)(b)- After condonation of delay","139(8A)"]
