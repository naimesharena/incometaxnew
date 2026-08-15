from fastapi import APIRouter, Query
from ..services import master_data

router = APIRouter(prefix="/api/master", tags=["master"])

@router.get("/bank-codes")
def get_bank_codes(q: str = Query("", description="Search query"), limit: int = 20):
    if q:
        return master_data.search_bank(q, limit)
    else:
        return master_data.load_bank_codes()[:limit]

@router.get("/pincode/{pincode}")
def get_pincode_info(pincode: str):
    state = master_data.search_pincode(pincode)
    if state:
        return {"pincode": pincode, "state": state, "valid": True}
    else:
        return {"pincode": pincode, "state": None, "valid": False}

@router.get("/ifsc/validate")
def validate_ifsc_endpoint(ifsc: str):
    valid = master_data.validate_ifsc(ifsc)
    return {"ifsc": ifsc.upper(), "valid": valid}

@router.get("/ifsc/search")
def search_ifsc(q: str = Query("", description="IFSC prefix"), limit: int = 20):
    all_ifsc = master_data.load_ifsc_list_paginated()
    q = q.upper()
    filtered = [code for code in all_ifsc if code.startswith(q)]
    return {"query": q, "results": filtered[:limit], "total": len(filtered)}

@router.get("/states")
def get_states():
    return {"states": master_data.get_state_list()}

@router.get("/employer-categories")
def get_employer_categories():
    return {"categories": master_data.get_employer_categories()}

@router.get("/dropdowns")
def get_all_dropdowns():
    return master_data.load_dropdowns()

@router.get("/hash-meta")
def get_hash_meta():
    return master_data.load_hash_meta()

@router.get("/return-file-sections")
def get_file_sections():
    return {"sections": master_data.get_return_file_sections()}
