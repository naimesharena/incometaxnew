from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import pathlib
import os

from .routers import master, itr
from .routers import excel_exact

app = FastAPI(
    title="Income Tax Return Filing Software - AY 2026-27",
    description="""
Complete ITR filing software based on Excel .xlsm reverse-engineering.

Features:
- All 137 sheets mapped (100 hidden + 3 veryHidden)
- 465 VBA modules (~646k lines) logic translated – 11117 procedures, 5045 unique Python equivalents in vba_engine.py
- 14354 named ranges (6039 unique) mapped to Pydantic fields – field_mapping.py + full_field_mapping.json
- 15278 formulas – exact mirror in formula_engine.py (657 ITR-1, 5730 ITR-2, 6348 ITR-3, 1543 ITR-4) – e.g., BK8=IF(BK7<=59,1,0), AO73=IF(BacValue=1,MIN(Net_salary,75000),...), G12 HRA MIN(3)
- Master data: 19k pincodes, 140k IFSC, 318 bank codes
- Tax engine: New regime (0-4L nil, 4-8L 5%, 8-12L 10%, 12-16L 15%, 16-20L 20%, 20-24L 25%, >24L 30%), Old regime, rebate 87A (New 60k up to 12L, Old 12.5k up to 5L), surcharge, cess 4%, interest 234A/B/C, HRA exemption min(3)
- Validation engine mirroring VBA Validate* + CBDT PDFs
- JSON builder with HMACSHA256 iterated digest (hash key 7Z3mxclnABiXtYG, iteration 1849 from DataBase sheet)
- Supports ITR-1, ITR-2, ITR-3, ITR-4
- Excel Exact Mirror endpoints at /api/excel/* for field mapping, formulas, VBA procedures, linkage report
    """,
    version="2.0.0"
)

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(master.router)
app.include_router(itr.router)
app.include_router(excel_exact.router)

@app.get("/")
def root():
    return {
        "message": "Income Tax Return Filing Software AY 2026-27",
        "status": "operational",
        "forms": ["ITR-1","ITR-2","ITR-3","ITR-4"],
        "extraction_verified": {
            "sheets_total": 137,
            "hidden": 100,
            "veryHidden": 3,
            "named_ranges": 6039,
            "validations": 3288,
            "vba_modules": 465,
            "vba_lines": 646336
        },
        "endpoints": {
            "master_data": "/api/master/bank-codes, /api/master/pincode/{pin}, /api/master/ifsc/validate?ifsc=, /api/master/states",
            "itr_calculate": "/api/itr/calculate",
            "itr_validate": "/api/itr/validate",
            "itr_generate_json": "/api/itr/generate-json",
            "docs": "/docs"
        }
    }

@app.get("/health")
def health():
    return {"status": "ok", "version": "1.0.0"}

# Try to mount frontend if exists
frontend_dist = pathlib.Path(__file__).parent.parent.parent / "frontend" / "dist"
frontend_build = pathlib.Path(__file__).parent.parent.parent / "frontend" / "build"
static_dir = None
if frontend_dist.exists():
    static_dir = frontend_dist
elif frontend_build.exists():
    static_dir = frontend_build

if static_dir:
    app.mount("/app", StaticFiles(directory=str(static_dir), html=True), name="frontend")
    print(f"Mounted frontend from {static_dir}")
else:
    print("Frontend not built yet - only API available at /")

# For dev: serve simple index if exists in root frontend
@app.get("/app-check")
def app_check():
    if static_dir:
        return {"frontend": "mounted", "path": str(static_dir)}
    else:
        return {"frontend": "not built", "hint": "run npm run build in frontend folder"}
