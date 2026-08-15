import openpyxl
from openpyxl import load_workbook
import sys, os, json, pathlib, zipfile
from collections import defaultdict

files = [
 "ITR-1/ITR1_AY_26-27_V1.2.xlsm",
 "ITR-2/ITR2_AY_26-27_V1.3.xlsm",
 "ITR-3/ITR3_AY_26-27_V1.2.xlsm",
 "ITR-4/ITR4_AY_26-27_V1.1.xlsm",
]

def analyze_file(path):
    print(f"\n{'='*80}\n ANALYZING {path}\n{'='*80}")
    wb = load_workbook(path, data_only=False, keep_vba=True)
    print(f"Sheet names ({len(wb.sheetnames)}): {wb.sheetnames}")
    # Examine sheet states
    for idx, ws in enumerate(wb.worksheets):
        # openpyxl stores sheet_state
        state = ws.sheet_state
        print(f"  [{idx}] {ws.title} state={state} max_row={ws.max_row} max_col={ws.max_column} dims={ws.dimensions} ")

    print("\n--- Named ranges ---")
    # openpyxl: wb.defined_names
    try:
        for dn_name in wb.defined_names:
            dn = wb.defined_names[dn_name]
            print(f"  Name: {dn.name} -> {dn.value} attr text? {str(dn.attr_text)[:500]}")
    except Exception as e:
        print(f"  Error listing defined names: {e}")
        # fallback
        try:
            for dn in wb.defined_names.definedName:
                print(f"  Name: {dn.name} -> {dn.value} attr text? {dn.attr_text[:200]}")
        except Exception as ex2:
            print(str(ex2))

    # Check externalLinks etc
    print("\n--- Data Validations sampling ---")
    for ws in wb.worksheets:
        if ws.data_validations.dataValidation:
            print(f"  Sheet {ws.title} has {len(ws.data_validations.dataValidation)} validations")
            for dv in ws.data_validations.dataValidation[:5]:
                print(f"    sqref={dv.sqref} type={dv.type} formula1={str(dv.formula1)[:200]} formula2={str(dv.formula2)[:200]}")
            break

    # formulas
    print("\n--- Formulas sample ---")
    formula_count=0
    tables=[]
    dropdowns=[]
    for ws in wb.worksheets:
        # tables
        if ws._tables:
            for t in ws._tables.values():
                tables.append((ws.title, t.displayName, t.ref))
        # count formulas
        for row in ws.iter_rows(min_row=1, max_row=min(200, ws.max_row), max_col=min(30, ws.max_column)):
            for cell in row:
                if cell.data_type == 'f' or (cell.value and isinstance(cell.value, str) and cell.value.startswith('=')):
                    formula_count+=1
                    if formula_count<30:
                        print(f"  {ws.title}!{cell.coordinate} = {cell.value}")
        # data validations for dropdowns already
    print(f"Total formula sample counted in first 200 rows per sheet (partial): {formula_count}")
    print(f"Tables: {tables[:20]}")

    # Check for hidden sheets more accurate via workbook.xml parsing
    # openpyxl already gives sheet_state but also veryHidden
    very_hidden = [ws.title for ws in wb.worksheets if ws.sheet_state != 'visible']
    print(f"\nHidden / Very Hidden sheets: {very_hidden}")

    # Check for conditional formatting?
    cf_count = sum(len(ws.conditional_formatting._cf_rules) for ws in wb.worksheets)
    print(f"Conditional formatting rules total: {cf_count}")

    # Try VBA extraction via oletools
    try:
        from oletools.olevba import VBA_Parser
        vba_parser = VBA_Parser(path)
        if vba_parser.detect_vba_macros():
            print("\n--- VBA Macros detected ---")
            for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                # print first 500 chars
                lines = vba_code.splitlines()
                print(f"  Stream: {filename} | {stream_path} | {vba_filename} | lines={len(lines)}")
                # print code preview
                preview = "\n".join(lines[:50])
                print(preview[:2000])
                print("... (truncated)")
                print("---")
        else:
            print("No VBA macros detected by oletools")
        vba_parser.close()
    except Exception as e:
        print(f"VBA extraction error: {e}")
        import traceback; traceback.print_exc()

    wb.close()

for f in files:
    analyze_file(f)
