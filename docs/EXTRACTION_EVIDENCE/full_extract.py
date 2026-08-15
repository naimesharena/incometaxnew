import openpyxl
from openpyxl import load_workbook
import zipfile, os, re, json, pathlib
from collections import Counter

files = [
 "ITR-1/ITR1_AY_26-27_V1.2.xlsm",
 "ITR-2/ITR2_AY_26-27_V1.3.xlsm",
 "ITR-3/ITR3_AY_26-27_V1.2.xlsm",
 "ITR-4/ITR4_AY_26-27_V1.1.xlsm",
]

# Helper to extract VBA module names via oletools
def get_vba_info(path):
    result = []
    try:
        from oletools.olevba import VBA_Parser
        parser = VBA_Parser(path)
        if parser.detect_vba_macros():
            for (filename, stream_path, vba_filename, vba_code) in parser.extract_macros():
                # get procedures via regex
                proc_regex = re.compile(r'^\s*(Public|Private)?\s*(Sub|Function)\s+(\w+)', re.MULTILINE | re.IGNORECASE)
                procs = proc_regex.findall(vba_code)
                proc_names = [p[2] for p in procs]
                # also detect event handlers Worksheet_Change etc
                result.append({
                    "stream_file": filename,
                    "stream_path": stream_path,
                    "module": vba_filename,
                    "lines": len(vba_code.splitlines()),
                    "procs_count": len(proc_names),
                    "procs": proc_names[:200],  # limit
                    "code_size": len(vba_code)
                })
        parser.close()
    except Exception as e:
        result.append({"error": str(e)})
    return result

def extract_workbook_structure(path):
    data = {}
    wb = load_workbook(path, data_only=False, keep_vba=True)
    data["file"] = path
    data["sheets"] = []
    total_formulas = 0
    total_validations = 0
    total_cf = 0
    total_tables = 0
    for ws in wb.worksheets:
        # count formulas scanning limited but attempt full scan
        formula_examples = []
        formula_count = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=min(ws.max_column, 100)):
            for cell in row:
                if cell.data_type == 'f' or (cell.value and isinstance(cell.value, str) and cell.value.startswith('=')):
                    formula_count+=1
                    if len(formula_examples)<15:
                        formula_examples.append(f"{cell.coordinate}={str(cell.value)[:300]}")
        total_formulas+=formula_count

        validations = len(ws.data_validations.dataValidation) if ws.data_validations else 0
        total_validations+=validations
        cf = len(ws.conditional_formatting._cf_rules) if hasattr(ws.conditional_formatting, '_cf_rules') else 0
        total_cf+=cf
        tables = list(ws._tables.values()) if ws._tables else []
        total_tables+=len(tables)

        merged = len(ws.merged_cells.ranges) if ws.merged_cells else 0

        # Data validations detailed list first 10
        dv_list = []
        for dv in (ws.data_validations.dataValidation[:10] if validations else []):
            dv_list.append({"sqref": str(dv.sqref), "type": dv.type, "formula1": str(dv.formula1)[:300], "formula2": str(dv.formula2)[:300] if dv.formula2 else ""})

        data["sheets"].append({
            "title": ws.title,
            "state": ws.sheet_state,
            "max_row": ws.max_row,
            "max_col": ws.max_column,
            "dimension": ws.dimensions,
            "formula_count": formula_count,
            "formulas_sample": formula_examples,
            "validations_count": validations,
            "validations_sample": dv_list,
            "cf_rules": cf,
            "tables": [(t.displayName, t.ref) for t in tables],
            "merged_cells": merged
        })
    data["totals"] = {
        "sheets": len(wb.sheetnames),
        "hidden_sheets": len([s for s in wb.worksheets if s.sheet_state!='visible']),
        "very_hidden": len([s for s in wb.worksheets if s.sheet_state=='veryHidden']),
        "formulas": total_formulas,
        "validations": total_validations,
        "cf": total_cf,
        "tables": total_tables,
        "merged": sum(s["merged_cells"] for s in data["sheets"])
    }

    # named ranges
    named = []
    try:
        for name in wb.defined_names:
            dn = wb.defined_names[name]
            named.append({"name": dn.name, "value": str(dn.value)[:500], "attr_text": str(dn.attr_text)[:500]})
    except Exception as e:
        named.append({"error": str(e)})
    data["named_ranges"] = named
    data["named_ranges_count"] = len(named)

    wb.close()
    return data

# extract JSON schema summaries
def extract_json_schema(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            content = json.load(f)
        definitions = content.get('definitions', {})
        return {
            "file": path,
            "definitions_count": len(definitions),
            "definitions": list(definitions.keys())[:200],
            "top_properties": list(content.get('properties', {}).keys()),
            "size_bytes": os.path.getsize(path)
        }
    except Exception as e:
        return {"file": path, "error": str(e)}

# Collect all
report = {"files": []}
all_outputs = {}
for itrxls in files:
    print(f"\nProcessing {itrxls} ...")
    struct = extract_workbook_structure(itrxls)
    vba_info = get_vba_info(itrxls)
    # store
    struct["vba"] = vba_info
    struct["vba_modules_count"] = len(vba_info)
    struct["vba_total_lines"] = sum(m.get("lines",0) for m in vba_info if "lines" in m)
    report["files"].append(struct)
    # Save individual detailed JSON for later review
    base = pathlib.Path(itrxls).stem
    out_path = f"/tmp/{base}_extraction.json"
    with open(out_path, 'w') as out:
        # avoid huge formulas sample listing?
        json.dump(struct, out, indent=2)
    print(f"  sheets:{struct['totals']['sheets']} hidden:{struct['totals']['hidden_sheets']} formulas:{struct['totals']['formulas']} validations:{struct['totals']['validations']} vba_modules:{struct['vba_modules_count']} vba_lines:{struct['vba_total_lines']}")

# JSON schemas
json_files = [
 "ITR-1/ITR-1_2026_Main_V1.1.json",
 "ITR-2/ITR-2_2026_Main_V1.1.json",
 "ITR-3/ITR-3_2026_Main_V1.1.json",
 "ITR-4/ITR-4_2026_Main_V1.1.json",
]
json_summaries=[]
for jf in json_files:
    summ = extract_json_schema(jf)
    json_summaries.append(summ)
    print(f"{jf}: defs {summ.get('definitions_count')} ")

report["json_schemas"] = json_summaries

# Write master report
with open("/tmp/master_report.json", "w") as mf:
    json.dump(report, mf, indent=2)

print("\nDone master report written to /tmp/master_report.json")

# Also extract repo code status
repo_files = []
for root, dirs, files_list in os.walk("."):
    # skip .git
    if ".git" in root:
        continue
    # skip __pycache__
    for fname in files_list:
        fpath = os.path.join(root, fname)
        if ".git" in fpath:
            continue
        # only count code files
        if fname.endswith((".py", ".js", ".ts", ".tsx", ".java", ".go", ".rs", ".json", ".pdf", ".xlsm", ".csv")):
            repo_files.append(fpath)

print(f"\nRepo file count considered: {len(repo_files)}")
for rf in sorted(repo_files)[:100]:
    print(rf)
