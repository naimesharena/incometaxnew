#!/usr/bin/env python3
"""Stream-dump every cell (values, formulas, types) + data validations + defined names
of an .xlsm directly from OOXML. Low memory; resolves sharedStrings."""
import sys, re, json, zipfile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
RNS = NS["r"]

def col_to_idx(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n

def expand_sqref(sqref):
    out = []
    for token in (sqref or "").split():
        if ":" in token:
            a, b = token.split(":")
            ma, mb = re.match(r"([A-Z]+)(\d+)", a), re.match(r"([A-Z]+)(\d+)", b)
            c1, r1, c2, r2 = col_to_idx(ma.group(1)), int(ma.group(2)), col_to_idx(mb.group(1)), int(mb.group(2))
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    out.append(f"{c}:{r}")
        else:
            m = re.match(r"([A-Z]+)(\d+)", token)
            if m:
                out.append(f"{col_to_idx(m.group(1))}:{m.group(2)}")
    return out

def load_shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    ss = []
    for ev, el in ET.iterparse(z.open("xl/sharedStrings.xml"), events=("end",)):
        if el.tag == "{%s}si" % NS["m"]:
            ss.append("".join(t.text or "" for t in el.iter("{%s}t" % NS["m"])))
            el.clear()
    return ss

#: pure reference-data sheets: keep only first N rows of values, but keep ALL formula cells
DATA_SHEET_CAP = {"ISIN List": 5, "IFSC": 5, "DropDownValues": 5, "DB": 5, "DataBase": 5}

def dump(path, out_path):
    z = zipfile.ZipFile(path)
    ss = load_shared_strings(z)
    wbroot = ET.fromstring(z.read("xl/workbook.xml"))
    sheets = []
    for s in wbroot.find("m:sheets", NS):
        sheets.append({"name": s.get("name"), "state": s.get("state", "visible"),
                       "rId": s.get("{%s}id" % RNS)})
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2t = {rel.get("Id"): rel.get("Target") for rel in rels}

    defined = []
    dn = wbroot.find("m:definedNames", NS)
    if dn is not None:
        for d in dn:
            defined.append({"name": d.get("name"), "ref": (d.text or "").strip(),
                            "hidden": d.get("hidden", "0"), "local": d.get("localSheetId")})

    result = {"file": path, "defined_names": defined, "sheets": []}
    for idx, s in enumerate(sheets, 1):
        sheet_name = s["name"]
        target = rid2t.get(s["rId"], f"worksheets/sheet{idx}.xml")
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        cells = {}
        dvs = []
        maxrow = 0
        for ev, el in ET.iterparse(z.open(target), events=("end",)):
            tag = el.tag.split("}")[-1]
            if tag == "c":
                ref = el.get("r")
                t = el.get("t", "n")
                m = re.match(r"([A-Z]+)(\d+)", ref)
                key = f"{col_to_idx(m.group(1))}:{m.group(2)}"
                maxrow = max(maxrow, int(m.group(2)))
                entry = {"t": t}
                f = el.find("m:f", NS)
                if f is not None:
                    entry["f"] = f.text or ""
                    if f.get("t") == "shared" and f.get("si") is not None:
                        entry["fshared"] = f.get("si")
                v = el.find("m:v", NS)
                if v is not None:
                    if t == "s":
                        entry["v"] = ss[int(v.text)]
                    else:
                        entry["v"] = v.text
                else:
                    ist = el.find("m:is", NS)
                    if ist is not None:
                        entry["v"] = "".join(x.text or "" for x in ist.iter("{%s}t" % NS["m"]))
                        entry["t"] = "inlineStr"
                cap = DATA_SHEET_CAP.get(sheet_name)
                if cap is not None and int(m.group(2)) > cap and "f" not in entry:
                    el.clear()
                    continue
                if len(entry) > 1 or entry.get("t") != "n":
                    cells[key] = entry
                el.clear()
            elif tag == "dataValidation":
                dv = {"type": el.get("type"), "op": el.get("operator"),
                      "allowBlank": el.get("allowBlank"), "sqref": el.get("sqref"),
                      "showDropDown": el.get("showDropDown")}
                for fn in ("formula1", "formula2"):
                    fe = el.find("m:" + fn, NS)
                    if fe is not None and fe.text:
                        dv[fn] = fe.text
                dvs.append(dv)
                el.clear()
        for dv in dvs:
            dv["cells"] = expand_sqref(dv["sqref"])
        result["sheets"].append({"name": s["name"], "state": s["state"],
                                 "maxrow": maxrow, "cells": cells, "validations": dvs})
        print(f"  sheet '{s['name']}': {len(cells)} cells, {len(dvs)} DVs", file=sys.stderr)
    with open(out_path, "w") as fh:
        json.dump(result, fh)
    print(f"wrote {out_path}", file=sys.stderr)

if __name__ == "__main__":
    dump(sys.argv[1], sys.argv[2])
