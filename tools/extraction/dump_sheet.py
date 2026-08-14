#!/usr/bin/env python3
"""Dump a named sheet from a cells.json with names + formulas + labels."""
import json, re, sys

path, sheet_name = sys.argv[1], sys.argv[2]
maxr = int(sys.argv[3]) if len(sys.argv) > 3 else 60
d = json.load(open(path))
s = next(x for x in d["sheets"] if x["name"] == sheet_name)
names = {}
for n in d["defined_names"]:
    m = re.match(r"'?" + re.escape(sheet_name) + r"'?!\$?([A-Z]+)\$?(\d+)(?::|$)", n["ref"])
    if m:
        x = 0
        for ch in m.group(1):
            x = x * 26 + ord(ch) - 64
        names.setdefault((int(m.group(2)), x), []).append(n["name"])
cells = s["cells"]
rows = {}
for k, e in cells.items():
    c, r = map(int, k.split(":"))
    if r <= maxr:
        rows.setdefault(r, {})[c] = e
for r in sorted(rows):
    parts = []
    for c in sorted(rows[r]):
        e = rows[r][c]
        nm = ",".join(names.get((r, c), []))
        if e.get("f"):
            parts.append(f"[c{c}|{nm}] F:{e['f'][:150]}")
        elif e.get("t") in ("s", "inlineStr") and str(e.get("v", "")).strip():
            parts.append(f"[c{c}|{nm}] {str(e['v']).replace(chr(10), ' ')[:60]}")
    if parts:
        print(f"R{r:>3}: " + " || ".join(parts)[:450])
