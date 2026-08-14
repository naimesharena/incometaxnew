# ITR Filing Software — AY 2026-27

Field- and logic-faithful reimplementation of the official CBDT Excel utilities
(`ITR1_AY_26-27_V1.2.xlsm`, `ITR2_AY_26-27_V1.3.xlsm`, `ITR3_…V1.2.xlsm`,
`ITR4_…V1.1.xlsm`) that ship with this repository.

Every rule in the code cites its Excel/VBA source (sheet!cell or
module.procedure) so completeness is auditable against the reverse-engineering
report: [`../docs/EXCEL_REVERSE_ENGINEERING_GAP_REPORT.md`](../docs/EXCEL_REVERSE_ENGINEERING_GAP_REPORT.md).

## Status (this milestone)

| Workstream | Status |
|---|---|
| Full reverse-engineering of the 4 `.xlsm` (sheets, hidden sheets, formulas, defined names, validations, VBA) | ✅ done (see report + extraction archive) |
| **ITR-1 core engine** (slabs, 87A + marginal relief, cess, Chapter VI-A caps, HP, OS, GTI/TI rounding, 234A/B, 234F/234-I) | ✅ implemented + 29 golden tests |
| **CBDT JSON output for ITR-1** | ✅ **100 % valid against the official draft-04 schema** (`ITR-1_2026_Main_V1.1.json`) |
| Validation rules (Category A subset) | ✅ started (identity, eligibility, 112A cap, TAN/IFSC, filing-section, agri > 5k) |
| Schedule EA 10(13A) HRA exemption (regime-aware, metro 50%/40%) | ✅ `schedules.hra_exemption_10_13a` |
| Schedule 80D AY 26-27 selection model (7 options, caps 25k/50k/75k/1L, medical 50k/1L, check-up 5k) | ✅ `schedules.Schedule80D` |
| 80DD / 80U (75k/125k), 80DDB caps (40k/1L) | ✅ `schedules.amount_80dd/80u/80ddb` |
| Schedule 80G (100%/50%, qualifying-limit 10% of adjusted TI, cash > 2k rule) | ✅ `schedules.eligible_donations_80g` |
| Interest 234C (exact port: 12%/36% tolerances, 15/45/75/100% instalments, 3-3-3-1-1 months) | ✅ `interest.interest_234c` |
| Entertainment allowance 16(ii) cap (least of actual/5k/1/5 salary) | ✅ `schedules.cap_entertainment_allowance_16ii` |
| **ITR-4 engine core** (Individual/HUF/Firm slabs, surcharge + marginal relief, cess, 44AD 6%/8% & 3Cr cap, 44ADA 50% & 75L cap, 44AE 1000/MT & 7500/month, 50L eligibility, 87A, 234A/B/F) | ✅ implemented + 10 tests |
| ITR-4 JSON output against official schema | ✅ **100 % schema-valid** (`itr4_json.py`) |
| ITR-2/3 CYLA set-off engine (17 income buckets, HP 2L cap, loss-order matrix, regime rules) | ✅ `setoff.py` (8 tests) |
| ITR-2/3 BFLA (brought-forward set-off incl. depreciation) + CFL tracker (8/4-AY expiry, indefinite depreciation/35AD) | ✅ `carry_forward.py` (11 tests) |
| CG engine: 112A grandfathering (31/01/2018 FMV), unified holding periods (12/24 mo), rate bands, IHLA intra-CG waterfall set-off | ✅ `capital_gains.py` (9 tests) |
| CG exemptions: 54 (2-house 2Cr rule), 54EC (50L cap), 54F (pro-rata, 10Cr cap) | ✅ `cg_exemptions.py` |
| AMT 115JC (18.5%, 20L threshold, N/A in new regime) + AMTC 115JD 15-AY credit grid | ✅ `amt.py` |
| DTAA relief u/s 90/91 (lower-of-two / lower-rate methods) | ✅ `dtaa.py` |
| Part B-TI/TTI capstone: special-rate band taxes (20/30/12.5/10%), 87A on TI excl. 112A, regime-aware surcharge (new capped 25%, old 37%), cess, MAX(regular, AMT) | ✅ `ti_tti.py` |
| ITR-2/3 JSON builders, Schedule 115AD(1)(iii), web UI | ⏳ next |
| ITR-2 / ITR-3 (CG, BP, P&L, CYLA/CFL, AMT, DTAA, SPI/SI) | ⏳ planned |
| Web UI (form entry + live validation) | ⏳ planned |

## Layout

```
itr_filing/
  constants.py      AY 2026-27 constants, caps, code lists (each sourced)
  tax_engine.py     slab tax (old regime by age / new regime), 87A, cess, 112A
  chapter_via.py    all Chapter VI-A caps (Income Details!AN115..AN161 port)
  house_property.py HP sheet computation (2 properties, loss cap, regime rules)
  interest.py       234A / 234B / 234F / 234-I (mdCalInterst234B port)
  itr1.py           complete ITR-1 return model + pipeline (IncD.* fields)
  json_builder.py   CBDT JSON payload builder (schema-conformant)
  validation.py     Category-A validation rules
tests/
  test_engine.py         golden vectors for every engine rule
  test_itr1_pipeline.py  end-to-end pipeline + STRICT schema validation
```

## Quick start

```bash
python3 -m pytest tests/ -q          # 29 tests
python3 run_demo.py                  # demo return -> computation + CBDT JSON
```

## Fidelity notes (deliberate ports of utility quirks)

* §80GG is **zeroed** for AY 2026-27 because the utility formula literally
  takes `MIN(...,0)` (`Income Details!AN152`).
* 87A under the **new** regime is tested on Total Income *excluding* 112A
  LTCG; under the **old** regime on Total Income *including* it (`AO177`).
* Total Income is rounded to the nearest ₹10 (`ROUND(MAX(0,x),-1)`, `AO164`).
* New regime: self-occupied HP interest deduction = 0; HP loss not set off
  in GTI (`AO82`, `AO111`).
* Standard deduction ₹75,000 (new) / ₹50,000 (old) (`AO73`).
* 80DD / 80U / 80GGA / 80GGC are unavailable under the new regime
  (`AN144/154/155/160`).
* ITR-1 eligibility caps LTCG u/s 112A at ₹1,25,000 (VBA validation message).
