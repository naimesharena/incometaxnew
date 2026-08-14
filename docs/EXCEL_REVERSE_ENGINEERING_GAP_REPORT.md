# ITR Filing Software — Excel Reverse-Engineering Verification & Gap Report

**Repo:** `naimesharena/incometaxnew` · **Branch:** `arena/019feeee-incometaxnew`
**Date:** 2026-08-11 · **Scope:** ITR-1, ITR-2, ITR-3, ITR-4 utilities for AY 2026-27
**Status of this document:** Pre-development verification report (no code has been written yet).

---

## 0. Executive summary (answers to the seven verification questions)

| # | Question | Verdict | Evidence |
|---|----------|---------|----------|
| 1 | Complete Excel structure extracted? | **Yes — verified** | All sheets, named ranges, formulas, data validations, conditional formats, tables, comments, controls and inter-sheet links parsed directly from the OOXML of all 4 workbooks (§2–§5) |
| 2 | Macros & VBA extracted? | **Yes — verified** | 465 VBA modules / **646,672 lines** decompiled with **zero extraction errors** via olevba (§6) |
| 3 | Hidden sheets? | **Yes — all found** | **100 hidden/veryHidden sheets** across the 4 books, each classified (logic / reference data / legacy) (§4) |
| 4 | Data & calculation logic captured? | **Yes** | 15,208 cell formulas, 14,706 defined names, 3,288 data-validation rules, 748 cell comments, tax engine sheets sampled and understood (§5) |
| 5 | Mapped to existing repository code? | **Not applicable — no code exists** | The repository contains **only** the 4 `.xlsm`, 4 JSON schemas, 5 PDFs and 13 CSVs. There are **0 source-code files** (§9) |
| 6 | Functional equivalence of current implementation? | **0 % — there is no implementation yet** | Nothing has been reproduced because nothing has been built. This repo is currently the *requirements source*, not the tool (§9) |
| 7 | Gap analysis | Provided | §10 feature-by-feature: everything is currently **Missing / To develop**, each mapped to its exact Excel evidence |

**Bottom line:** the `.xlsm` files are now fully reverse-engineered and inventoried (nothing hidden was missed — see §4 and §12 caveats). However, the repository contains **no existing software** to map against. If an "existing tool" exists in a different repository/location, point me to it and I will complete the mapping; otherwise the gap report below *is* the build backlog for the complete ITR filing software.

---

## 1. Verification methodology (how this was checked — reproducible)

All findings below were produced in this session by directly unpacking the OOXML containers and decompiling the VBA projects; none of it is assumed from file names:

| Step | Tool | Output location |
|------|------|-----------------|
| Unzip workbooks, parse `workbook.xml`, sheet XML, rels, external links | Python `zipfile` + ElementTree streaming | `/home/user/itr_extraction/structure/ITR{1..4}.json` |
| Sheet state (visible/hidden/**veryHidden**), formulas per sheet, Excel functions used, data validations (dropdowns), conditional formatting, tables, calc-chain | same | same |
| Full VBA decompilation of `xl/vbaProject.bin` | `oletools/olevba` | `/home/user/itr_extraction/vba/ITR{1..4}.vba` (28 MB total) |
| Module-level inventory (subs / functions / event handlers / line counts) | custom parser | `/home/user/itr_extraction/vba/*.inventory.json` |
| JSON submission schemas inspected | `json` | §7 |
| Validation-rule PDFs inspected | `pypdf` | §7 |
| Reference CSV templates inspected | `csv`/head | §7 |

Re-run any check with the scripts kept in `/home/user/itr_extraction/` (`analyze_xml.py`, `vba_inventory.py`).

---

## 2. Workbook-level inventory

| Item | ITR-1 `V1.2` | ITR-2 `V1.3` | ITR-3 `V1.2` | ITR-4 `V1.1` | **Total** |
|------|------|------|------|------|------|
| Worksheets | 21 | 66 | 66 | 24 | **177** |
| — visible | 5 | 25 | 40 | 7 | 77 |
| — hidden | 16 | 40 | 24 | 17 | 97 |
| — **veryHidden** | 0 | **1** (`OLDAL`) | **2** (`ITold`, `OLDAL`) | 0 | **3** |
| Cell formulas | 1,018 | 6,196 | 6,449 | 1,545 | **15,208** |
| Data-validation (dropdown/input) rules | 387 | 1,122 | 1,257 | 522 | **3,288** |
| Defined names (workbook + sheet scope) | 1,355 | 6,037 | 5,821 | 1,493 | **14,706** |
| Cell comments (developer notes) | 69 | 323 | 261 | 95 | **748** |
| Conditional-formatting blocks | 9 | 41+ | 29+ | 5 | ~85 |
| VML drawings / form-control bindings | 18 | 59 | 60 | 22 | 159 |
| ActiveX controls | 1 | 0 | 0 | 1 | 2 |
| VBA modules | 73 | 155 | 158 | 79 | **465** |
| VBA code lines | 73,500 | 236,221 | 240,157 | 96,794 | **646,672** |
| VBA Subs / Functions / Event handlers | 497 / 731 / 135 | 1,276 / 2,539 / 323 | 1,330 / 2,623 / 317 | 595 / 618 / 133 | **3,698 / 6,511 / 908** |

Visible navigation sheets: ITR-1 (`Income Details`, `HP`, `TDS`, `TCS`, `Taxes Paid and Verification`); ITR-2/ITR-3 (`Home` dashboard, `PART A - General`, `Schedule S`, `House Property`, `CG`, `Schedule 112A`, `Schedule 115AD(1)(iii) proviso`, `VDA`, `OS`, `CYLA - BFLA`, `CFL`, `VI-A`, `SPI - SI`, `AMTC`, `EI`, `PTI`, `FSI`, `Sch 5A`, `TR_FA`, `AL`, `Part B - TI TTI`, `IT`, `ESOP`, `TDS`, `Verification`); ITR-4 (`Income Details`, `HP`, `BP`, `TDS`, `TCS`, `IT`, `Taxes Paid and Verification`).

---

## 3. Form-control / interactivity layer

- ITR-1 carries **144 `ctrlProps` form-control bindings** plus 1 ActiveX OCX (`classid {D7053240-…}` — legacy MSForms control) and 17 drawing canvases: dropdowns, check-boxes and option buttons wired to VBA (`checkBoxModule`, `FilingSectRadioButton`, `ePayPrefill`).
- Dropdown list sources are both **inline** (e.g. `"(Select),Self Occupied, Let Out,Deemed Let Out"`), **named ranges** (`StateList`, `Selection80DB`, `ResStatus`) and **INDIRECT()-driven** (e.g. `INDIRECT(IF(MID(TRIM(sheet1.Status),1,1)="F","ResStatus_2","ResStatus"))` — residence-status options change with filing status).
- Inter-sheet wiring is done through defined names (e.g. ITR-4: `_44AD_Income → TaxCalc!$T$17`, `_Final_44AD → TaxCalc!$T$19`, `_NOB1 → DB!$CA$1:$CA$345`) — i.e. visible input sheets → hidden calc engine → visible output cells. All 14,706 names are captured in the structure JSONs.

---

## 4. Hidden & veryHidden sheets — complete classified list

Nothing was missed: visibility state was read from `workbook.xml` for every sheet (not from what Excel shows).

### 4.1 Hidden sheets containing **calculation logic** (must be re-implemented)
| Sheet | Workbook(s) | Content |
|---|---|---|
| `Tax Calculated` | ITR-2 (1,981 formulas), ITR-3 (2,019) | **The full tax computation engine**: normal + special-rate tax, 87A rebate, surcharge with marginal relief (`MarginalRelief` named formula), 4 % cess (`*(0.04)`), AMT, relief, interest 234A/B/C |
| `TaxCalc` | ITR-4 (710 formulas) | Same engine for presumptive-return filers (slabs, 87A by entity type `IS_FIRM/IS_HUF/IS_NRI`, net liability, interest) |
| `DataBase` | ITR-1 (361 formulas, 19,303 rows) | Code masters + lookup logic |
| `DropDownValues` | ITR-2/ITR-3 (≈103 formulas, 19,302 rows) | Dropdown master data + helper formulas |
| `DB` | ITR-4 (10 formulas, 63,101 rows) | Master data for ITR-4 |
| `SUMMARY` | all 4 | Roll-up used by Home/dashboard & verification |
| `Part B ATI`, `Part A Gen_139(8A)` | all 4 | Assessment-type & late-filing 139(8A) logic |
| `80C/80CCC/80CCD(80)/80D/80DD/80U/80E/80EE/80EEA/80EEB/80G/80GGA/80GGC/80U-80DD/10AA/RA` | ITR-1..4 (e.g. hidden `80G` has 144–196 formulas) | Chapter VI-A deduction math & caps (incl. 80G donee-type limit logic) |
| `AMT` | ITR-2/ITR-3 | Alternate Minimum Tax (115JC/115JD) worksheet |
| `FSI1` | ITR-2/ITR-3 | Foreign Source Income computation helper (DTAA) |
| `TPSA` | ITR-2 (hidden), ITR-3 (visible) | Advance-tax instalment schedule (234C driver) |
| `GST`, `BA`, `ICDS`, `10AA` | ITR-2/ITR-3 | Ancillary schedule logic |
| `Unabsorbed Depreciation` | ITR-2 (hidden) | Depreciation carry-forward |
| `BP`, `DPM - DOA`, `DEP_DCG`, `ESR`, `Part A - BS`, `Manufacturing/Trading Account`, `Profit and Loss`, `Part A - OI`, `Quantitative Details`, `Nature Of Business` | hidden in **ITR-2** (visible in ITR-3) | Full business-income machinery — present, just switched off for individuals without business income |
| `CG Pop up_prefill`, `Temporary Values`, `Sheet1`, `ITold`, `OLDAL` (veryHidden) | various | Prefill staging, scratch pads, **legacy** prior-year AL/IT layouts |

### 4.2 Hidden sheets that are **reference data** (must be migrated to the new system's databases)
| Sheet | Rows | Content |
|---|---|---|
| `ISIN List` | 88,568 (ITR-2) / 124,306 (ITR-3) | Security ISIN master for 112A/CG/115AD |
| `IFSC` | 45,138 (ITR-1) | Bank branch master (refund account lookup) |
| `BankCode` | 318 (ITR-1) | Bank name/code master |
| `Help` / `HelpCSV` / `Instructions` | up to 19k rows | User help content and instructions |

### 4.3 Conclusion for Q3
All 100 hidden/veryHidden sheets are extracted and their role identified. The three **veryHidden** sheets (`OLDAL` ×2, `ITold`) are legacy leftovers with negligible logic (1–64 formulas each), but they were still inspected — no business logic lives only there. All real hidden logic (§4.1) is mapped into the gap table (§10).

---

## 5. Formula & calculation logic (Q4)

- **15,208 formulas** captured with their text (functions frequency-analysed per sheet — dominated by `IF`, `ROUND`, `SUM`, `SUMIF`, `VLOOKUP`, `INDEX/MATCH`, `INDIRECT`, `OFFSET`, `MIN/MAX`, `CONCATENATE`, date math).
- Sampled and confirmed the **AY 2026-27 business rules** embedded in the engine, e.g.:
  - Slab-based tax with entity/status awareness (`IS_FIRM`, `IS_HUF`, `IS_NRI`, `GrpC`, taxable-income thresholds);
  - Rebate u/s **87A** (incl. special treatment around 115BAC — VBA refs: 169–691 hits per workbook for `87A/Surcharge/Cess/MarginalRelief`);
  - **Surcharge** tiers with **marginal relief** (`MarginalRelief` named formula, pro-rata relief `SurchargeOnTaxPayable*(x/BalTaxPayable)`);
  - **Health & Education Cess 4 %** (`(tax+surcharge)*0.04`);
  - **Interest u/s 234A/234B/234C** — dedicated VBA modules (`mdCalInterst234B` 707 lines; `mdInterestCalc`; 54–57 identifier hits per book) plus sheet-level period math (`234BOnPeriod`, `234Bprinciple`, `234Ci..iv`);
  - Special rates: **111A, 112, 112A, 115AD, 115BBA, 115BAC** (115AD alone: 2,413 VBA identifiers in ITR-2), indexed cost of acquisition via `DEP_DCG`;
  - **AMT 115JC/115JD** + credit carry-forward (`AMTC`);
  - **Set-off / carry-forward**: `CYLACalculations.bas` (8,836 lines, 168–214 functions), `BFLA_Calculations`, `mdCFL` (8-year CFL grid up to row 65,520), unabsorbed depreciation;
  - **DTAA relief**: FSI/FSI1, PTI, TR sheets + `SchTR_FA`; relief u/s 89/90/91;
  - House property: annuality, 30 % standard deduction, interest u/s 24(b), co-ownership, let-out/deemed-let (hidden `Schedule 24(b)`);
  - Salary: perquisites, ESOP (perquisite value + 89A relief refs: 383), VDA/Form 12BA;
  - Presumptive taxation **44AD/44ADA/44AE** (ITR-4 `mdNOBBP`, `md44AE`, named formulas `_44AD_Income` etc.);
  - Refund/balance payable with status (`Tax Refundable / Tax Payable / Nil Tax Balance` dropdown);
  - Verification, filing-section selection incl. late filing u/s 139(8A) (hidden `Part A Gen_139(8A)` + `mdGen139_8A`).
- **Validations:** 3,288 data-validation rules (list/whole/decimal/date/custom incl. `INDIRECT`-based dynamic lists) plus conditional-format highlighting and ~748 developer comments. These complement the official **Category A/B/C validation-rule PDFs** (§7) — both layers must be implemented.

---

## 6. VBA / macro layer (Q2) — fully extracted

- **465 modules, 646,672 lines, 3,698 Subs, 6,511 Functions, 908 event handlers** decompiled with olevba — **no extraction errors**, no obfuscation, no p-code-only stubs. Full source is in `/home/user/itr_extraction/vba/`.
- Third-party dependency noticed inside the code: **VBA-JSON** (github.com/VBA-tools/VBA-JSON) — used for JSON parse/serialize (only IOC/URLs flagged by olevba are this library and MS KB links).
- Key infrastructure modules (per workbook):

| Concern | ITR-1 | ITR-2 | ITR-3 | ITR-4 |
|---|---|---|---|---|
| JSON generation (official schema output) | `GenerateJson` 7,120 ln | `Generate_XML` 18,333 ln | `Generate_XML` (~18k) | `GenerateJSON` 8,256 ln |
| JSON/Excel import (draft load) | `ImportJson` 6,763 ln, `ImportExcel` | `mdImportXL` 3,636 ln | `mdImportXL` | `ImportJSON` 9,638 ln |
| Govt prefill import | `PreFillJson` 4,430 ln, `mdImportXML` | `mdImportXML` | `mdImportXML` | `PreFillJson` 6,233 ln |
| Hashing/signing | `mdHashing` + `HS256.cls` (526 ln) | `mdHashing` | `mdHashing` | `mdHashing` |
| Interest 234 engine | `mdCalInterst234B`, `mdInt24b` | `mdInterestCalc` | `mdInterestCalc` | in `mdTaxCalc` |
| Sheet protection | `PWD` | `pwd` | `pwd` | `pwd` |
| UX | progress bar, MessageBox forms, `mfMessage` | same | same | same |
| Schedule modules | SchHP, Sch80G/D/GGA/GGC/U_DD/E/EE/EEA/EEB/C, SchBA, SchDI, SchAL, ModuleEA10_13A, AY23_24Changes | Part_A_General (8,080 ln), SchCG (9,175), SchTR_FA (5,678), SchDPM_DOA (311 fns), CGDeductions, md112A, md115AD, ImportSchedule112A/115AD, Tax_Calc, mdAMT/AMTC, mdEI, mdPTI, SchFSI, Sch5A, SchRA, SchUD, GST, Verification | + BP/P&L/Manufacturing/Trading modules, ICDS, Quantitative_Details | mdIncomeDetails, mdNOBBP, md44AE, SchBA, mdTDS (4,377 ln) + mdTDS2/3, mdTCS, mdIT, md80G, Sch80D, mdAL, mdPAN |
| Event surface | Worksheet Change/Activate/SelectionChange handlers on every input sheet drive recalculation + dependent dropdown refresh (`Worksheet_Change` in 90+ sheet classes) | | | |

**Event handlers matter:** sheet-level `Worksheet_Change/Activate` code enforces field dependencies, re-computes hidden engines and refreshes validations — this workflow behaviour is part of the spec and must be reproduced in the new UI.

---

## 7. Companion CBDT artefacts in the repo (all inspected)

| File | Nature | Verified content |
|---|---|---|
| `ITR-{1..4}_2026_Main_V1.1.json` | Official **JSON Schema (draft-04)** of the e-filing payload | ITR-1: 64 definitions (`Form_ITR1`, `PersonalInfo`, `AddressDetail…`, `ITR1_IncomeDeductions`, `ITR1_TaxComputation`, `Schedule80C/D/DD/E/EE/EEA/EEB/G`, `LTCG112A`, `BankAccountDtls`, `Refund`, `IntrstPay`, `FilingStatus`, `CreationInfo`…). ITR-2/3/4 analogous but much larger (390 KB / 1.06 MB / 252 KB). **This is the exact output contract our software must generate.** |
| `CBDT_e-Filing_ITR {1,2,4}_Validation Rules` + ITR-3 rules | Rulebooks | 22 / 51 / 73 / 24 pages, Categories A (blocking), B, C — must be implemented as server/client validation |
| `ITR 1_Schema change document_AY2026-27_V1.1.pdf` | Delta vs previous AY | 5 pages — tracks field changes |
| `CSV_TDS1/2/3.csv`, `CSV_TCS.csv`, `CSV_IT.csv`, `CSV_112A.csv`, `CSV_115AD.csv` (ITR-2 & ITR-3 folders) | Bulk-import templates | Column layouts for TDS (salary / non-salary / ESOP), TCS, challan IT, 112A and 115AD imports — mirrored by the workbook importers (`ImportSchedule112A`, `mdTDS*`, `ImportExcel`) |

---

## 8. Functional decomposition of the Excel system (the "spec" the new software must meet)

1. **Return selection & eligibility** (ITR-1..4 chooser; ITR-2/3 `Home` dashboard with status-driven sheet visibility — 38/28 conditional-format blocks).
2. **Part A – General**: PAN, Aadhaar linkage, residential status, address, filing section 139(1)/139(4)/139(8A), porting u/s 139(9), filing-status-driven option sets.
3. **Income capture**: Salary (+ESOP/VDA/12BA), House Property (incl. co-owned, 24(b)), Capital Gains (incl. 112A/115AD(1)(iii), CII indexation, exemptions 54/54EC…), Other Sources (dividends, interest, VDA, gambling), Business/Profession (P&L, BS, Mfg/Trading, Quantitative, ICDS, GST, depreciation DPM/DOA, presumptive 44AD/44ADA/44AE).
4. **Deductions**: full Chapter VI-A set + 10AA + EI exempt-income schedules.
5. **Aggregation**: CYLA/BFLA set-off, CFL carry-forward (8 AYs), unabsorbed depreciation.
6. **Tax computation engine**: regime choice (115BAC), slabs, 87A, surcharge tiers + marginal relief, cess, special rates, AMT + credit, reliefs 89/90/91/DTAA, interest 234A/B/C, advance-tax schedule, net payable/refund.
7. **Taxes paid**: TDS (3 kinds), TCS, challan IT, ESOP TDS — with CSV/prefill import & TAN/BSR validation.
8. **Auxiliary schedules**: TR/FA (foreign assets & DTAA), AL (assets & liabilities), Sch 5A, BA, DI, GST, 139(8A).
9. **Refund & bank**: bank account selection against 45k-row IFSC master, refund re-issue details.
10. **Data exchange**: generate CBDT JSON per schema; import previous draft JSON; import govt prefill JSON/XML; hash verification (HS256).
11. **Validation layer**: 3,288 DV rules + conditional formats + VBA messages + official Category A/B/C PDF rules.
12. **UX behaviours**: dependent dropdowns (`INDIRECT`), progressive disclosure of hidden sheets, progress bars, help/instructions content, sheet protection.

---

## 9. Mapping against the existing repository (Q5 & Q6)

A complete file census of the repo (excluding `.git`):

```
13 csv · 5 pdf · 4 xlsm · 4 json   ← that is the entire repository
```

**There is no application code whatsoever** (no source files of any language, no build config, no tests, no CI). Therefore:

- Q5 (mapping to existing tool): there is nothing in this repository to map to. *If your "existing tool" lives in another repository or system, share it and I will extend this report with a line-by-line mapping.*
- Q6 (functional equivalence): **0 %** — nothing has been implemented yet. The repo currently holds the **complete requirements and reference data**; the software must be built from scratch on top of them.

---

## 10. Gap analysis (Q7)

Legend: ✅ mapped/implemented · 🟡 partial · ❌ missing. Since the repo has no code, **every row is ❌ today** — the table's value is defining the exact build backlog with Excel evidence.

| # | Excel functionality (evidence) | Where in Excel | Repo status | To develop |
|---|---|---|---|---|
| 1 | Form chooser + eligibility rules | Home dashboards, conditional formats, `Part_A_General.bas` | ❌ | ITR-eligibility wizard |
| 2 | Personal / Part-A General data model | PART A - General (87–151 DV rules), `Part_A_General.bas` (8,080+ ln) | ❌ | Identity module |
| 3 | Salary incl. ESOP/VDA/perquisites | Schedule S, VDA, ESOP sheets; `TI_TTI_Salary.bas` | ❌ | Salary module |
| 4 | House property incl. 24(b), co-ownership | HP + hidden `Schedule 24(b)`; `SchHP.bas` | ❌ | HP module |
| 5 | Capital gains + 112A/115AD + indexation | CG (689–757 formulas), Schedule 112A/115AD, `SchCG.bas`, `CGDeductions`, `DEP_DCG`, ISIN List | ❌ | CG module + ISIN DB |
| 6 | Other sources incl. VDA income | OS (331–385 formulas), `SchOS.bas` | ❌ | OS module |
| 7 | Business & profession full suite (ITR-3) / presumptive (ITR-4) | BP, P&L, BS, Mfg/Trading, Quantitative, DPM-DOA, ICDS, GST sheets; 44AD/ADA/AE logic (`mdNOBBP`, `md44AE`, `_44AD_Income`) | ❌ | BP module (biggest workstream) |
| 8 | Chapter VI-A deductions | Hidden 80x sheets + VI-A; `md80_.bas`, `Sch80G/D/...` | ❌ | Deductions module |
| 9 | Exempt income (10/10A/10AA) | EI, `10AA`, `mdEI`, `Sch10A` | ❌ | EI module |
| 10 | Set-off CYLA/BFLA + CFL carry-forward | CYLA - BFLA, CFL (65,520-row grid), `CYLACalculations.bas` (168–214 fns), `mdCFL` | ❌ | Set-off/CFL engine |
| 11 | Tax engine: slabs, 87A, surcharge + marginal relief, cess | Hidden `Tax Calculated` / `TaxCalc` (710–2,019 formulas), `Tax_Calc.bas` | ❌ | Tax computation engine |
| 12 | Special rates 111A/112/112A/115AD/115BBA/115BAC | SPI - SI sheets (784–927 formulas), `md112A`, `md115AD` | ❌ | Special-rate engine |
| 13 | AMT 115JC/JD + credit | AMT (hidden), AMTC, `mdAMT/AMTC` | ❌ | AMT module |
| 14 | Interest 234A/B/C | `mdCalInterst234B`, `mdInterestCalc`, engine sheets | ❌ | Interest engine |
| 15 | Relief 89/90/91 + DTAA (FSI/PTI/TR) | FSI, FSI1, PTI, TR_FA, Sch 5A, `SchFSI`, `mdPTI` | ❌ | Relief/foreign module |
| 16 | Taxes paid: TDS/TCS/IT + CSV importers | TDS/TCS/IT sheets; `mdTDS*`, `SchIT`, CSV templates | ❌ | Taxes-paid module + CSV import |
| 17 | AL / TR-FA / BA / DI / GST schedules | AL, TR_FA, BA, SchDI, GST | ❌ | Auxiliary schedules |
| 18 | Refund & bank (IFSC 45k rows, BankCode) | Hidden `IFSC`, `BankCode`; `SchTaxVerify` | ❌ | Refund module + IFSC DB |
| 19 | CBDT JSON generation (schema contract) | `GenerateJson`/`Generate_XML` + `*_Main_V1.1.json` schemas | ❌ | JSON builder per schema |
| 20 | Draft save/load (JSON import) | `ImportJson`/`mdImportXL` | ❌ | Persistence layer |
| 21 | Govt prefill import (JSON/XML) | `PreFillJson`, `mdImportXML`, `ePayPrefill` | ❌ | Prefill integration |
| 22 | Hash/integrity (HS256) | `HS256.cls`, `mdHashing` | ❌ | Integrity checks |
| 23 | Validation suite (3,288 DV + PDF Category A/B/C) | DV rules, CF blocks, validation PDFs (22/51/73/24 pp) | ❌ | Validation engine (blocking + warnings) |
| 24 | UX: dependent dropdowns, progressive disclosure, help | `INDIRECT` lists, hidden-sheet toggling, Help/Instructions | ❌ | Web UI workflows |
| 25 | Reference databases | ISIN List (88k/124k), DropDownValues (19k), DB (63k), DataBase (19k) | ❌ | Data migration into DB |

**Partial mappings: none. Successful mappings: none — by construction (no code exists).**

---

## 11. Residual risks / caveats (stated for completeness)

1. **Cached values** of formulas were not needed for logic extraction (formula text + VBA fully define behaviour), but we have not *executed* the workbook (no Excel in this environment); numeric golden outputs should be verified against the official utility once we build the engine (test-vectors strategy, §12).
2. The **external links** found (ITR-1/2/4) are *vestigial developer references* to old workbook versions on local drives/SharePoint — verified dead, no runtime dependency.
3. Two ActiveX OCX controls (ITR-1/4) are legacy MSForms; their behaviour (date/popup pickers) is specified by the surrounding VBA, not by the binary.
4. Worksheets are password-protected (`PWD.bas`); protection is a UX guardrail, not logic — decide the equivalent in the new product.
5. The official **prefill API contract** is not in the repo (only its consumer code); integration must follow the e-filing portal specs.
6. VBA identifier `AY23_24Changes` shows accumulated legacy patches; AY 2026-27 behaviour must be taken from current-year cells/PDFs, which is what we did.

---

## 11b. Development progress update (2026-08-14, post-confirmation)

Development has started on the confirmed extraction. Status of the §10 backlog:

| Area | Status | Where |
|---|---|---|
| Complete cell-level field registry (inputs + internal fields) | ✅ | `itr_extraction/registry/` (ITR-1: **2,554 input cells, 1,018 internal formula fields, 1,355 defined names** captured) |
| Full VBA decompilation archive in repo | ✅ | `docs/extraction/vba_and_structure.tar.gz` |
| ITR-1 tax engine (old-regime age slabs + AY 26-27 new-regime slabs) | ✅ | `software/itr_filing/tax_engine.py` (ported from `calcTaxPayableOnTI` / `calcTaxPayableOnTINTR`) |
| Rebate 87A incl. marginal relief (12L/60k new, 5L/12.5k old) | ✅ | `tax_engine.rebate_87a` (port of `Income Details!AO177`) |
| Health & Education Cess 4 % | ✅ | `tax_engine.education_cess` (`AO180`) |
| Chapter VI-A caps (80C/CCC/CCD(1)/(1B)/(2)/CCG/D/DD/DDB/E/EE/EEA/EEB/G/GG/GGA/GGC/QQB/RRB/TTA/TTB/U/CCH) | ✅ | `chapter_via.py` (port of `AN115..AN163` incl. AY-26-27 quirks: 80GG zeroed, new-regime blocks) |
| House Property (2 properties, 30 %, 24(b) regime rules, −2L loss cap, GTI flooring) | ✅ | `house_property.py` (port of HP sheet + `AO80..AO84`) |
| Salary pipeline (17(1)/(2)/(3), §10 exemptions, 16(ia) 75k/50k, 16(ii)/(iii)) | ✅ | `itr1.py` (port of `AO54..AO76`) |
| Other Sources (dividend quarters, 57(iia), 89A retirement accounts) | ✅ | `itr1.py` (port of `AO85..AO110`) |
| GTI / GTI incl. 112A / Total Income rounded to ₹10 | ✅ | `itr1.py` (`AO111/112/164/165`) |
| 112A LTCG not-chargeable + ITR-1 ₹1.25L eligibility cap | ✅ | `tax_engine.ltcg_112a_not_chargeable` (`AO175`) + `validation.py` |
| Interest 234A / 234B (10k threshold, 90 % rule, floor-to-100, month count, senior exemption) | ✅ | `interest.py` (port of `mdCalInterst234B`) |
| Fee 234F (₹1,000/₹5,000) and 234-I (revised return) | ✅ | `interest.py` (VBA fee blocks) |
| CBDT JSON output for ITR-1 | ✅ **100 % schema-valid** | `json_builder.py` — verified against official `ITR-1_2026_Main_V1.1.json` (0 errors) incl. TAN jurisdiction codes, state codes, ISO dates, Digest b64-SHA256 |
| Validation rules (Category A subset: identity, Aadhaar, PIN, mobile, TAN/IFSC formats, 112A cap, filing-section linkage, bank mandatory) | ✅ started | `validation.py` |
| Golden-vector test suite | ✅ 30 tests | `software/tests/` |
| Schedule EA 10(13A) HRA (regime-aware), 80D AY 26-27 selection model (25k/50k/75k/1L), 80DD/80U (75k/125k), 80DDB (40k/1L), 80G qualifying-limit engine, 234C exact port, 16(ii) cap, agri-income eligibility rule | ✅ | `software/itr_filing/schedules.py`, `interest.interest_234c` (43 tests) |
| Remaining for ITR-1: Form 10E relief-89 worksheet, prefill import, 80QQB/80RRB detail sheets, Part B ATI refund-interest (244A) grid | ⏳ next | |
| ITR-4 engine core: entity-aware slabs (TaxCalc GrpA/B/C/HUF/Firm), surcharge 10/15/12% + marginal relief, cess post-rebate, 87A individuals-only, presumptive 44AD (6%/8%, 3Cr cap), 44ADA (50%, 75L), 44AE (₹1,000/MT & ₹7,500/month), 50L eligibility | ✅ `software/itr_filing/itr4.py` (10 tests) |
| ITR-4 JSON builder (Form 10IEA flags, ScheduleBP 44AD/ADA/AE detail, due date 31/08/2026) - 100 % schema-valid | ✅ `software/itr_filing/itr4_json.py` |
| ITR-2/3 CYLA current-year set-off engine: 17 income buckets incl. AY 26-27 CG rates (STCG 15/20/30, LTCG 10/12.5/20/DTAA), HP loss 2L cap + set-off order, BP/OS loss orders, new-regime HP-loss rule | ✅ `software/itr_filing/setoff.py` (8 tests) |
| ITR-2/3 BFLA brought-forward set-off (HP/BP/speculation/specified/STCG/LTCG/racehorse orders, unabsorbed depreciation any-head-except-salary) + CFL tracker (8-AY HP/BP/CG, 4-AY speculation/race-horse, indefinite depreciation & 35AD) decoded from CFL grid rows 6-25 | ✅ `carry_forward.py` (11 tests) |
| CG engine: 112A grandfathering MAX(cost, MIN(FMV 31/01/2018, sale)), unified holding periods (12 mo securities / 24 mo other, post 23/07/2024), rate bands (STCG 20/30/applicable, LTCG 12.5/20), IHLA intra-CG waterfall set-off decoded from CG!rows 498-523 | ✅ `capital_gains.py` (9 tests) |
| CG exemptions 54 (two-house Rs 2Cr option), 54EC (Rs 50L cap [VBA 31622]), 54F (pro-rata, Rs 10Cr cap [VBA 31637]) | ✅ `cg_exemptions.py` |
| AMT 115JC: adjusted TI = TI + VI-A(excl QQB/RRB) + 10AA + 35AD; 18.5% + surcharge only if adj TI > 20L & VI-A claimed; N/A new regime. AMTC 115JD: 15-AY credit grid, zeroed in new regime | ✅ `amt.py` (7 tests) |
| DTAA relief u/s 90 (MIN foreign tax vs Indian incremental tax) & u/s 91 (lower-rate method) | ✅ `dtaa.py` (2 tests) |
| Part B-TI/TTI capstone: special-rate band taxes (111A 20%, 112A 12.5% with 1.25L exempt, 30%/10% legacy bands), 87A tested on TI excl. 112A [AO177 semantics], surcharge new capped 25% vs old 37%, cess, MAX(regular, AMT) rule | ✅ `ti_tti.py` (7 tests) — 105 tests total |
| ITR-2 return model (heads + CG waterfall + CYLA/BFLA/CFL + VI-A + TI/TTI + interest + refund) and JSON builder; introduced schema-autofill technique (recursive required-leaf skeleton from the draft-04 schema, then real-value overlay) | ✅ `itr2.py` + `itr2_json.py` — **100 % schema-valid** (5 tests, 110 total) |
| ITR-3 business engine: P&L net profit + inadmissible add-backs + IT-rules depreciation (block-wise, 15%/full/half for <180-day additions) [PARTA_PL + DPM sheet rules], speculation & 35AD tracks isolated per CYLA/BFLA rules | ✅ `business.py` (4 tests) |
| ITR-3 model + JSON builder (PartA_GEN2 audit flags, PARTA_BS/PL, ITR3ScheduleBP adjustment chain incl. DepreciationAllowITAct32/NetPLAftAdjBusOthThanSpec/NetPLBusOthThanSpec7A7B7C, ProfBusGain, ITR-3 due-date enum 31/08/2026) | ✅ **100 % schema-valid** (4 tests) — **all 4 forms complete, 116 tests** |
| Web application: FastAPI + single-page filing UI for all 4 forms — spec-driven dynamic fields, live tax computation with full breakdown, Category-A validation messages, CBDT JSON generation with in-request schema validation (all 4 forms verified 0 errors through the API) | ✅ `software/web/` |
| Schedule 115AD(1)(iii) proviso detail, e-filing portal prefill integration, draft save/load | ⏳ backlog (§10) | |

Demo: `software/run_demo.py` computes a 14 LPA new-regime return (tax 78,750 + cess 3,150 = 81,900; balance payable 22,776 after TDS 60,000) and emits a **schema-valid** CBDT JSON.

---

## 12. Recommended build plan (post-confirmation)

- **Phase 0 (done):** this extraction & gap report.
- **Phase 1:** Core platform — data model from JSON schemas, validation-rule engine (Category A/B/C), reference DBs (IFSC, banks, ISIN, dropdown masters).
- **Phase 2:** Tax engines — computation (slabs/87A/surcharge/marginal relief/cess/special rates/AMT), set-off/CFL, interest 234A/B/C, reliefs. Golden test vectors: replicate known scenarios and cross-check with the official utility.
- **Phase 3:** Income modules per form — ITR-1 → ITR-4 → ITR-2 → ITR-3 (increasing complexity), including all auxiliary schedules.
- **Phase 4:** I/O — CBDT JSON generation (schema-validated), draft persistence, CSV importers, prefill import, hashing.
- **Phase 5:** UX — eligibility wizard, dependent-dropdown workflows, progressive disclosure, help content, validation messaging.

---

## Appendix — extraction artefacts (outside Git, in workspace)

- `/home/user/itr_extraction/structure/ITR{1..4}.json` — per-workbook structural inventory (sheets incl. hidden state, formulas, functions, validations, defined names, external links)
- `/home/user/itr_extraction/vba/ITR{1..4}.vba` — full decompiled VBA source (~28 MB)
- `/home/user/itr_extraction/vba/*.inventory.json` — module/procedure inventories
- Scripts: `analyze_xml.py`, `analyze_xlsm.py`, `vba_inventory.py` (reproducible re-runs)
