"""ITR Filing Software - AY 2026-27.

Field- and logic-faithful reimplementation of the official CBDT Excel utilities
(ITR1_AY_26-27_V1.2.xlsm, ITR2 V1.3, ITR3 V1.2, ITR4 V1.1).

Every module documents the Excel/VBA source of each rule (sheet!cell or
module.procedure) so completeness can be audited against the extraction
registry (see docs/EXCEL_REVERSE_ENGINEERING_GAP_REPORT.md).
"""

__version__ = "0.1.0"
ASSESSMENT_YEAR = "2026"  # AY 2026-27
PREVIOUS_YEAR = "2025"    # FY 2025-26
