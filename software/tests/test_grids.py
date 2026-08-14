"""Grid input tests: 80G donations grid (ITR-1) and CG asset grid (ITR-2)."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "web"))

from datetime import date
from builder import build_itr1, build_itr2


BASE = {
    "first_name": "Test", "last_name": "User", "pan": "ABCPS1234K",
    "dob": "1986-08-14", "state": "11-Gujarat", "pincode": "395009",
    "email": "t@x.com", "mobile": "9876543210", "locality": "X", "city": "Surat",
    "addr1": "1", "ifsc": "SBIN0001234", "account_no": "1", "bank_name": "SBI",
    "verification_date": "2026-07-20", "regime": "old",
}


def test_80g_donations_grid():
    i = dict(BASE)
    i["gross_salary"] = 1600000
    i["donations_80g"] = [
        {"amount": 100000, "percent": "100", "qualifying": "no", "mode": "noncash"},
        {"amount": 50000, "percent": "50", "qualifying": "no", "mode": "noncash"},
        {"amount": 20000, "percent": "100", "qualifying": "no", "mode": "cash"},  # cash > 2k -> 0
    ]
    r = build_itr1(i)
    assert len(r.chapter_via.donations_80g) == 3
    s = r.compute_summary()
    # 100000 + 25000 (50% of 50k) + 0 = 125000
    assert s["ChapterVIAComponents"]["80G"] == 125000


def test_80g_qualifying_limit_grid():
    i = dict(BASE)
    i["gross_salary"] = 1600000
    i["donations_80g"] = [
        {"amount": 300000, "percent": "100", "qualifying": "yes", "mode": "noncash"},
    ]
    r = build_itr1(i)
    s = r.compute_summary()
    # qualifying limit = 10% of adjusted TI (< donation) -> capped
    assert s["ChapterVIAComponents"]["80G"] <= 160000
    assert s["ChapterVIAComponents"]["80G"] > 0


def test_cg_assets_grid():
    i = dict(BASE)
    i["cg_assets"] = [
        {"asset_class": "securities", "listed_stt": "yes",
         "acq_date": "2025-11-01", "transfer_date": "2026-02-01",
         "sale": 300000, "cost": 200000, "fmv": 0},                       # STCG20 +1L
        {"asset_class": "securities", "listed_stt": "yes",
         "acq_date": "2023-01-01", "transfer_date": "2026-02-01",
         "sale": 500000, "cost": 300000, "fmv": 0},                       # LTCG125 +2L
    ]
    r = build_itr2(i)
    assert len(r.cg_assets) == 2
    s = r.compute()
    assert s["cg_summary"]["STCG20"] == 100000
    assert s["cg_summary"]["LTCG125"] == 200000
    assert s["CapGainTotal"] == 300000


def test_cg_grid_grandfathering():
    i = dict(BASE)
    i["cg_assets"] = [
        {"asset_class": "securities", "listed_stt": "yes",
         "acq_date": "2017-01-01", "transfer_date": "2026-02-01",
         "sale": 500000, "cost": 100000, "fmv": 300000},
    ]
    r = build_itr2(i)
    s = r.compute()
    # cost basis = MAX(100000, MIN(300000, 500000)) = 300000 -> LTCG125 = 200000
    assert s["cg_summary"]["LTCG125"] == 200000
