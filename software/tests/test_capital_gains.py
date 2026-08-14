"""Capital gains engine tests: 112A grandfathering, holding periods, bands,
and the IHLA waterfall set-off."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import date
from itr_filing.capital_gains import (CapitalAsset, grandfathered_cost_112a,
                                      cg_current_year_setoff, compute_cg_summary)


# ---------------------------------------------------------------- 112A
def test_grandfathering_fmv_between_cost_and_sale():
    assert grandfathered_cost_112a(100000, 200000, 300000) == 200000


def test_grandfathering_fmv_above_sale():
    # gain nil - cost capped at sale consideration
    assert grandfathered_cost_112a(100000, 350000, 300000) == 300000


def test_grandfathering_fmv_below_cost():
    assert grandfathered_cost_112a(100000, 50000, 300000) == 100000


# ---------------------------------------------------------------- holding periods
def test_holding_period_securities():
    lt = CapitalAsset(asset_class="securities", listed_stt_paid=True,
                      acquisition_date=date(2025, 1, 10), transfer_date=date(2026, 2, 15))
    st = CapitalAsset(asset_class="securities", listed_stt_paid=True,
                      acquisition_date=date(2025, 5, 10), transfer_date=date(2026, 2, 15))
    assert lt.is_long_term()      # 13 months > 12
    assert not st.is_long_term()  # 9 months


def test_holding_period_immovable():
    lt = CapitalAsset(asset_class="immovable",
                      acquisition_date=date(2024, 1, 1), transfer_date=date(2026, 2, 15))
    st = CapitalAsset(asset_class="immovable",
                      acquisition_date=date(2024, 5, 1), transfer_date=date(2026, 2, 15))
    assert lt.is_long_term()      # 25 months > 24
    assert not st.is_long_term()  # 21 months


# ---------------------------------------------------------------- bands
def test_band_allocation():
    assert CapitalAsset(asset_class="securities", listed_stt_paid=True,
                        acquisition_date=date(2023, 1, 1),
                        transfer_date=date(2026, 3, 1)).band() == "LTCG125"
    assert CapitalAsset(asset_class="securities", listed_stt_paid=True,
                        acquisition_date=date(2025, 10, 1),
                        transfer_date=date(2026, 3, 1)).band() == "STCG20"
    assert CapitalAsset(asset_class="immovable",
                        acquisition_date=date(2023, 1, 1),
                        transfer_date=date(2026, 3, 1)).band() == "LTCG125"
    assert CapitalAsset(asset_class="other",
                        acquisition_date=date(2025, 10, 1),
                        transfer_date=date(2026, 3, 1)).band() == "STCG_RATE"


def test_asset_gain_and_exemptions():
    a = CapitalAsset(asset_class="immovable",
                     acquisition_date=date(2020, 1, 1), transfer_date=date(2026, 3, 1),
                     sale_consideration=5000000, cost_of_acquisition=2000000,
                     transfer_expenses=100000, exemptions={"54": 1000000})
    assert a.gain() == 2900000
    assert a.gain_after_exemptions() == 1900000


# ---------------------------------------------------------------- waterfall
def test_stcl_flows_stcg_then_ltcg():
    out = cg_current_year_setoff({"STCG20": 100000, "LTCG125": 50000,
                                  "STCG30": -120000})
    assert out["STCG20"] == 0
    assert out["LTCG125"] == 30000     # 20k of the loss reached LTCG
    assert out["STCG30"] == 0


def test_ltcl_only_ltcg():
    out = cg_current_year_setoff({"STCG20": 100000, "LTCG125": 50000,
                                  "LTCG20": -80000})
    assert out["STCG20"] == 100000      # untouched by LTCL
    assert out["LTCG125"] == 0
    assert out["LTCG20"] == -30000      # remaining -> CFL


def test_intra_band_netting_then_waterfall():
    a1 = CapitalAsset(asset_class="securities", listed_stt_paid=True,
                      acquisition_date=date(2025, 11, 1), transfer_date=date(2026, 3, 1),
                      sale_consideration=150000, cost_of_acquisition=100000)
    a2 = CapitalAsset(asset_class="securities", listed_stt_paid=True,
                      acquisition_date=date(2025, 11, 1), transfer_date=date(2026, 3, 1),
                      sale_consideration=90000, cost_of_acquisition=120000)
    s = compute_cg_summary([a1, a2])
    assert s["STCG20"] == 20000         # +50k - 30k netted in band
