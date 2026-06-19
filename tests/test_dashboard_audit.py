"""Tests for dashboard audit CSV exports."""

import pandas as pd

from mrp.viz.dashboard_audit import (
    STATUS_CODE_MAP,
    build_delta_dashboard_audit_df,
    build_sku_dashboard_audit_df,
    dashboard_audit_path,
)
from mrp.viz.dashboard_audit import DASHBOARD_ROOT, STATUS_CODE_MAP


def _minimal_enriched_row(sku, date, order_type="Stable", planned_receipts=0):
    return {
        "SKU_ID": sku,
        "Date_Index": date,
        "Demand": 10,
        "Scheduled_Receipts": 0,
        "Planned_Receipts": planned_receipts,
        "Planned_Releases": 0,
        "Locked_Inv": 50,
        "Order_Type": order_type,
        "Unit_Cost": 100.0,
        "Unit_Revenue": 200.0,
        "Safety_Stock_Cost": 2000.0,
        "Capital_Tied_Up": 5000.0,
        "Capital_Required": 0.0,
        "Revenue_at_Risk": 0.0,
    }


def _master_data():
    return {
        "SKU_A": {"SS": 20, "Max_Cap": 100, "Unit_Cost": 10.0},
        "SKU_B": {"SS": 30, "Max_Cap": 200, "Unit_Cost": 20.0},
    }


def test_dashboard_audit_path_alpha():
    path = dashboard_audit_path("alpha")
    assert path == DASHBOARD_ROOT / "alpha" / "dashboard_audit.csv"


def test_build_sku_dashboard_audit_status_code():
    df = pd.DataFrame(
        [
            _minimal_enriched_row("SKU_A", "2026-06", "Stable"),
            _minimal_enriched_row("SKU_A", "2026-07", "⚠️ CAPACITY BREACH", planned_receipts=150),
        ]
    )
    audit = build_sku_dashboard_audit_df(df, _master_data())

    assert list(audit["Status_Code"]) == [
        STATUS_CODE_MAP["Stable"],
        STATUS_CODE_MAP["⚠️ CAPACITY BREACH"],
    ]
    assert audit.loc[0, "Safety_Stock_Units"] == 20
    assert audit.loc[0, "Max_Cap"] == 100


def test_build_delta_dashboard_audit_action_delta():
    alpha = pd.DataFrame(
        [
            _minimal_enriched_row("SKU_A", "2026-06"),
            _minimal_enriched_row("SKU_B", "2026-06"),
        ]
    )
    alpha.loc[0, "Planned_Releases"] = 5
    alpha.loc[1, "Planned_Releases"] = 10

    beta = alpha.copy()
    beta.loc[0, "Planned_Releases"] = 15
    beta.loc[1, "Planned_Releases"] = 8

    audit = build_delta_dashboard_audit_df(alpha, beta, _master_data())

    assert len(audit) == 2
    row_a = audit[audit["SKU_ID"] == "SKU_A"].iloc[0]
    row_b = audit[audit["SKU_ID"] == "SKU_B"].iloc[0]

    assert row_a["Action_Delta"] == 10
    assert row_a["Capital_Delta"] == 100.0
    assert row_b["Action_Delta"] == -2
    assert row_b["Capital_Delta"] == -40.0


def test_delta_audit_row_count_matches_inner_join():
    dates_alpha = ["2026-06", "2026-07", "2026-08"]
    dates_beta = ["2026-07", "2026-08", "2026-09"]

    alpha_rows = [_minimal_enriched_row("SKU_A", d) for d in dates_alpha]
    beta_rows = [_minimal_enriched_row("SKU_A", d) for d in dates_beta]

    alpha = pd.DataFrame(alpha_rows)
    beta = pd.DataFrame(beta_rows)

    audit = build_delta_dashboard_audit_df(alpha, beta, _master_data())
    expected = pd.merge(alpha, beta, on=["SKU_ID", "Date_Index"])

    assert len(audit) == len(expected) == 2
