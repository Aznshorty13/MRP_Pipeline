"""Machine-readable audit CSVs for dashboard chart source data."""

from pathlib import Path

import pandas as pd

STATUS_CODE_MAP = {
    "Stable": 0,
    "Normal Execution": 1,
    "⚠️ CAPACITY BREACH": 2,
    "🚨 MAGIC FIX (Past Due)": 3,
}

DASHBOARD_ROOT = Path("output") / "dashboards"

SKU_AUDIT_COLUMNS = [
    "SKU_ID",
    "Date_Index",
    "Demand",
    "Scheduled_Receipts",
    "Planned_Receipts",
    "Planned_Releases",
    "Locked_Inv",
    "Order_Type",
    "Status_Code",
    "Unit_Cost",
    "Unit_Revenue",
    "Safety_Stock_Cost",
    "Capital_Tied_Up",
    "Capital_Required",
    "Revenue_at_Risk",
    "Safety_Stock_Units",
    "Max_Cap",
]


def dashboard_audit_path(phase: str) -> Path:
    """Full path for a phase dashboard audit CSV under output/dashboards/{phase}/."""
    folder = DASHBOARD_ROOT / phase.lower()
    folder.mkdir(parents=True, exist_ok=True)
    return folder / "dashboard_audit.csv"


def build_sku_dashboard_audit_df(df_enriched: pd.DataFrame, master_data: dict) -> pd.DataFrame:
    """Build long-format audit data matching Alpha/Beta dashboard inputs."""
    df = df_enriched.copy()
    df["Status_Code"] = df["Order_Type"].map(STATUS_CODE_MAP)
    df["Safety_Stock_Units"] = df["SKU_ID"].map(lambda sku: master_data[sku]["SS"])
    df["Max_Cap"] = df["SKU_ID"].map(lambda sku: master_data[sku]["Max_Cap"])
    return df[SKU_AUDIT_COLUMNS]


def build_delta_dashboard_audit_df(
    df_alpha: pd.DataFrame,
    df_beta: pd.DataFrame,
    master_data: dict,
) -> pd.DataFrame:
    """Build long-format audit data matching Delta dashboard inputs."""
    df_join = pd.merge(
        df_alpha,
        df_beta,
        on=["SKU_ID", "Date_Index"],
        suffixes=("_Alpha", "_Beta"),
    )
    df_join["Unit_Cost"] = df_join["SKU_ID"].map(lambda sku: master_data[sku]["Unit_Cost"])
    df_join["Action_Delta"] = (
        df_join["Planned_Releases_Beta"] - df_join["Planned_Releases_Alpha"]
    )
    df_join["Capital_Delta"] = df_join["Action_Delta"] * df_join["Unit_Cost"]
    df_join["Revenue_Risk_Delta"] = (
        df_join["Revenue_at_Risk_Beta"] - df_join["Revenue_at_Risk_Alpha"]
    )
    df_join["Safety_Stock_Units"] = df_join["SKU_ID"].map(lambda sku: master_data[sku]["SS"])
    df_join["Max_Cap"] = df_join["SKU_ID"].map(lambda sku: master_data[sku]["Max_Cap"])

    delta_columns = [
        "SKU_ID",
        "Date_Index",
        "Locked_Inv_Alpha",
        "Locked_Inv_Beta",
        "Planned_Releases_Alpha",
        "Planned_Releases_Beta",
        "Planned_Receipts_Alpha",
        "Planned_Receipts_Beta",
        "Demand_Alpha",
        "Demand_Beta",
        "Revenue_at_Risk_Alpha",
        "Revenue_at_Risk_Beta",
        "Action_Delta",
        "Capital_Delta",
        "Revenue_Risk_Delta",
        "Unit_Cost",
        "Safety_Stock_Units",
        "Max_Cap",
    ]
    return df_join[delta_columns]


def export_sku_dashboard_audit(
    phase: str,
    df_enriched: pd.DataFrame,
    master_data: dict,
) -> Path:
    """Write Alpha or Beta dashboard audit CSV."""
    out_path = dashboard_audit_path(phase)
    df_audit = build_sku_dashboard_audit_df(df_enriched, master_data)
    df_audit.to_csv(out_path, index=False)
    print(f" -> Dashboard audit trail saved to {out_path}")
    return out_path


def export_delta_dashboard_audit(
    df_alpha: pd.DataFrame,
    df_beta: pd.DataFrame,
    master_data: dict,
) -> Path:
    """Write Delta dashboard audit CSV."""
    out_path = dashboard_audit_path("delta")
    df_audit = build_delta_dashboard_audit_df(df_alpha, df_beta, master_data)
    df_audit.to_csv(out_path, index=False)
    print(f" -> Dashboard audit trail saved to {out_path}")
    return out_path
