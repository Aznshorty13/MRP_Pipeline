"""CSV export utilities."""

from pathlib import Path

import pandas as pd

from mrp.exports.output_paths import (
    alpha_cadence_path,
    alpha_exception_log_path,
    alpha_trace_path,
)


def export_exception_log(df_enriched, filename: str | Path | None = None):
    """Filters for timing exceptions and saves them to a CSV."""
    filename = filename or alpha_exception_log_path()
    print(f"Generating Exception Log: {filename}...")
    df_exceptions = df_enriched[df_enriched["Order_Type"] == "🚨 MAGIC FIX (Past Due)"]
    if df_exceptions.empty:
        print("✅ No critical timing exceptions detected. Log is empty.")
    else:
        action_cols = ["SKU_ID", "Date_Index", "Planned_Receipts", "Capital_Required"]
        df_exceptions[action_cols].to_csv(filename, index=False)
        print(f"🚨 Caught {len(df_exceptions)} exceptions. Saved to {filename}.")


def export_full_pedagogical_trace(
    df_enriched,
    master_data,
    months_to_trace=24,
    filename: str | Path | None = None,
):
    """Writes pedagogical audit trace for all SKUs."""
    filename = filename or alpha_trace_path()
    print(f"Generating Complete Audit Trace: {filename}...")
    sku_list = df_enriched["SKU_ID"].unique()

    with open(filename, "w", encoding="utf-8") as f:
        for sku in sku_list:
            target_ss = master_data[sku]["SS"]
            f.write("=========================================\n")
            f.write(f"=== 🔍 PEDAGOGICAL TRACE: {sku} ===\n")
            f.write(f"Target Safety Stock (SS): {target_ss}\n")
            f.write("=========================================\n\n")

            df_sku = df_enriched[df_enriched["SKU_ID"] == sku].head(months_to_trace)
            first_row = df_sku.iloc[0]
            prev_inv = (
                first_row["Locked_Inv"]
                - first_row["Planned_Receipts"]
                + first_row["Demand"]
                - first_row["Scheduled_Receipts"]
            )

            for _, row in df_sku.iterrows():
                t = row["Date_Index"]
                d_t = row["Demand"]
                r_t = row["Planned_Receipts"]
                sr_t = row["Scheduled_Receipts"]
                i_t = row["Locked_Inv"]
                unhealed = prev_inv + sr_t - d_t
                breach_flag = "⚠️ BREACH" if unhealed < target_ss else "✅ STABLE"
                f.write(f"Month: {t}\n")
                f.write(
                    f"  1. Raw Balance  -> {prev_inv} (I_t-1) + {sr_t} (SR_t) - {d_t} (D_t) = {unhealed}\n"
                )
                f.write(f"  2. Constraint   -> {unhealed} < {target_ss} ? {breach_flag}\n")
                if unhealed < target_ss:
                    f.write(f"  3. Arrival Fix  -> Injecting {r_t} (R_t) to heal timeline.\n")
                f.write(f"  4. Locked State -> {i_t} (I_t passed to next month)\n\n")
                prev_inv = i_t

        f.write("END OF MASTER AUDIT TRACE\n")
    print("✅ Master Trace written successfully.")


def generate_cadence_matrix(
    df_enriched, calendar_array, filename: str | Path | None = None
):
    """Pivots planned releases into a purchasing cadence matrix."""
    filename = filename or alpha_cadence_path()
    print("Initializing Component: Purchasing Cadence Matrix...\n")
    df_cadence = (
        df_enriched.pivot_table(
            index="SKU_ID",
            columns="Date_Index",
            values="Planned_Releases",
            aggfunc="sum",
        )
        .fillna(0)
    )
    for month in calendar_array:
        if month not in df_cadence.columns:
            df_cadence[month] = 0
    df_cadence = df_cadence[calendar_array]
    df_cadence.to_csv(filename)
    print("✅ Cadence Matrix Generated.")
