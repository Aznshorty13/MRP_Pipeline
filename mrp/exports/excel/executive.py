"""Alpha executive workbook export (deduped upgraded version with condensed alerts)."""

from pathlib import Path

import pandas as pd

from mrp.enrichment import calculate_inventory_health
from mrp.exports.output_paths import alpha_exec_report_path, executive_variance_report_path
from mrp.viz.dashboards import dashboard_path


def build_executive_workbook(
    df_enriched, capacity_alerts_list, filename: str | Path | None = None
):
    """Compiles metrics, action plans, and raw matrices into an executive .xlsx file."""
    filename = filename or alpha_exec_report_path()
    print(f"Generating Enterprise Excel Handoff: {filename}...")

    agg_capital_commitment = df_enriched["Capital_Required"].sum()
    health_grade, dead_stock = calculate_inventory_health(df_enriched)
    df_action_plan = df_enriched[
        (df_enriched["Planned_Releases"] > 0) & (df_enriched.groupby("SKU_ID").cumcount() < 3)
    ]

    condensed_capacity_alerts = []
    if capacity_alerts_list:
        alert_counts = {}
        for alert in capacity_alerts_list:
            parts = alert.split(" | ")
            if len(parts) >= 2:
                sku = parts[1].split(" in ")[0]
                alert_counts[sku] = alert_counts.get(sku, 0) + 1
            else:
                alert_counts[alert] = alert_counts.get(alert, 0) + 1

        for sku, count in alert_counts.items():
            if count > 1:
                condensed_capacity_alerts.append(
                    f"🚨 CAPACITY BREACH | {sku} | {count} Occurrences across 24-month horizon"
                )
            else:
                original_alert = next(a for a in capacity_alerts_list if sku in a)
                condensed_capacity_alerts.append(original_alert)

    df_magic_fixes = df_enriched[df_enriched["Order_Type"] == "🚨 MAGIC FIX (Past Due)"]
    condensed_magic_fixes = []
    if not df_magic_fixes.empty:
        grouped_fixes = df_magic_fixes.groupby("SKU_ID").agg(
            Count=("Date_Index", "count"),
            Total_Qty=("Planned_Receipts", "sum"),
        ).reset_index()
        for _, row in grouped_fixes.iterrows():
            if row["Count"] > 1:
                condensed_magic_fixes.append(
                    f"🚨 EXPEDITE | {row['SKU_ID']} | {row['Count']} Occurrences | "
                    f"Total Qty Forced to Day 0: {row['Total_Qty']}"
                )
            else:
                condensed_magic_fixes.append(
                    f"🚨 EXPEDITE | {row['SKU_ID']} | 1 Occurrence | "
                    f"Qty Forced to Day 0: {row['Total_Qty']}"
                )

    writer = pd.ExcelWriter(filename, engine="xlsxwriter")
    workbook = writer.book
    header_format = workbook.add_format({"bold": True, "font_size": 14, "bottom": 2})
    money_format = workbook.add_format({"num_format": "$#,##0"})
    percent_format = workbook.add_format({"num_format": "0.0%"})

    ws_exec = workbook.add_worksheet("Executive Summary")
    ws_exec.write("A1", "Alpha Engine: 24-Month Master Resource Plan", header_format)
    ws_exec.write("A3", "Aggregate Capital Commitment (PO Values):")
    ws_exec.write("B3", agg_capital_commitment, money_format)
    ws_exec.write("A4", "Inventory Health Grade (Active vs Dead):")
    ws_exec.write("B4", health_grade / 100, percent_format)
    ws_exec.write("A5", "Total Dead Stock Risk Exposure:")
    ws_exec.write("B5", dead_stock, money_format)

    ws_exec.write("A7", "Capacity/Risk Breaches (Condensed):", header_format)
    row = 8
    if not condensed_capacity_alerts:
        ws_exec.write(row, 0, "No capacity breaches detected in Day 0 baseline.")
        row += 1
    else:
        for alert in condensed_capacity_alerts:
            ws_exec.write(row, 0, str(alert))
            row += 1

    row += 1
    ws_exec.write(row, 0, "Applied 'Magic' Resolutions (Condensed):", header_format)
    row += 1
    if not condensed_magic_fixes:
        ws_exec.write(row, 0, "No automated fixes were necessary or applied.")
    else:
        for fix in condensed_magic_fixes:
            ws_exec.write(row, 0, str(fix))
            row += 1

    try:
        ws_exec.insert_image("D2", str(dashboard_path("alpha", "SKU_003")), {"x_scale": 0.6, "y_scale": 0.6})
    except FileNotFoundError:
        ws_exec.write("D2", "(Run dashboard generation first to embed visual here)")

    action_cols = ["SKU_ID", "Date_Index", "Order_Type", "Planned_Releases", "Capital_Required"]
    df_action_plan[action_cols].to_excel(writer, sheet_name="90-Day Action Plan", index=False)
    df_enriched.to_excel(writer, sheet_name="Raw Horizon Matrix", index=False)
    writer.close()
    print("✅ Executive Workbook saved successfully.")


def export_variance_workbook(df_raw_baseline, df_exceptions, df_summary):
    """Compiles the 4-tab enterprise variance workbook."""
    from datetime import datetime

    from mrp.delta import compress_to_campaigns

    print("Executing Blueprint 11.3: Compiling Executive Excel Workbook...")
    today_str = datetime.now().strftime("%Y-%m-%d_%H%M")
    filename = executive_variance_report_path(today_str)

    if df_raw_baseline.empty:
        print("🚨 Error: No baseline data provided to export.")
        return filename

    df_campaigns = compress_to_campaigns(df_exceptions)
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        if not df_summary.empty:
            df_summary.to_excel(writer, sheet_name="1_Executive_Summary", index=False)
        if not df_campaigns.empty:
            df_campaigns.to_excel(writer, sheet_name="2_Action_Campaigns", index=False)
        if not df_exceptions.empty:
            df_action = df_exceptions.sort_values(by=["Date_Index", "SKU_ID"])
            buyer_columns = [
                "SKU_ID",
                "Date_Index",
                "Action_Delta",
                "Capital_Variance",
                "Revenue_at_Risk_Beta",
            ]
            df_action[buyer_columns].to_excel(writer, sheet_name="3_Monthly_Execution_List", index=False)
        df_raw_baseline.to_excel(writer, sheet_name="4_Raw_Physics_Baseline", index=False)

        workbook = writer.book
        money_format = workbook.add_format({"num_format": "$#,##0"})
        for sheet_name in writer.sheets:
            writer.sheets[sheet_name].set_column("A:Z", 18)
            if sheet_name == "2_Action_Campaigns":
                writer.sheets[sheet_name].set_column("G:G", 18, money_format)

    print(f"✅ Executive Workbook successfully saved to disk: {filename}\n")
    return filename
