"""SKU and delta visualization dashboards."""

from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import pandas as pd
import seaborn as sns
from matplotlib.patches import Patch

from mrp.viz.dashboard_audit import DASHBOARD_ROOT, STATUS_CODE_MAP


def dashboard_dir(phase: str) -> Path:
    """Return the output directory for a phase (alpha, beta, or delta), creating it if needed."""
    folder = DASHBOARD_ROOT / phase.lower()
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def dashboard_path(phase: str, sku_id: str) -> Path:
    """Full path for a SKU dashboard PNG under output/dashboards/{phase}/."""
    return dashboard_dir(phase) / f"{sku_id}_dashboard.png"


def plot_sku_dashboard(sku_id, df_enriched, master_data, prefix="Alpha"):
    df = df_enriched[df_enriched["SKU_ID"] == sku_id].copy()
    dates = df["Date_Index"]
    max_cap = master_data[sku_id]["Max_Cap"]

    fig, axes = plt.subplots(3, 1, figsize=(14, 12), gridspec_kw={"height_ratios": [3, 1, 2]})
    fig.suptitle(f"Operational Dashboard ({prefix} Timeline): {sku_id}", fontsize=16, fontweight="bold")

    ax1 = axes[0]
    ax1.step(dates, df["Locked_Inv"], where="post", color="#2c3e50", linewidth=2, label="Locked Inventory (I_t)")
    ax1.bar(dates, df["Planned_Receipts"], color="#3498db", alpha=0.6, label="Planned Receipts (Arrivals)")
    ax1.axhline(
        y=df["Safety_Stock_Cost"].iloc[0] / df["Unit_Cost"].iloc[0],
        color="red",
        linestyle="--",
        alpha=0.5,
        label="Safety Stock Target",
    )
    ax1.axhline(
        y=max_cap,
        color="#d35400",
        linestyle="-.",
        linewidth=2,
        alpha=0.8,
        label=f"Supplier Max Capacity ({max_cap})",
    )
    ax1.set_title("Physical Inventory & Replenishment", loc="left")
    ax1.set_ylabel("Units")
    ax1.tick_params(axis="x", rotation=45)
    ax1.legend(loc="upper right")
    ax1.grid(True, linestyle=":", alpha=0.6)

    ax2 = axes[1]
    df["Status_Code"] = df["Order_Type"].map(STATUS_CODE_MAP)
    heatmap_data = df["Status_Code"].values.reshape(1, 24)
    cmap = sns.color_palette(["#bdc3c7", "#2ecc71", "#f39c12", "#e74c3c"])

    sns.heatmap(heatmap_data, cmap=cmap, cbar=False, ax=ax2, linewidths=0.5, linecolor="white", vmin=0, vmax=3)
    ax2.set_title(
        "Horizon Map (Gray=Stable, Green=Routine, Orange=Capacity Breach, Red=Timing Breach)",
        loc="left",
    )
    ax2.set_xticks(np.arange(24) + 0.5)
    ax2.set_xticklabels(dates, rotation=45, ha="right")
    ax2.set_yticks([])

    ax3 = axes[2]
    ax3.fill_between(dates, df["Capital_Tied_Up"], color="#9b59b6", alpha=0.3)
    ax3.plot(dates, df["Capital_Tied_Up"], color="#8e44ad", linewidth=2, label="Working Capital ($)")
    ax3.set_ylabel("Working Capital ($)", color="#8e44ad", fontweight="bold")
    ax3.tick_params(axis="y", labelcolor="#8e44ad")
    ax3.tick_params(axis="x", rotation=45)
    ax3.grid(True, linestyle=":", alpha=0.6)

    ax3_risk = ax3.twinx()
    ax3_risk.bar(dates, df["Revenue_at_Risk"], color="#e74c3c", alpha=0.4, width=0.6, label="Revenue at Risk")
    ax3_risk.set_ylabel("Revenue at Risk ($)", color="#e74c3c", fontweight="bold")
    ax3_risk.tick_params(axis="y", labelcolor="#e74c3c")

    lines_1, labels_1 = ax3.get_legend_handles_labels()
    lines_2, labels_2 = ax3_risk.get_legend_handles_labels()
    ax3.legend(lines_1 + lines_2, labels_1 + labels_2, loc="upper right")
    ax3.set_title("Capital Allocation vs. Revenue at Risk", loc="left")

    plt.tight_layout()
    out_path = dashboard_path(prefix, sku_id)
    plt.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close()
    return out_path


def generate_all_sku_dashboards(df_enriched, master_data, prefix="Alpha"):
    sku_list = df_enriched["SKU_ID"].unique()
    folder = dashboard_dir(prefix)
    print(f"Generating 4-Color Visual Dashboards ({prefix} Timeline) for {len(sku_list)} SKUs...")
    print(f" -> Saving to {folder}\n")
    for sku in sku_list:
        plot_sku_dashboard(sku, df_enriched, master_data, prefix)
    print(f"✅ All {prefix} Dashboards processed and saved.")


def plot_delta_dashboard(sku_id, df_alpha, df_beta, master_data):
    """Four-panel executive dashboard comparing Alpha vs Beta."""
    df_a = df_alpha[df_alpha["SKU_ID"] == sku_id].reset_index(drop=True)
    df_b = df_beta[df_beta["SKU_ID"] == sku_id].reset_index(drop=True)

    df_join = pd.merge(df_a, df_b, on="Date_Index", suffixes=("_Alpha", "_Beta"))
    dates = df_join["Date_Index"]

    unit_cost = master_data[sku_id]["Unit_Cost"]
    max_cap = master_data[sku_id]["Max_Cap"]

    action_delta = df_join["Planned_Releases_Beta"] - df_join["Planned_Releases_Alpha"]
    capital_delta = action_delta * unit_cost
    revenue_risk_delta = df_join["Revenue_at_Risk_Beta"] - df_join["Revenue_at_Risk_Alpha"]

    fig, axes = plt.subplots(4, 1, figsize=(14, 16), gridspec_kw={"height_ratios": [3, 2, 2, 3]})
    fig.suptitle(f"Delta Variance Dashboard: {sku_id}", fontsize=16, fontweight="bold")

    ax1 = axes[0]
    ax1.step(
        dates,
        df_join["Locked_Inv_Alpha"],
        where="post",
        color="#95a5a6",
        linestyle="--",
        linewidth=2,
        alpha=0.7,
        label="Day 0 Plan (Ghost)",
    )
    ax1.step(
        dates,
        df_join["Locked_Inv_Beta"],
        where="post",
        color="#2c3e50",
        linewidth=3,
        label="Day X Reality (Actual)",
    )
    ax1.axhline(y=master_data[sku_id]["SS"], color="red", linestyle="--", alpha=0.5, label="Safety Stock")
    ax1.axhline(
        y=max_cap,
        color="#d35400",
        linestyle="-.",
        linewidth=2,
        alpha=0.8,
        label=f"Capacity ({max_cap})",
    )
    ax1.set_title("Comparative Inventory Physics (Alpha vs. Beta)", loc="left")
    ax1.set_ylabel("Units")
    ax1.tick_params(axis="x", rotation=45)
    ax1.legend(loc="upper right")
    ax1.grid(True, linestyle=":", alpha=0.6)

    ax2 = axes[1]
    x = np.arange(len(dates))
    width = 0.35
    ax2.bar(x - width / 2, df_join["Planned_Releases_Alpha"], width, color="#bdc3c7", alpha=0.8, label="Day 0 Plan (Ghost)")
    ax2.bar(x + width / 2, df_join["Planned_Releases_Beta"], width, color="#3498db", alpha=0.9, label="Day X Reality (Required)")
    ax2.set_title("Absolute Pipeline Volume (Comparing Total Order Scale)", loc="left")
    ax2.set_ylabel("Order Qty (Units)")
    ax2.set_xticks(x)
    ax2.set_xticklabels(dates, rotation=45, ha="right")
    ax2.legend(loc="upper right")
    ax2.grid(True, linestyle=":", alpha=0.6)

    ax3 = axes[2]
    delta_colors = ["#e74c3c" if val > 0 else "#2ecc71" for val in action_delta]
    ax3.bar(dates, action_delta, color=delta_colors, alpha=0.9)
    ax3.axhline(0, color="black", linewidth=1.5)
    legend_elements = [
        Patch(facecolor="#e74c3c", label="ACTION: Buy / Expedite"),
        Patch(facecolor="#2ecc71", label="ACTION: Cancel / Push Out"),
    ]
    ax3.legend(handles=legend_elements, loc="upper right")
    ax3.set_title("Human Action Required (Net Variance from Day 0 Plan)", loc="left")
    ax3.set_ylabel("Action (Units)")
    ax3.set_xticks(x)
    ax3.set_xticklabels(dates, rotation=45, ha="right")
    ax3.grid(True, linestyle=":", alpha=0.6)

    ax4 = axes[3]
    cap_colors = ["#2ecc71" if val < 0 else "#e74c3c" for val in capital_delta]
    ax4.bar(dates, capital_delta, color=cap_colors, alpha=0.7, label="Net Capital Variance ($)")
    ax4.set_ylabel("Unplanned Capital Flow ($)", fontweight="bold")
    ax4.tick_params(axis="x", rotation=45)
    ax4.axhline(0, color="black", linewidth=1)
    ax4.grid(True, linestyle=":", alpha=0.6)

    ax4_risk = ax4.twinx()
    ax4_risk.axhline(0, color="#f39c12", linestyle="--", alpha=0.4)
    ax4_risk.plot(
        dates,
        revenue_risk_delta,
        color="#f39c12",
        marker="o",
        linewidth=2,
        markersize=6,
        label="Net Revenue Impact ($)",
    )
    ax4_risk.set_yscale("symlog", linthresh=1000)
    ax4_risk.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: f"${y:,.0f}"))
    ax4_risk.set_ylabel("Net Rev Impact ($) [Spike=Loss, Dip=Save]", color="#f39c12", fontweight="bold")
    ax4_risk.tick_params(axis="y", labelcolor="#f39c12")

    lines_1, labels_1 = ax4.get_legend_handles_labels()
    lines_2, labels_2 = ax4_risk.get_legend_handles_labels()
    ax4.legend(lines_1 + lines_2, labels_1 + labels_2, loc="upper right")
    ax4.set_title("Financial Consequence: Capital Flow vs. Net Revenue Impact", loc="left")

    plt.tight_layout()
    out_path = dashboard_path("delta", sku_id)
    plt.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close()
    return out_path


def generate_all_delta_dashboards(df_alpha, df_beta, master_data):
    folder = dashboard_dir("delta")
    print("Generating Comparative Delta Dashboards...")
    print(f" -> Saving to {folder}\n")
    for sku in df_alpha["SKU_ID"].unique():
        plot_delta_dashboard(sku, df_alpha, df_beta, master_data)
    print("✅ All Delta Dashboards processed and saved.")
