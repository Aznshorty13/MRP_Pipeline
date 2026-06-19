"""Orchestrates Alpha -> Beta -> Delta MRP pipeline phases."""

from dataclasses import dataclass, field

import pandas as pd

from data.fixtures import CHAOS_PAYLOAD, CONSTRAINTS, DEMAND, INVENTORY
from mrp.ai.supplier_comms import draft_supplier_emails, generate_ai_payload
from mrp.calendar import advance_rolling_horizon, generate_calendar_horizon
from mrp.delta import (
    execute_calendar_join,
    generate_executive_alerts,
)
from mrp.enrichment import enrich_baseline_matrix
from mrp.exports.csv_exports import (
    export_exception_log,
    export_full_pedagogical_trace,
    generate_cadence_matrix,
)
from mrp.exports.output_paths import (
    alpha_raw_horizon_path,
    beta_cadence_path,
    beta_trace_path,
    cleanup_legacy_root_artifacts,
    enterprise_test_fixture_path,
)
from mrp.exports.excel.executive import build_executive_workbook, export_variance_workbook
from mrp.exports.excel.fixtures import (
    generate_enterprise_sandbox,
    generate_enterprise_test_fixture,
    run_semantic_shadow_test,
)
from mrp.exports.excel.shadow_ledgers import (
    build_alpha_shadow_ledger,
    build_beta_shadow_ledger,
    build_delta_shadow_ledger,
)
from mrp.simulation import execute_beta_run, execute_sku_simulation
from mrp.state import apply_chaos_events, deep_copy_beta_state, extract_inherited_state
from mrp.viz.dashboard_audit import export_delta_dashboard_audit, export_sku_dashboard_audit
from mrp.viz.dashboards import generate_all_delta_dashboards, generate_all_sku_dashboards


@dataclass
class AlphaResult:
    df_raw: pd.DataFrame
    df_enriched: pd.DataFrame
    calendar_array: list[str]
    all_alerts: list[str] = field(default_factory=list)


@dataclass
class BetaResult:
    df_raw: pd.DataFrame
    df_enriched: pd.DataFrame
    calendar_array: list[str]
    mutated_master: dict
    mutated_initial: dict
    mutated_demand: dict
    beta_alerts: list[dict] = field(default_factory=list)


@dataclass
class DeltaResult:
    df_joined: pd.DataFrame
    df_executive_summary: pd.DataFrame
    df_detailed_exceptions: pd.DataFrame
    llm_json_string: str = ""


def _format_capacity_alerts(sku: str, alerts: list[dict]) -> list[str]:
    formatted = []
    for alert in alerts:
        month = alert.get("Month", alert.get("Month", ""))
        formatted.append(
            f"🚨 CAPACITY BREACH | {sku} in {month} | "
            f"Required: {alert['Required']} vs Limit: {alert['Limit']}"
        )
    return formatted


def run_alpha(
    constraints: dict | None = None,
    inventory: dict | None = None,
    demand: dict | None = None,
    start_date: str = "2026-06-01",
    export_csv: bool = True,
    export_excel: bool = True,
    generate_dashboards: bool = True,
) -> AlphaResult:
    """Run Phase I Alpha simulation and enrichment."""
    constraints = constraints or CONSTRAINTS
    inventory = inventory or INVENTORY
    demand = demand or DEMAND

    removed = cleanup_legacy_root_artifacts()
    if removed:
        print(f"Removed {len(removed)} legacy root artifact(s).")

    calendar_array = generate_calendar_horizon(start_date=start_date)
    print(f"✅ Data Generation Suite Initialized. Horizon anchored at: {calendar_array[0]}")
    print("Initializing Alpha Engine State Machine...")

    all_sku_dfs = []
    all_alerts: list[str] = []

    for sku in constraints.keys():
        df, alerts = execute_sku_simulation(
            sku, constraints[sku], inventory[sku], demand[sku], calendar_array
        )
        all_sku_dfs.append(df)
        all_alerts.extend(_format_capacity_alerts(sku, alerts))

    df_enterprise_matrix = pd.concat(all_sku_dfs, ignore_index=True)
    df_alpha_enriched = enrich_baseline_matrix(df_enterprise_matrix, constraints)
    print("✅ Baseline Matrix Enriched. Revenue at Risk calculated.")

    export_sku_dashboard_audit("alpha", df_alpha_enriched, constraints)

    if export_csv:
        df_enterprise_matrix.to_csv(alpha_raw_horizon_path(), index=False)
        export_exception_log(df_alpha_enriched)
        export_full_pedagogical_trace(df_alpha_enriched, constraints)
        generate_cadence_matrix(df_alpha_enriched, calendar_array)

    if generate_dashboards:
        generate_all_sku_dashboards(df_alpha_enriched, constraints, prefix="Alpha")

    if export_excel:
        build_executive_workbook(df_alpha_enriched, all_alerts)
        build_alpha_shadow_ledger(
            df_alpha_enriched, list(constraints.keys()), calendar_array, constraints
        )

    print(f"Master Matrix generated with {len(df_enterprise_matrix)} rows.")
    return AlphaResult(df_enterprise_matrix, df_alpha_enriched, calendar_array, all_alerts)


def run_beta(
    alpha: AlphaResult,
    constraints: dict | None = None,
    demand: dict | None = None,
    chaos_payload: list | None = None,
    export_artifacts: bool = True,
    generate_dashboards: bool = True,
) -> BetaResult:
    """Run Phase II Beta: inheritance, chaos, recalculation."""
    constraints = constraints or CONSTRAINTS
    demand = demand or DEMAND
    chaos_payload = chaos_payload or CHAOS_PAYLOAD

    print("=== INITIATING PHASE II: THE BETA ENGINE ===")
    beta_calendar_array = advance_rolling_horizon(alpha.calendar_array)

    beta_initial_state, beta_master_data, beta_demand = extract_inherited_state(
        alpha.df_enriched, constraints, demand
    )

    mutated_initial, mutated_master, mutated_demand = deep_copy_beta_state(
        beta_initial_state, beta_master_data, beta_demand
    )
    mutated_initial, mutated_master, mutated_demand = apply_chaos_events(
        mutated_initial, mutated_master, mutated_demand, beta_calendar_array, chaos_payload
    )
    print("\n✅ Day X Matrix Successfully Mutated.")

    df_beta_raw, beta_alerts = execute_beta_run(
        beta_calendar_array, mutated_master, mutated_initial, mutated_demand
    )
    df_beta_enriched = enrich_baseline_matrix(df_beta_raw, mutated_master)
    print("✅ Beta Run Complete. Day X Timeline Enriched and Costed.")

    export_sku_dashboard_audit("beta", df_beta_enriched, mutated_master)

    if export_artifacts:
        export_full_pedagogical_trace(
            df_beta_enriched, mutated_master, filename=beta_trace_path()
        )
        generate_cadence_matrix(
            df_beta_enriched, beta_calendar_array, filename=beta_cadence_path()
        )

    if generate_dashboards:
        generate_all_sku_dashboards(df_beta_enriched, mutated_master, prefix="Beta")

    build_beta_shadow_ledger(
        df_beta_enriched,
        list(mutated_master.keys()),
        beta_calendar_array,
        mutated_master,
        chaos_payload,
    )

    return BetaResult(
        df_beta_raw,
        df_beta_enriched,
        beta_calendar_array,
        mutated_master,
        mutated_initial,
        mutated_demand,
        beta_alerts,
    )


def run_delta(
    alpha: AlphaResult,
    beta: BetaResult,
    constraints: dict | None = None,
    export_artifacts: bool = True,
    generate_dashboards: bool = True,
    run_ai: bool = True,
) -> DeltaResult:
    """Run Phase III Delta: join, exceptions, exports, AI payload."""
    constraints = constraints or CONSTRAINTS

    df_joined = execute_calendar_join(alpha.df_enriched, beta.df_enriched)
    df_executive_summary, df_detailed_exceptions = generate_executive_alerts(
        df_joined, constraints
    )

    export_delta_dashboard_audit(
        alpha.df_enriched, beta.df_enriched, beta.mutated_master
    )

    if generate_dashboards:
        generate_all_delta_dashboards(
            alpha.df_enriched, beta.df_enriched, beta.mutated_master
        )

    if export_artifacts:
        export_variance_workbook(df_joined, df_detailed_exceptions, df_executive_summary)
        build_delta_shadow_ledger(
            alpha.df_enriched,
            beta.df_enriched,
            list(beta.mutated_master.keys()),
            beta.calendar_array,
            beta.mutated_master,
        )

    llm_json_string = ""
    if run_ai:
        llm_json_string = generate_ai_payload(df_detailed_exceptions)
        if llm_json_string:
            draft_supplier_emails(llm_json_string, constraints)

    return DeltaResult(
        df_joined, df_executive_summary, df_detailed_exceptions, llm_json_string
    )


def run_full(
    export_fixtures: bool = False,
    run_semantic_test: bool = False,
    generate_dashboards: bool = True,
    export_csv: bool = True,
    export_excel: bool = True,
) -> tuple[AlphaResult, BetaResult, DeltaResult]:
    """Run complete Alpha -> Beta -> Delta pipeline."""
    alpha = run_alpha(
        generate_dashboards=generate_dashboards,
        export_csv=export_csv,
        export_excel=export_excel,
    )
    beta = run_beta(alpha, generate_dashboards=generate_dashboards)
    delta = run_delta(alpha, beta, generate_dashboards=generate_dashboards)

    if export_fixtures:
        generate_enterprise_sandbox(CONSTRAINTS, INVENTORY, DEMAND, alpha.calendar_array)
        generate_enterprise_test_fixture(
            list(CONSTRAINTS.keys()),
            CONSTRAINTS,
            INVENTORY,
            DEMAND,
            beta.mutated_demand,
            beta.mutated_master,
            alpha.calendar_array,
        )

    if run_semantic_test:
        run_semantic_shadow_test(
            enterprise_test_fixture_path(),
            alpha.df_enriched,
            beta.df_enriched,
            list(CONSTRAINTS.keys()),
            alpha.calendar_array,
        )

    return alpha, beta, delta
