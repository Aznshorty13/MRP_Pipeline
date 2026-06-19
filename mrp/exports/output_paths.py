"""Path helpers for phase export artifacts under output/exports/ and output/fixtures/."""

from pathlib import Path

OUTPUT_ROOT = Path("output")
EXPORTS_ROOT = OUTPUT_ROOT / "exports"
FIXTURES_ROOT = OUTPUT_ROOT / "fixtures"


def exports_dir(phase: str) -> Path:
    """Return the export directory for a phase (alpha, beta, or delta), creating it if needed."""
    folder = EXPORTS_ROOT / phase.lower()
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def export_path(phase: str, filename: str) -> Path:
    """Full path for a named artifact under output/exports/{phase}/."""
    return exports_dir(phase) / filename


def fixture_path(filename: str) -> Path:
    """Full path for a fixture workbook under output/fixtures/."""
    FIXTURES_ROOT.mkdir(parents=True, exist_ok=True)
    return FIXTURES_ROOT / filename


def alpha_raw_horizon_path() -> Path:
    return export_path("alpha", "Alpha_Raw_Horizon.csv")


def alpha_exception_log_path() -> Path:
    return export_path("alpha", "Alpha_Exception_Log.csv")


def alpha_trace_path() -> Path:
    return export_path("alpha", "Alpha_All_SKUs_Trace.txt")


def alpha_cadence_path() -> Path:
    return export_path("alpha", "Alpha_Purchasing_Cadence.csv")


def alpha_exec_report_path() -> Path:
    return export_path("alpha", "Alpha_Exec_Report.xlsx")


def beta_trace_path() -> Path:
    return export_path("beta", "Beta_All_SKUs_Trace.txt")


def beta_cadence_path() -> Path:
    return export_path("beta", "Beta_Purchasing_Cadence.csv")


def executive_variance_report_path(date_str: str) -> Path:
    return export_path("delta", f"Executive_Variance_Report_{date_str}.xlsx")


def enterprise_sandbox_path() -> Path:
    return fixture_path("Enterprise_MRP_Sandbox.xlsx")


def enterprise_test_fixture_path() -> Path:
    return fixture_path("Enterprise_MRP_TEST_FIXTURE.xlsx")


LEGACY_ROOT_GLOBS = (
    "Alpha_*",
    "Beta_*",
    "Executive_Variance_Report_*.xlsx",
    "Enterprise_MRP_*.xlsx",
    "System_of_Record_*.xlsx",
)


def cleanup_legacy_root_artifacts(root: Path | None = None) -> list[Path]:
    """Remove pipeline artifacts left in the project root before output/ migration."""
    root = root or Path(".")
    removed: list[Path] = []
    for pattern in LEGACY_ROOT_GLOBS:
        for path in root.glob(pattern):
            if path.is_file():
                path.unlink()
                removed.append(path)
    return removed
