"""Tests for phase export and fixture output path helpers."""

from mrp.exports.output_paths import (
    EXPORTS_ROOT,
    FIXTURES_ROOT,
    alpha_exec_report_path,
    alpha_raw_horizon_path,
    enterprise_sandbox_path,
    enterprise_test_fixture_path,
    export_path,
    exports_dir,
    fixture_path,
)


def test_alpha_raw_horizon_path():
    path = alpha_raw_horizon_path()
    assert path == EXPORTS_ROOT / "alpha" / "Alpha_Raw_Horizon.csv"


def test_alpha_exec_report_path():
    path = alpha_exec_report_path()
    assert path == EXPORTS_ROOT / "alpha" / "Alpha_Exec_Report.xlsx"


def test_export_path_case_insensitive():
    path = export_path("BETA", "Beta_All_SKUs_Trace.txt")
    assert path == EXPORTS_ROOT / "beta" / "Beta_All_SKUs_Trace.txt"


def test_fixture_paths():
    assert enterprise_sandbox_path() == FIXTURES_ROOT / "Enterprise_MRP_Sandbox.xlsx"
    assert enterprise_test_fixture_path() == FIXTURES_ROOT / "Enterprise_MRP_TEST_FIXTURE.xlsx"


def test_fixture_path_creates_folder(tmp_path, monkeypatch):
    monkeypatch.setattr("mrp.exports.output_paths.FIXTURES_ROOT", tmp_path / "fixtures")
    path = fixture_path("Enterprise_MRP_Sandbox.xlsx")
    assert path.parent.is_dir()
    assert path == tmp_path / "fixtures" / "Enterprise_MRP_Sandbox.xlsx"


def test_exports_dir_creates_folder(tmp_path, monkeypatch):
    monkeypatch.setattr("mrp.exports.output_paths.EXPORTS_ROOT", tmp_path / "exports")
    folder = exports_dir("delta")
    assert folder.is_dir()
    assert folder == tmp_path / "exports" / "delta"


def test_cleanup_legacy_root_artifacts(tmp_path):
    (tmp_path / "Alpha_Raw_Horizon.csv").write_text("old", encoding="utf-8")
    (tmp_path / "Beta_Purchasing_Cadence.csv").write_text("old", encoding="utf-8")
    (tmp_path / "Executive_Variance_Report_2026-01-01_1200.xlsx").write_bytes(b"")
    (tmp_path / "Enterprise_MRP_Sandbox.xlsx").write_bytes(b"")
    (tmp_path / "main.py").write_text("# keep", encoding="utf-8")

    from mrp.exports.output_paths import cleanup_legacy_root_artifacts

    removed = cleanup_legacy_root_artifacts(tmp_path)
    assert len(removed) == 4
    assert (tmp_path / "main.py").exists()
    assert not (tmp_path / "Alpha_Raw_Horizon.csv").exists()
