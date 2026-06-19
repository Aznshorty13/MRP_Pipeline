# MRP Pipeline

Modular Python refactor of the original Colab notebook. See **[Project_Documentation.md](Project_Documentation.md)** for blueprint-by-blueprint documentation. Legacy source files are in [`legacy/`](legacy/).

## Setup

```powershell
py -m pip install -r requirements.txt
```

On Windows, use **`py`** (not `python`) and UTF-8 mode so emoji in log output render correctly:

```powershell
py -X utf8 main.py --phase full
```

## Run

```powershell
# Full Alpha → Beta → Delta pipeline (includes dashboard PNGs)
py -X utf8 main.py --phase full

# Individual phases
py -X utf8 main.py --phase alpha
py -X utf8 main.py --phase beta
py -X utf8 main.py --phase delta

# Skip matplotlib dashboards (faster, no PNG files)
py -X utf8 main.py --phase full --no-dashboards

# Generate Excel sandbox + test fixture workbooks
py -X utf8 main.py --phase full --fixtures

# Smoke test (calendar only, no simulation)
py -X utf8 main.py --smoke-test
```

## Dashboard PNGs

By default, the pipeline generates matplotlib dashboard images under `output/dashboards/`:

```
output/dashboards/
├── alpha/   → SKU_001_dashboard.png … SKU_007_dashboard.png, dashboard_audit.csv
├── beta/    → SKU_001_dashboard.png … SKU_007_dashboard.png, dashboard_audit.csv
└── delta/   → SKU_001_dashboard.png … SKU_007_dashboard.png, dashboard_audit.csv
```

Each chart is a multi-panel view: inventory sawtooth, horizon status heatmap, and capital/risk curves (Alpha/Beta), or Alpha vs Beta variance (Delta).

**Dashboard audit CSV** (`dashboard_audit.csv` in each phase folder) exports the exact data used to render the charts — one row per SKU and month (plus derived delta columns for the delta phase). It is written on every phase run, including with `--no-dashboards`, so you can validate chart inputs without reading PNG files.

The Alpha executive workbook embeds `output/dashboards/alpha/SKU_003_dashboard.png` when that file exists.

Implementation: [`mrp/viz/dashboards.py`](mrp/viz/dashboards.py), [`mrp/viz/dashboard_audit.py`](mrp/viz/dashboard_audit.py), wired from [`pipeline/runner.py`](pipeline/runner.py).

## Phase exports

CSV, trace, and executive workbook artifacts are written under `output/exports/`:

```
output/exports/
├── alpha/   → Alpha_Raw_Horizon.csv, Alpha_Exec_Report.xlsx, …
├── beta/    → Beta_All_SKUs_Trace.txt, Beta_Purchasing_Cadence.csv
└── delta/   → Executive_Variance_Report_*.xlsx
```

Optional Excel sandbox and CI test fixture workbooks (`--fixtures`) are written to `output/fixtures/`.

On each pipeline run, any legacy copies of these artifacts still sitting in the project root are removed automatically.

Implementation: [`mrp/exports/output_paths.py`](mrp/exports/output_paths.py), [`mrp/exports/csv_exports.py`](mrp/exports/csv_exports.py), [`mrp/exports/excel/executive.py`](mrp/exports/excel/executive.py), [`mrp/exports/excel/fixtures.py`](mrp/exports/excel/fixtures.py).

## System of Record

Shadow ledger workbooks (Excel + Google Sheets-ready CSV tabs) are written under `output/system_of_record/`:

```
output/system_of_record/
├── alpha/
│   ├── System_of_Record_Alpha.xlsx
│   ├── google_sheets/
│   │   ├── Executive_Summary.csv
│   │   ├── Alpha_Master_Plan.csv
│   │   └── 90-Day_Action_Plan.csv
│   └── google_sheets_manifest.json   (when API upload succeeds)
├── beta/
│   └── ...
└── delta/
    └── ...
```

CSV tabs are generated automatically after each Excel workbook is saved. They carry data only (no conditional formatting); the `.xlsx` files remain the formatted audit artifact.

### Optional Google Sheets upload

To create live Google Spreadsheets (one per phase) and save shareable links in `google_sheets_manifest.json`:

1. Create a Google Cloud service account with **Google Sheets API** and **Google Drive API** enabled.
2. Download the JSON key to a local path (e.g. `google_sheets_service_account.json` — do not commit it).
3. Set the environment variable before running:

```powershell
$env:GOOGLE_SHEETS_CREDENTIALS = "google_sheets_service_account.json"
py -X utf8 main.py --phase full
```

Optional: set `GOOGLE_SHEETS_FOLDER_ID` to a Drive folder ID and share that folder with the service account email.

If credentials are not set, the pipeline still writes Excel and CSV exports and skips the API upload with a log message.

Implementation: [`mrp/exports/system_of_record_paths.py`](mrp/exports/system_of_record_paths.py), [`mrp/exports/google_sheets_export.py`](mrp/exports/google_sheets_export.py), [`mrp/exports/excel/shadow_ledgers.py`](mrp/exports/excel/shadow_ledgers.py).

## Project layout

| Path | Purpose |
|------|---------|
| [`data/fixtures.py`](data/fixtures.py) | Master data, inventory, demand, chaos payload |
| [`mrp/simulation.py`](mrp/simulation.py) | SKU state-machine simulation |
| [`mrp/state.py`](mrp/state.py) | Beta inheritance and chaos injection |
| [`mrp/enrichment.py`](mrp/enrichment.py) | Baseline enrichment and health metrics |
| [`mrp/delta.py`](mrp/delta.py) | Calendar join and exception rollup |
| [`mrp/viz/dashboards.py`](mrp/viz/dashboards.py) | Dashboard PNG generation |
| [`mrp/viz/dashboard_audit.py`](mrp/viz/dashboard_audit.py) | Dashboard audit CSV export |
| [`mrp/exports/output_paths.py`](mrp/exports/output_paths.py) | Phase export and fixture output paths |
| [`mrp/exports/system_of_record_paths.py`](mrp/exports/system_of_record_paths.py) | System of Record output paths |
| [`mrp/exports/google_sheets_export.py`](mrp/exports/google_sheets_export.py) | CSV tab export + optional Google Sheets upload |
| [`pipeline/runner.py`](pipeline/runner.py) | `run_alpha()`, `run_beta()`, `run_delta()`, `run_full()` |
| [`main.py`](main.py) | CLI entry point |
| [`legacy/`](legacy/) | Original code, Colab backup, and prior documentation |

Full blueprint documentation: **[Project_Documentation.md](Project_Documentation.md)**
