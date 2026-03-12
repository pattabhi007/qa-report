# QA Automation Report Tool

Generates daily QA test execution reports by pulling data from the [Testmo](https://www.testmo.com/) test management platform via its REST API. Designed for the **Go-To-Market & Business Solutions** Salesforce UI test suites that run nightly against GenMax QA and GenMax Staging environments.

## Scope

This tool covers **Salesforce UI Test Runs** executed through GitHub Actions workflows (`salesforce-ui-tests-matrix.yml`) that use Playwright and report results to Testmo. It tracks the following automation sources:

| Source | Description |
|--------|-------------|
| `contract-lifecycle-management` | CLM dashboard and workflow tests |
| `lead-to-opportunity` | L2O pipeline tests |
| `self-service-experience` | SSE portal tests |
| `quote-to-order` | Q2O quoting flow tests |
| `issue-to-resolution` | I2R support flow tests |
| `regression` | Cross-capability regression suite |

### Environments

| Environment | Tag | Run Name |
|-------------|-----|----------|
| GenMax QA | `env-genmax-qa` | `UI Test Run - GenMax QA` |
| GenMax Staging | `env-genmax-staging` | `UI Test Run - GenMax Staging` |

## Architecture

```
Testmo REST API                     Python Script                    Outputs
─────────────────                   ──────────────                   ───────
                                                                    
GET /projects/{id}/automation/runs  ┌──────────────────────┐        qa_automation_report.xlsx
GET /projects/{id}/automation/      │ qa_testmo_api_report  │───────►  └─ Latest Runs (styled)
    sources                         │                      │        
GET /automation/runs/{id}           │  TestmoClient         │        trend_genmax_qa.png
                                    │  ├─ fetch & paginate  │───────► trend_genmax_staging.png
    Bearer $TESTMO_TOKEN            │  ├─ filter UI runs    │        summary_genmax_qa.png
                                    │  ├─ map sources       │───────► summary_genmax_staging.png
                                    │  └─ build report      │
                                    └──────────────────────┘
```

### Testmo API Endpoints Used

| Endpoint | Purpose |
|----------|---------|
| `GET /api/v1/projects/{id}/automation/sources` | List all automation sources (capabilities) |
| `GET /api/v1/projects/{id}/automation/runs` | List automation runs with date filtering |
| `GET /api/v1/automation/runs/{id}` | Fetch a single run with aggregated status counts |

Authentication uses a Bearer token in the `Authorization` header, identical to the `TESTMO_TOKEN` used in GitHub Actions workflows.

### How Status Counts Are Extracted

Each automation run object includes pre-aggregated fields:

| API Field | Meaning |
|-----------|---------|
| `status1_count` | Passed |
| `failure_count` | Failed |
| `status5_count` | Blocked |
| `status6_count` | Skipped |
| `total_count` | Total tests |

## Prerequisites

- **Python 3.10+**
- **Testmo API token** — generate one from your [Testmo profile](https://genesys.testmo.net/users/profile) under the API access section

### Install Dependencies

```bash
pip install pandas matplotlib openpyxl requests
```

## Configuration

Set the following environment variables (same ones used by the GitHub Actions workflows):

```bash
export TESTMO_URL="https://genesys.testmo.net"
export TESTMO_TOKEN="testmo_api_..."
```

Optionally override the project ID (defaults to `11`):

```bash
export TESTMO_PROJECT_ID="11"
```

## Usage

### Fetch the last 7 days (default)

```bash
python qa_testmo_api_report.py
```

### Custom date range

```bash
# Last 14 days
python qa_testmo_api_report.py --days 14

# Last 30 days
python qa_testmo_api_report.py --days 30
```

### Fetch specific run IDs

```bash
python qa_testmo_api_report.py --run-ids 206847 206857 206909 206918 206940
```

### Pass credentials inline

```bash
python qa_testmo_api_report.py --url https://genesys.testmo.net --token YOUR_TOKEN
```

### All CLI options

```
usage: qa_testmo_api_report.py [-h] [--url URL] [--token TOKEN]
                                [--project-id PROJECT_ID] [--days DAYS]
                                [--run-ids [RUN_IDS ...]]

  --url           Testmo instance URL (or set TESTMO_URL env var)
  --token         Testmo API token (or set TESTMO_TOKEN env var)
  --project-id    Testmo project ID (default: 11)
  --days          Number of days to look back (default: 7)
  --run-ids       Specific automation run IDs (overrides --days)
```

## Output Files

### Excel Report — `qa_automation_report.xlsx`

A single **"Latest Runs"** sheet showing the most recent run per source, interleaved by environment:

| Source | Environment | RunID | Date | Status | Total | Passed | Failed | Skipped | Pass Rate % |
|--------|-------------|-------|------|--------|-------|--------|--------|---------|-------------|
| contract-lifecycle-management | QA | 206847 | 2026-03-10 | Success | 5 | 5 | 0 | 0 | 100.00 |
| contract-lifecycle-management | Staging | 206846 | 2026-03-10 | Success | 5 | 5 | 0 | 0 | 100.00 |
| lead-to-opportunity | QA | 206857 | 2026-03-10 | Failure | 68 | 18 | 43 | 7 | 26.47 |
| ... | | | | | | | | | |

**Styling:**
- Dark navy (`#1F3864`) headers for data columns
- Burnt orange (`#E36C09`) headers for Source and Environment
- Conditional formatting: green (>=80%), yellow (>=50%), red (<50%) on Pass Rate
- Green/red on Status (Success/Failure)
- Alternating row striping, frozen header, auto-fitted column widths

### Trend Charts

| File | Description |
|------|-------------|
| `trend_genmax_qa.png` | Daily pass rate trend lines per source for QA |
| `trend_genmax_staging.png` | Daily pass rate trend lines per source for Staging |
| `summary_genmax_qa.png` | Stacked bar chart (passed/failed/skipped) per source for QA |
| `summary_genmax_staging.png` | Stacked bar chart (passed/failed/skipped) per source for Staging |

Charts use the **full date range** (all runs in the `--days` window), not just the latest runs.

## Project Structure

```
qa-report/
├── qa_testmo_api_report.py    # Main script — fetches from Testmo API
├── qa_testmo_report_tool.py   # Legacy script — parses exported CSV files
├── test_runs/                 # Folder for CSV exports (used by legacy script)
├── qa_automation_report.xlsx  # Generated Excel report
├── trend_genmax_qa.png        # Generated QA trend chart
├── trend_genmax_staging.png   # Generated Staging trend chart
├── summary_genmax_qa.png      # Generated QA summary chart
├── summary_genmax_staging.png # Generated Staging summary chart
└── README.md
```

## Legacy CSV-Based Tool

The original `qa_testmo_report_tool.py` parses CSV files exported manually from Testmo. It is retained for backward compatibility but the API-based tool (`qa_testmo_api_report.py`) is the recommended approach.

```bash
# Place exported CSVs in test_runs/ then run:
python qa_testmo_report_tool.py
```

## Relationship to GitHub Workflows

The test data consumed by this tool is produced by the following GitHub Actions workflow chain:

```
salesforce-ui-tests-matrix.yml
  └─ salesforce-ui-tests.yml
       ├─ testmo automation:run:create     (creates a run in Testmo)
       ├─ testmo automation:run:submit-thread  (submits Playwright results per shard)
       └─ testmo automation:run:complete   (marks the run as complete)
```

Both the workflows and this tool authenticate using the same `TESTMO_URL` and `TESTMO_TOKEN` credentials.
