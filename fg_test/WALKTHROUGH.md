# FacilityGrid automation walkthrough

Use these scripts to investigate missing manufacturer, model, and asset information in FacilityGrid issue exports. They create Excel reports for review.

| Script | Purpose |
| --- | --- |
| [fg_web_automation.py](fg_web_automation.py) | Opens issue pages and inspects their hierarchy for manufacturer and model information. |
| [fg_api_automation.py](fg_api_automation.py) | Queries the issue API, compares selected results with PostgreSQL, and optionally cross-checks the web UI. |

## 1. Prepare the environment

Run these commands in PowerShell from the repository root. Use the existing `.venv`, or create it first with `python -m venv .venv`.

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m playwright install chromium
```

The root `requirements.txt` includes the dependencies needed by both scripts. Install the Chromium binaries separately with the command above. Playwright is needed even for API-only runs because the API script imports helpers from the web script.

If `.env` does not exist, copy the template without overwriting an existing configuration:

```powershell
if (!(Test-Path .env)) { Copy-Item env_example .env }
```

## 2. Configure `.env`

Edit the repository-root `.env`. The following values are examples only:

```dotenv
FG_USERNAME=user@example.com
FG_PASSWORD=dummy-password
FG_FOLDER_NAME=Day3_Data
FG_WEB_TEST_FILE_NAME=issues.xlsx
FG_API_TEST_FILE_NAME=issues.xlsx
PHASE_1_FILE_FOR_API_WEB_TEST=api_scan_results_YYYYMMDD_HHMMSS.xlsx

EXTERNAL_API_GRANT_TYPE=client_credentials
EXTERNAL_API_CLIENT_ID=00000000-0000-0000-0000-000000000000
EXTERNAL_API_CLIENT_SECRET=dummy-client-secret
EXTERNAL_API_AUTH_ENDPOINT=https://auth.example.com/oauth/token
EXTERNAL_API_BASE_URL=https://facilitygrid.example.com

DATABASE_URL=postgresql://demo_user:dummy-password@localhost:5432/demo_db
```

The web script needs browser credentials or an existing login session. Full API runs also need valid API credentials and endpoints. Standalone Phase 2 uses the browser without requesting an API token or connecting to the database.

`DATABASE_URL` takes precedence over `DB_HOST`, `DB_PORT`, `DB_NAME`, `DB_USER`, and `DB_PASSWORD`. If the database connection fails, the full API run continues with DB verification skipped. The comparison queries the `issues` table using `fg_issue_uuid` and reads `equipment_uuid`, `asset_uuid`, `subtask_uuid`, and `responsibility_name`. It checks for missing information, not exact equality of every UUID.

Run all remaining commands from `fg_test` so that `FG_FOLDER_NAME=Day3_Data` resolves correctly:

```powershell
Set-Location .\fg_test
```

Relative file and folder paths resolve from the current working directory. You can also use absolute paths. Both scripts load the root `.env` through `find_dotenv()`.

## 3. Prepare the input workbook

Both scripts read the first worksheet.

| Mode | Input columns |
| --- | --- |
| Web scan | `issue_url`; optional `manufacturer` controls the default missing-manufacturer filter. |
| Full API scan | Exact column names `fg_project_id` and `fg_issue_uuid` are required. `issue_code`, `project_name`, and `issue_url` add context. |
| Standalone Phase 2 | Prefer an existing API report. It accepts display headers such as `Issue URL`, `FG Project ID`, and `FG Issue ID`, as well as their source column names. |

Use complete issue URLs for web scans. The API script can construct issue URLs from project and issue IDs when a URL is absent. Rows without an issue UUID are excluded from full API scans.

## 4. Run a web scan

Start with a small sample:

```powershell
python fg_web_automation.py --limit 10 --workers 2
```

Chromium opens visibly. The script tries automatic login; if prompted, complete login or MFA in the browser and press Enter in PowerShell. Browser sessions persist under `.fg_playwright_profile`.

Scan all eligible rows:

```powershell
python fg_web_automation.py --limit 0 --workers 5 --batch-size 10
```

By default, rows with an existing `manufacturer` value are excluded when that column exists. Use `--all-rows` to include them. The default limit is 10; `--limit 0` removes the limit.

To choose files explicitly:

```powershell
python fg_web_automation.py --file .\Day3_Data\issues.xlsx --output .\Day3_Data\scan_results_review.xlsx --limit 0 --all-rows
```

Other options: `--timeout 30000` increases the per-page timeout to 30 seconds; `--headless` hides the browser after a working login session is established. Use `--workers` explicitly to control concurrency: the CLI default of 5 means the environment worker fallback is not normally used.

The report records extracted manufacturer/model values and their source. In the web report, **red means both manufacturer and model were found**; this differs from the API report's color scheme.

## 5. Run the API and database scan

Start with 10 rows:

```powershell
python fg_api_automation.py --limit 10 --workers 3 --web-workers 2
```

Phase 1 authenticates, queries issue details, and checks selected rows against the local database. Phase 2 uses browser pages to inspect issues marked for web checking. Its session is stored separately under `.fg_api_playwright_profile`.

Full run, including web checks for rows where the API supplies asset information:

```powershell
python fg_api_automation.py --limit 0 --workers 15 --web-workers 5 --check-all-web
```

API and DB checks only:

```powershell
python fg_api_automation.py --limit 0 --workers 15 --no-web-check
```

Unlike the web script, the API script defaults to all rows. `--workers` controls API threads; `--web-workers` controls browser pages. Reduce them if requests fail or the browser becomes overloaded.

Yellow rows are not checked on the web by default. `--check-all-web` (alias `--check-yellow`) enables those checks; `--no-check-yellow` skips them. If `FG_CHECK_ALL_WEB` is set in `.env`, its value overrides these flags (`true` enables checks; other values disable them).

## 6. Run Phase 2 separately

Use `PHASE_1_FILE_FOR_API_WEB_TEST` from `.env`:

```powershell
python fg_api_automation.py --phase-2 --limit 0 --web-workers 5 --output .\Day3_Data\api_scan_results_web_review.xlsx
```

Or select the input report explicitly:

```powershell
python fg_api_automation.py --phase-2 --file .\Day3_Data\api_scan_results_YYYYMMDD_HHMMSS.xlsx --output .\Day3_Data\api_scan_results_web_review.xlsx --web-workers 5
```

Replace the timestamp placeholder with an actual report name. Without `--file`, the script tries the configured Phase 1 file, then the latest API result in `FG_FOLDER_NAME`, then the configured API input workbook.

**Without `--output`, Phase 2 overwrites the input when its name contains `api_scan_results_`.** Use a separate output path when keeping the original report. A limited Phase 2 run writes only the selected subset.

Current limitation: standalone Phase 2 detects previous reports using lowercase internal column names, while generated reports use display headers and omit `color_assigned`. It can therefore recheck all rows and recompute colors rather than preserve the original classifications. For a combined API/DB/web classification, use a full run with `--check-all-web`.

## 7. Read the API report

The workbook contains `API Scan Results` plus filtered sheets named `Our Code Issues`, `FG API Issues`, `Cyan - Web Mfg Found`, and `Yellow - Web Mfg Missing`.

| Color | Interpretation |
| --- | --- |
| Cyan | API asset information was provided and a web check found manufacturer information. |
| Yellow | API asset information was provided. A web check may be skipped, pending, or unable to find a manufacturer; inspect `Web Cross-Check`. |
| Purple | DB comparison found missing asset information, missing recognized responsibility, or an issue absent from the DB. Inspect `DB Status`. |
| Orange | Web inspection found equipment, subtask, manufacturer, or model information for a fallback row. Investigate the API/web discrepancy. |
| Light blue | API asset information is absent, but responsibility matches the script's configured list. Inspect web-check status before concluding it is also absent on the web. |
| Red | API asset information is absent and responsibility does not match the configured list. Inspect web-check status and errors. |

Review `Status`, `DB Status`, `Web Cross-Check`, and `Web Source` alongside the colors. An unchecked row or failed request is not proof that data is missing. The responsibility list is defined in `fg_api_automation.py`.

## 8. Outputs and troubleshooting

Reports normally use timestamped `scan_results_*.xlsx` or `api_scan_results_*.xlsx` names. Output defaults to the configured folder when it exists, otherwise the input workbook's directory. An explicit output path containing a directory makes the destination unambiguous.

The web script saves checkpoints every 10 completed rows by default (`--batch-size`). API Phase 1 saves every 50 completed rows and at completion. These are saved progress reports, not automatic resume checkpoints.

| Problem | Action |
| --- | --- |
| Input file not found | Check the current directory, `FG_FOLDER_NAME`, and filename, or use an absolute `--file` path. |
| Missing Python module or Chromium executable | Run the installation commands in step 1 using the activated environment. |
| Login or MFA required | Run without `--headless`, complete browser login, then press Enter when prompted. |
| API authentication or HTTP errors | Check API credentials, endpoint, and access to the project; inspect `Status`. |
| DB verification skipped or DB error | Check database connectivity, credentials, and the `issues` table schema; inspect `DB Status`. |
| Excel file is open | Close it before running. On a save permission error, the scripts attempt an alternate timestamped filename and print its location. |
| Browser profile already in use | Close the previous run's browser before starting another instance of the same script. |

`.env`, browser profiles, generated scan reports, caches, and logs are covered by the repository's `.gitignore`. Keep reusable source workbooks distinct from generated report names.

For the full argument lists:

```powershell
python fg_web_automation.py --help
python fg_api_automation.py --help
```
