# 🛠️ Repository tooling

This directory indexes maintenance tools and the helpers of the SA
prototype, milestone v0.2.0 (issue #20). None of them holds calculation logic.

| Tool | Purpose |
| --- | --- |
| `Build-Workbook.ps1` | Builds `dist\K-FRTB-ASA.xlsm` from the template and the VBA modules through Excel COM on Windows (`-FromSource` without the template; `-Template` saves a source-built template) |
| `Build-Workbook.cmd` | Double-click launcher for `Build-Workbook.ps1`: asks for a production or development build and whether to overwrite |
| `Sync-FRTB-VBA.ps1` | Replaces the standard modules of an existing `.xlsm` through Excel COM, from the checkout (`-Mode Production` or `Development`) or from a folder (`-SourceDir`), with a safety copy; never writes cells or sheets |
| `Sync-FRTB-VBA.cmd`, `Sync-FRTB-VBA-Downloads.cmd` | Drag-and-drop launchers for `Sync-FRTB-VBA.ps1`: from the checkout, or from a folder that defaults to `%USERPROFILE%\Downloads` |
| `Export-FRTB-VBA.ps1` | Exports every VBA component of a workbook, read-only and with macros disabled, to an empty folder with `export-manifest.json` (workbook and file SHA-256, component types) |
| `compare_vba_export.py` | Compares an export with the source of a commit, read from Git, and checks its binding to the workbook; never runs Excel |
| `build_template.py` | Rebuilds (or `--check`s) the macro-free template `src/workbook/FRTB_Template.xlsx` headless in LibreOffice |
| `lo_runner.py` | Runs the VBA modules headless in LibreOffice Calc and saves the result as `.xlsx` |
| `vba_lint.py` | Static checks approximating Excel's compile checks and the module conventions |
| `check_excel_evidence.py` | Validates an Excel evidence record and its logs against the exact commit (`--evidence`), prints the source digests (`--inventory`) and checks the policy against `M_Setup` (`--check-policy`); never runs Excel. See [Excel evidence](../docs/EXCEL_EVIDENCE.md) |
| `Record-FRTB-ExcelRun.ps1`, `Record-FRTB-ExcelRun.cmd` | Record an Excel run of the checked-out commit in one step: build, compile, export and compare both workbooks, write `session.log` and the logs under `test-results\`, then call `record_excel_run.py`. See [Excel evidence](../docs/EXCEL_EVIDENCE.md#recorded-run) |
| `Check-FRTB-Excel.ps1`, `Check-FRTB-Excel.cmd` | Run every Excel check of a branch unattended: check it out, record the run, run `TEST_Acceptance` in Excel, the source and template builds, the Trust Center error path when asked, and the synchronization checks; write a summary and a zip. See [Excel evidence](../docs/EXCEL_EVIDENCE.md#unattended-check) |
| `build_acceptance_data.py` | Regenerates (or `--check`s) `tests/modules/TEST_AcceptanceData.bas`: the invalid-row variant, its reference results and the template's panes and filters, for the unattended check |
| `compare_templates.py` | Lists the differences between two templates (values, number formats, column widths, frozen panes); used by `Check-FRTB-Excel.ps1` for the Excel-built template |
| `record_excel_run.py` | Writes `host.json` of a recorded bundle from its logs only, and validates it with `check_excel_evidence.py`; never runs Excel |
| `check_external_links.py` | Requests the external links of the tracked Markdown files and reports them as verified, unverifiable or broken; run weekly by `link-check.yml`, not a required check. See [Repository checks](../docs/CI.md) |
| `build_market_diagrams.py` | Renders the five market risk and CVA method maps of `docs/assets/market-risk/` as SVG and PNG; needs CairoSVG 2.9.1, which is not a pinned test dependency |
| `update_expected.py` | Regenerates (or `--check`s) the expected values embedded in `M_Setup` from the independent reference, and rebuilds the template when they change |
| `template_presentation.py` | Restores, in the template's OOXML, the view and print properties and README guide borders that the headless build cannot set; used by `build_template.py` |
| `check_repository.py` | Offline, tracked-file repository checks; the universal maintenance check below. No Excel or regulatory validation |
| `check_issues.py` | Read-only check of the issue metadata and dependency status that issue forms cannot enforce; run by Issue triage. See [Labels](../docs/LABELS.md) |
| `traffic_history.py` | Merges collected traffic API responses into the CSV history and writes traffic alerts; see *Traffic tooling* below |
| `traffic_publish.py` | Publishes the merged traffic history atomically to the data-only `traffic-history` branch; see *Traffic tooling* below |

[Validation](../docs/validation.md) lists their prerequisites (Python 3.12 or newer,
LibreOffice Calc with the Python UNO bridge, the pinned packages in
`tests/requirements.txt` and `.github/requirements.txt`) and how they are used
together. The [workbook guide](../docs/workbook-guide.md#updating-an-existing-workbook)
describes building and updating a workbook on Windows, and
[verifying the round trip](../docs/workbook-guide.md#verifying-the-round-trip).

## Label tooling

Run these from the repository root with Node.js providing built-in fetch
support:

```sh
node .github/scripts/labels-sync.mjs --self-test
node .github/scripts/labels-sync.mjs --mode validate
node .github/scripts/labels-drift.mjs --self-test
```

They run offline fixtures or validate the label catalogue; these commands do
not change live GitHub labels. The scripts stay at their existing authoritative
paths under `.github/scripts/`; do not duplicate them here.
[Labels](../docs/LABELS.md) owns live verification and trusted reconciliation.

## 🧪 Universal maintenance check

Install `.github/requirements.txt`, `.github/requirements-quality.txt` and
`tests/requirements.txt`, then run `python tools/check_repository.py`. It runs
the quality gates (Ruff, mypy and actionlint), checks tracked configuration,
whitespace, local Markdown targets and their anchors, and calls the existing
label scripts and offline maintenance tests. A gate whose tool is missing is
skipped unless `REQUIRE_QUALITY_TOOLS=1`, as in CI. Coverage and
dependency-update policy are in [Repository checks](../docs/CI.md).

## 📈 Traffic tooling

The curl step of
[`traffic-export.yml`](../.github/workflows/traffic-export.yml) collects
aggregate traffic with the GitHub App token, only from trusted main in the
analytics environment. `traffic_history.py` merges the collected responses
into the CSV history and its badge files, offline and without a token, and
writes the title and body of a traffic alert issue for the workflow to open;
`traffic_publish.py` publishes the merged history.
[Traffic history](../docs/TRAFFIC_HISTORY.md) owns the contract and credential
readiness. Their offline tests run without credentials through the universal
gate. The credentialed daily collection has been active
since October 18, 2026 (#23).
Generated output and durable evidence follow
[Repository structure](../docs/REPOSITORY_STRUCTURE.md).
