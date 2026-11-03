# 🧪 Validation

[![Reconciliation: LibreOffice vs reference](https://img.shields.io/badge/reconciliation-LibreOffice%20vs%20reference-2ea44f)](#automated-checks)
[![Excel evidence: recorded](https://img.shields.io/badge/Excel%20evidence-recorded%20323260b-2ea44f)](evidence/2026-10-31-excel-run-323260b.md)
[![Parameters: BIS checked](https://img.shields.io/badge/parameters-BIS%20checked-1f3a5f)](methodology.md#bis-check)

This document states what is checked automatically, how, and what is not
checked. Treat any result as unvalidated for regulatory use.

<a id="automated-checks"></a>

## ✅ Automated checks

| Check | Command | What it shows |
| --- | --- | --- |
| VBA static lint | `python3 tools/vba_lint.py` | Module banner, `Option Explicit` and `VB_Name` in every module, `CORE_`/`M_` names in `src/modules` and `TEST_` names in `tests/modules`, `Option Private Module` in core, no undeclared identifiers, no duplicate declarations, no cross-module use of private procedures, ASCII-only source |
| Hand-calculated tests | `python3 -m pytest tests/test_reference.py` | The reference reproduces small examples worked out from the MAR21-MAR23 formulas: correlation scenarios, GIRR tenor/curve/inflation/basis correlations, alternative `Sb`, other-sector buckets, CSR cross-bucket correlations, FX crosses, vega and curvature correlations and risk factors, curvature selection and `psi`, DRC JTD, maturity scaling, seniority netting, HBR, RRAO |
| Parameter spot check | `tests/test_vba_engine.py::test_parameter_sheets_match_regulatory_defaults` | The `P_*` sheets written by `M_Setup` equal an independently typed list in `tests/regulatory_defaults.py` |
| VBA engine reconciliation | `python3 -m pytest tests/test_vba_engine.py` | The VBA modules, executed headless in LibreOffice, reproduce every `Out_RiskClass`, `Out_Buckets`, `Out_Summary` and `Out_DRC` value of the independent Python reference, on the sample portfolio and on a variant with all discretions enabled and invalid rows, each of which must be logged; empty inputs give zero capital |
| Numeric input contract | `python3 -m pytest tests/test_numeric_contract.py` | Valid numeric text converts as expected and invalid text is rejected in the VBA and the reference; a required key, parameter-table value or vertex-list entry that is blank or not a number stops the run with its location named and leaves every output sheet empty; **Clear outputs** and a failed run restore the template's output headers and note (#41, #45, #85) |
| In-workbook regression test | run inside the reconciliation tests | `Test_RunRegression` passes all 52 items embedded by `M_Setup` |
| Expected values current | `python3 tools/update_expected.py --check` | The values embedded in `M_Setup` equal the reference results for the sample |
| Sync script logic | `python3 -m pytest tests/test_sync_script.py` | The functions of `tools/Sync-FRTB-VBA.ps1`, run offline in PowerShell 7: module identity from `Attribute VB_Name`, folder selection and newest copy, CRLF conversion, mode answers, module sets equal to the build sets; the launchers never echo a path directly. The PowerShell tests are skipped when `pwsh` is not installed (the hosted runner has it); no test runs Excel |
| Export comparison logic | `python3 -m pytest tests/test_compare_vba_export.py` | `tools/compare_vba_export.py` on synthetic exports of the source at `HEAD`: an exact or re-cased export passes. These fail: changed code, comments, attributes or string case; missing, unexpected or test modules in production; document modules holding code; a changed file or another workbook against the manifest. A renamed workbook is a note. The exporter's format and read-only open match what the tool reads. No test runs Excel |
| Excel evidence validator | `python3 -m pytest tests/test_excel_evidence.py`, `python3 tools/check_excel_evidence.py --check-policy` | The validator accepts a consistent synthetic record and rejects these: wrong sources or candidate; missing or foreign `CANDIDATE=` lines; tampered, missing, escaping, symlinked or reused logs; stale regression counts; comparison logs with the wrong mode, template, result or binding; dependency violations; environment errors. It reports failed, not-run, timed-out and unavailable runs. The policy's regression count equals the expected values `M_Setup` embeds. No test runs Excel, and a synthetic record is not evidence |
| Template current | `python3 tools/build_template.py --check`, `tests/test_template.py` | The committed template equals a fresh build of the modules in values and layout (cell styles, number formats, explicitly set column widths and row heights, merged ranges, tab colors, frozen panes, gridlines, autofilters and print settings; automatic heights depend on the host's fonts and are not compared), carries no VBA project, controls or personal properties, and a build from it passes the regression test and reconciles with the reference. Saved-presentation assertions check the intended style independently of rebuild equality; view/print normalization must preserve content and be idempotent. |

Run everything with:

```sh
python3 -m pip install -r tests/requirements.txt -r .github/requirements.txt
python3 tools/vba_lint.py
REQUIRE_LIBREOFFICE=1 python3 -m pytest -v tests
```

LibreOffice Calc and the Python UNO bridge (`python3-uno` on Debian/Ubuntu)
are required for the reconciliation tests; without them those tests are
skipped unless `REQUIRE_LIBREOFFICE=1` is set. The `Prototype checks`
workflow runs the same commands on pull requests into `main` and `release/**`
that change `src/`, `tests/`, `tools/`, `.github/requirements.txt` or the
workflow itself. The universal `Repository integrity` check runs on every
pull request; other workflows, such as the label checks, have path filters
of their own.

The separate universal [Repository integrity](CI.md) job covers repository
maintenance and its offline tests. Its PyYAML dependency is installed from
the CI manifest above when pytest discovers the shared tests; it is not an
engine dependency or evidence of Excel execution.

### How the LibreOffice run works

[`tools/lo_runner.py`](../tools/lo_runner.py) inserts the modules of
`src/modules` and `tests/modules` into an empty Calc document (or the
template) with `Option VBASupport 1`, calls a probe function appended to every module
(LibreOffice does not report compile errors otherwise), runs the requested
macros and saves the result as `.xlsx`.

<a id="not-validated"></a>

## ⚠️ Not validated

* **Excel, beyond one recorded run.** No automated check here runs Excel.
  [One run](evidence/2026-10-31-excel-run-323260b.md) on `323260b` is
  recorded and validated (#28). On one host (Windows 11, Excel 16.0 64-bit,
  comma decimal separator) it shows:
  * the development and production builds of `tools\Build-Workbook.ps1`;
  * **Debug → Compile VBAProject** with no error in both;
  * the regression test, 52 of 52;
  * both round trips through `tools/Export-FRTB-VBA.ps1` and
    `tools/compare_vba_export.py`;
  * the sample run, Clear outputs, a failed run and the `RefReportingCcy`
    cases, checked by eye.

  A change after that commit needs a new run. Other Excel versions, 32-bit
  Excel and other regional settings are untested.
* **The sync script** `tools/Sync-FRTB-VBA.ps1` and its `.cmd` launchers
  have not been executed in Excel (#72). `tests/test_sync_script.py` runs
  the script's file selection, name checks, line-ending conversion and module
  sets offline in PowerShell 7; the Excel import, backup and save are
  untested.
* **Not yet run in Excel:**
  * the invalid-row variant with every discretion on;
  * numeric text typed with each decimal separator;
  * the `-FromSource` and `-Template` builds;
  * the build's Trust Center error path;
  * `Dev_ExportModules`.
* **Regulatory correctness.** The reference and the VBA share the same
  reading of MAR21-MAR23, so their agreement does not prove that reading is
  right. The [BIS check](methodology.md#bis-check), the hand-calculated
  tests and the parameter spot checks tie both to the regulatory text as
  transcribed; the
  [open interpretations](methodology.md#interpretations-and-simplifications-to-review)
  still need a decision.
* **Real portfolios and performance.** Only the synthetic sample and small
  test cases have been run; the engine uses simple loops and has not been
  timed on large inputs.
