# 🪟 Excel check — November 20, 2026, on `15bcf82`

The sixth recorded session of K-FRTB-ASA in desktop Excel, and the first with one
failed check. It checks the new look of the navigation bar and the buttons
(#347, #363), the key highlights and the Log columns (#350, #362), the catalogue
of the 52 regression items (#352, #364) and the protection of the sheets (#349).
I ran the unattended check `tools\Check-FRTB-Excel.cmd` (#274), full profile, on my
Windows host, on November 20, 2026, from 18:52 (UTC+01:00). `RESULT=FAIL (1
failed)`: 27 of 28 passed. The one that failed, `KEY_HIGHLIGHTS`, found a real
defect of the workbook, which #368 fixed; the section "After the run" says how it
was confirmed.

| | |
| --- | --- |
| Candidate | `release/v0.3.0` at `15bcf82d263bdfa27521905f5f6aeb4ad7b6e99b` |
| Template | `src/workbook/FRTB_Template.xlsx`, SHA-256 `40693831ec960cef8f30b8214e546268967f1c6f73dbcf2970804f1f10beed15` |
| Validator outcome | **PASS**: `tools/check_excel_evidence.py` on the recorded run, on the host |
| Operator | Daniele Penza |

On this host, the workbook builds from the template and from source, compiles,
passes its regression test (52 items) and 15 of its 16 acceptance checks with the
sheets and the structure protected, and round-trips. The build also reports the
Trust Center error, and the modules synchronize into an existing workbook. The run
is not regulatory validation, and it says nothing about methodology (#29) or
performance.

Like the [check of November 15](2026-11-15-excel-check-308cbc6.md), this record does
not list the digests of the bundle's logs: the bundle stays with me, because it
records the computer name and local paths.

## 🖥️ Host

The same host as the earlier runs: Excel 16.0, comma decimal separator.

## 🚦 Recorded run

| Stage | Result |
| --- | --- |
| `build_development` | PASS, exit code 0 |
| `compile_development` | PASS, **Debug → Compile VBAProject** with no error |
| `regression` | PASS, through `-RunTests`: 52 items |
| `round_trip_development` | PASS: 13 modules, 0 differences, binding verified |
| `build_production` | PASS, exit code 0 |
| `compile_production` | PASS, **Debug → Compile VBAProject** with no error |
| `round_trip_production` | PASS: 9 modules, 0 differences, binding verified |

## ✅ Acceptance checks in Excel

`TEST_Acceptance.Test_RunAcceptanceInExcel_Silent`, in a copy of the development
workbook:

| Check | Result |
| --- | --- |
| `SAMPLE_RUN` | PASS: total 3,877,792.81, 0 errors, 0 warnings |
| `FAILED_RUN` | PASS: `Results` C6 shows "Run failed (see the Log sheet): Invalid numeric parameter in [GIRR_RW] data row 1, column 2 (P_GIRR): blank" |
| `CLEAR_OUTPUTS` | PASS: headers in row 5 and the note in row 6 on `Results`, the `Out_*` sheets and `Log` |
| `NUMERIC_TEXT` | PASS: rejected as text `0,5`, `1.000,5`, `1,5`; accepted `0.5`, `-1.25e2` and the number 0.5 |
| `UNDEFINED_SUB_TYPES` | PASS: the nine rows are rejected with their messages, and the results are the sample's |
| `FROZEN_PANES` | PASS: 23 sheets frozen as in the template |
| `AUTOFILTERS` | PASS: five input sheets filtered as in the template; `In_Delta` filtered on GIRR shows its 10 GIRR rows only |
| `CONTROL_DESCRIPTIONS` | PASS: the Description column is 125.14 wide; its texts need 121.43 |
| `DISPLAYED_VALUES` | PASS: 16,526 cells on 23 sheets after a sample run; none shows `####` |
| `NO_DATES` | PASS: none of these cells holds a date |
| `SHEET_VIEW` | PASS: 23 sheets without gridlines, each with the tab color of its role |
| `DATA_VALIDATION` | PASS: 92 rules on the input and parameter sheets, as in the template, also after **Load sample data** and **Reset parameters** |
| `NAVIGATION` | PASS: 23 sheets, each with a marker and 22 links, all listed on `Home`; the buttons as set (#347) |
| `KEY_HIGHLIGHTS` | **FAIL**: after **Reset parameters**, nine tables had no rule: `P_General`!C7:C9, C12:C14 and C28:C32, `P_GIRR`!C20:C29, `P_CSR`!C28:C33, `P_EQ`!C23:C24, `P_COM`!C21:C23, `P_FX`!C7:C11 and `P_DRC`!C26:C28 (#350) |
| `PROTECTION` | PASS: 23 sheets and the structure protected without a password; 8,596 unlocked cells, all in the input fill; a script's write to a header is refused (#349) |
| `INVALID_ROW_VARIANT` | PASS: total 3,917,301.09, 22 errors and 1 warning, as the independent reference |

`KEY_HIGHLIGHTS` is new in this run (#350). It runs **Reset parameters** and then
reads the conditional formats of the ten `Key | Value` tables the run reads: the
nine on the `P_*` sheets and the one on `Control`. `Control` is not reset and
keeps the template's rule, so it was not among the failures. **Reset parameters**
clears the formats of the `P_*` sheets, the rules included, and `M_Setup` then
failed to write them again.

## 🔁 After the run

- **A wrong reading.** My first reading of the failure (#367) was that the check
  read a property that Excel does not fill, and that the red fill showed. That
  was wrong. #367 made the check look at what Excel shows: a repeated key typed
  into the row below a table must be red.
- **The quick profile on `0d7fbf9`.** With that check, the quick profile (the build
  with the regression test and the acceptance checks, not evidence) failed the
  same way: a repeated key was not red on any of the nine `P_*` tables, while it
  was on `Control`. The workbook was at fault, not the check.
- **The cause and the fix.** `FormatConditions.Add` reads its formula in the language
  and with the separators of the Excel that runs it, and this host is an Italian
  Excel, so the English formula did not parse and `M_Setup.MarkDuplicateKeys`
  failed silently. #368 writes the formula into a scratch cell, reads
  `FormulaLocal` and adds the rule with it.
- **A freeze.** With #368, the acceptance checks froze Excel on the runner. My
  reading of the code: the scratch cell was the last cell of the sheet, which
  stretches the used range to the last row and column, and three checks walk
  `UsedRange` cell by cell. #384 moved the cell to `BH1`, and the next run ended.
- **The quick profile on `fcf25c3`, on my runner after #384:** `RESULT=PASS`, in
  about 90 seconds. The build with the regression test (52 items) and the 16
  acceptance checks passed, `KEY_HIGHLIGHTS` among them: ten tables, a key is
  plain, and the same key in the row below is red. The quick profile is not
  evidence, because it skips the stages that follow the acceptance checks; the
  acceptance checks themselves are the same ones.

## 📸 Sheet pictures (#347, #363)

`SHEET_PICTURES` passed: 46 of 46 pictures, every sheet before and after a sample
run, with the sheets protected. The `PICTURES` of the sample run, of the failed run
and of the cleared outputs passed as well. I read the pictures after the run:
- **Navigation bar and buttons:** each sheet shows the dark bar in column A, black
  with dark entries, amber section labels and a strip in the tab color of each
  entry; the current sheet is a blue entry, `START` is the last section, and the
  bar covers no content. `Home` lists the 23 sheets once, with their roles.
- **Sheets:** all 22 pictures after a sample run are in the repository as
  `check-15bcf82-<sheet>.png`: [Home](../assets/screenshots/check-15bcf82-home.png), [Control](../assets/screenshots/check-15bcf82-control.png), [Results](../assets/screenshots/check-15bcf82-results.png);
  inputs [In_Delta](../assets/screenshots/check-15bcf82-in-delta.png), [In_Vega](../assets/screenshots/check-15bcf82-in-vega.png), [In_Curvature](../assets/screenshots/check-15bcf82-in-curvature.png),
  [In_DRC](../assets/screenshots/check-15bcf82-in-drc.png), [In_RRAO](../assets/screenshots/check-15bcf82-in-rrao.png); outputs [Out_Summary](../assets/screenshots/check-15bcf82-out-summary.png),
  [Out_RiskClass](../assets/screenshots/check-15bcf82-out-riskclass.png), [Out_Buckets](../assets/screenshots/check-15bcf82-out-buckets.png), [Out_DRC](../assets/screenshots/check-15bcf82-out-drc.png),
  [Log](../assets/screenshots/check-15bcf82-log.png); parameters [P_General](../assets/screenshots/check-15bcf82-p-general.png), [P_GIRR](../assets/screenshots/check-15bcf82-p-girr.png), [P_CSR](../assets/screenshots/check-15bcf82-p-csr.png),
  [P_EQ](../assets/screenshots/check-15bcf82-p-eq.png), [P_COM](../assets/screenshots/check-15bcf82-p-com.png), [P_FX](../assets/screenshots/check-15bcf82-p-fx.png), [P_DRC](../assets/screenshots/check-15bcf82-p-drc.png), [P_RRAO](../assets/screenshots/check-15bcf82-p-rrao.png)
  and [README](../assets/screenshots/check-15bcf82-readme.png).

The picture of `Test_Expected` is not added: its `Source` column was cut in this
run, and #367 widened it. It is not covered by a picture until the next run.

## 🏗️ Builds and the Trust Center

| Check | Result |
| --- | --- |
| `BUILD_FROM_SOURCE` | PASS: `-FromSource -RunTests`, exit code 0 |
| `BUILD_TEMPLATE` | PASS: `-Template` saved a macro-free template, SHA-256 `265c48aa951a5f54f7a8be2d3880f3a61eab1353f59201b24129161738a7190e` |
| `TRUST_CENTER_ERROR` | PASS: with the VBA project access off, the build exits 1 with a message that names the setting; the access was restored to 1 |

I compared the Excel-built template with the committed one using
`tools/compare_templates.py`; the host could not run the comparison, because
`openpyxl` was missing, so I ran it on the received bundle. The differences are
the ones the [check of November 12](2026-11-12-excel-check-1badef1.md) explains,
and they come from the host that saved the file:
- the right margin column of each sheet is 11.53 wide in the committed template and
  9.00 in Excel's;
- the Boolean cells of `Control` and `In_Curvature` are a number with a TRUE/FALSE
  format in LibreOffice and a Boolean in Excel;
- the two square roots of 2 (`P_GIRR`!D28 and `P_FX`!D9) are stored with 15
  significant digits by LibreOffice and in full by Excel.

The comparison does not look at data validation, conditional formats or protection.

## 🔄 Synchronization (#72)

| Check | Result |
| --- | --- |
| `SYNC_PRODUCTION` | PASS: compiled; total 3,877,792.81 as expected; calculation mode kept |
| `SYNC_DEVELOPMENT` | PASS: compiled; total 3,877,792.81; regression 52 of 52; calculation mode kept |
| `SYNC_FOLDER` | PASS: `CORE_Config` replaced from a folder of downloaded modules |
| `SYNC_FAILURE` | PASS: the synchronization fails while the workbook is open elsewhere, keeps its safety copy and leaves the workbook unchanged |
| `LAUNCHER_Sync-FRTB-VBA.cmd` | PASS, exit code 0 |
| `LAUNCHER_Sync-FRTB-VBA-Downloads.cmd` | PASS, exit code 0 |

## ⏭️ Not covered by this run

- A full run that passes every check on a commit after #368 and #384: the later runs
  are quick runs.
- The optional ribbon (#353): the check does not build it.
- Typing into a validated cell, and into an unlocked cell of a protected sheet, as a
  person does it; filtering, sorting and inserting rows; a click on a navigation
  link or on a button.
- The user-password path of the protection: the checks run without a password.
- The unlocking of an older workbook that the synchronization updated.
- Printing, other Excel versions, 32-bit Excel and other regional settings: this
  host uses the comma decimal separator only.
- Regulatory correctness and performance on large portfolios.
