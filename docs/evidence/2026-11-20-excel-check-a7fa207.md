# 🪟 Excel check — November 20, 2026, on `a7fa207`

The fifth recorded session of K-FRTB-ASA in desktop Excel. It checks the start
sheet and the navigation bar (#347), the protection of the sheets and of the
workbook structure (#349) and the 11 pt table headers (#351, #359). I ran the
unattended check `tools\Check-FRTB-Excel.cmd` (#274) on my Windows host, on
November 20, 2026, from 16:55 (UTC+01:00). Every check passed: `RESULT=PASS`,
27 of 27. Its sheet pictures show every header and the navigation bar as the
template intends.

| | |
| --- | --- |
| Candidate | `release/v0.3.0` at `a7fa2075c2e4cc4ad0bfb457afdc3adda10dd4e8` |
| Template | `src/workbook/FRTB_Template.xlsx`, SHA-256 `430ea566ee4b204f41ef72a0cd63eba4b6bb1f827142b353ba1e0d547b37fbaa` |
| Validator outcome | **PASS**: `tools/check_excel_evidence.py` on the recorded run, on the host |
| Operator | Daniele Penza |

On this host, the workbook builds from the template and from source,
compiles, passes its regression and acceptance checks with the sheets and the
structure protected, and round-trips. The build also reports the Trust Center
error, and the modules synchronize into an existing workbook. The run is not
regulatory validation, and it says nothing about methodology (#29) or
performance.

Like the [check of November 15](2026-11-15-excel-check-308cbc6.md), this
record does not list the digests of the bundle's logs: the bundle stays with
me, because it records the computer name and local paths.

An earlier run the same day, on `9390788`, passed 26 of 26 with the new
`NAVIGATION` check and showed two capitals that broke mid-word, "SectorGroup"
on `P_CSR` and "RHO_NAME" on `P_EQ`, and a navigation heading cut to "TESTS
AND". #359 fixes them, and this run shows that it does. That bundle is not a
record of its own.

## 🖥️ Host

The same host as the earlier runs: Excel 16.0, comma decimal separator.

## 🚦 Recorded run

| Stage | Result |
| --- | --- |
| `build_development` | PASS, exit code 0 |
| `compile_development` | PASS, **Debug → Compile VBAProject** with no error |
| `regression` | PASS, through `-RunTests` |
| `round_trip_development` | PASS: 13 modules, 0 differences, binding verified |
| `build_production` | PASS, exit code 0 |
| `compile_production` | PASS, **Debug → Compile VBAProject** with no error |
| `round_trip_production` | PASS: 9 modules, 0 differences, binding verified |

## ✅ Acceptance checks in Excel

`TEST_Acceptance.Test_RunAcceptanceInExcel_Silent`, in a copy of the
development workbook:

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
| `DISPLAYED_VALUES` | PASS: 16,262 cells on 23 sheets after a sample run; none shows `####` |
| `NO_DATES` | PASS: none of these cells holds a date |
| `SHEET_VIEW` | PASS: 23 sheets without gridlines, each with the tab color of its role |
| `DATA_VALIDATION` | PASS: 92 rules on the input and parameter sheets, as in the template, also after **Load sample data** and **Reset parameters** |
| `NAVIGATION` | PASS: 23 sheets, each with a marker and 22 links, all listed on `Home` (#347) |
| `PROTECTION` | PASS: 23 sheets and the structure protected without a password; 8,596 unlocked cells, all in the input fill; a script's write to a header is refused (#349) |
| `INVALID_ROW_VARIANT` | PASS: total 3,917,301.09, 22 errors and 1 warning, as the independent reference |

`NAVIGATION` reads the shapes and the links back after **Load sample data** and
**Reset parameters**, and counts the entries of `Home`. `PROTECTION` runs after
every command above: the run, **Clear outputs**, **Load sample data**, **Reset
parameters**, the regression test and the acceptance report all opened their
write window and closed it, so the sheets and the structure are protected again
when the check reads them. It compares every cell's lock with the input fill, reads the
actions each sheet allows, and tries a script's write to a header on
`Out_Summary` (refused) and to an input cell on `In_Delta` (accepted). It does
not type into a cell, filter, sort or insert a row as a person does, and it
does not click a link.

## 📸 Sheet pictures (#351, #347)

`SHEET_PICTURES` passed: 46 of 46 pictures, every sheet after **Clear outputs**
and after a sample run, with the sheets protected: the export lifts a sheet's
protection to draw its picture, and the next command protects it again. I read
the pictures after the run:
- **Headers:** every 11 pt header fits on two lines at most and none breaks
  mid-word, including "SectorGroup" on `P_CSR`, "RHO_NAME" on `P_EQ`, the three
  of `In_Vega` and `In_Curvature`, and "Reference" on `P_CSR`, `P_EQ` and
  `P_COM`.
- **Navigation bar:** each sheet shows the bar in column A with its own
  entry as the marker, the five headings read "START", "INPUTS",
  "PARAMETERS", "OUTPUTS" and "TESTS & REFERENCE", and the bar covers no
  content. `Home` lists the 23 sheets once, with their roles.
- **Sheets:** all 23 pictures after a sample run are in the repository as
  `check-a7fa207-<sheet>.png`:
  [Home](../assets/screenshots/check-a7fa207-home.png),
  [Control](../assets/screenshots/check-a7fa207-control.png),
  [Results](../assets/screenshots/check-a7fa207-results.png);
  inputs
  [In_Delta](../assets/screenshots/check-a7fa207-in-delta.png),
  [In_Vega](../assets/screenshots/check-a7fa207-in-vega.png),
  [In_Curvature](../assets/screenshots/check-a7fa207-in-curvature.png),
  [In_DRC](../assets/screenshots/check-a7fa207-in-drc.png),
  [In_RRAO](../assets/screenshots/check-a7fa207-in-rrao.png);
  outputs
  [Out_Summary](../assets/screenshots/check-a7fa207-out-summary.png),
  [Out_RiskClass](../assets/screenshots/check-a7fa207-out-riskclass.png),
  [Out_Buckets](../assets/screenshots/check-a7fa207-out-buckets.png),
  [Out_DRC](../assets/screenshots/check-a7fa207-out-drc.png),
  [Log](../assets/screenshots/check-a7fa207-log.png);
  parameters
  [P_General](../assets/screenshots/check-a7fa207-p-general.png),
  [P_GIRR](../assets/screenshots/check-a7fa207-p-girr.png),
  [P_CSR](../assets/screenshots/check-a7fa207-p-csr.png),
  [P_EQ](../assets/screenshots/check-a7fa207-p-eq.png),
  [P_COM](../assets/screenshots/check-a7fa207-p-com.png),
  [P_FX](../assets/screenshots/check-a7fa207-p-fx.png),
  [P_DRC](../assets/screenshots/check-a7fa207-p-drc.png),
  [P_RRAO](../assets/screenshots/check-a7fa207-p-rrao.png);
  [Test_Expected](../assets/screenshots/check-a7fa207-test-expected.png) and
  [README](../assets/screenshots/check-a7fa207-readme.png).

The pictures of the first run on `9390788` are not added: they show the headers
and the heading that #359 fixes.

## 🏗️ Builds and the Trust Center

| Check | Result |
| --- | --- |
| `BUILD_FROM_SOURCE` | PASS: `-FromSource -RunTests`, exit code 0 |
| `BUILD_TEMPLATE` | PASS: `-Template` saved a macro-free template, SHA-256 `8206e5ffd271419537981a24971fd0664d02687e05d9a63774e5d68339adb61c` |
| `TRUST_CENTER_ERROR` | PASS: with the VBA project access off, the build exits 1 with a message that names the setting; the access was restored to 1 |

`-FromSource` and `-Template` run with the protection on: the build opens a
write window to delete the shapes of a template and protects the workbook when
it closes.

I compared the Excel-built template with the committed one using
`tools/compare_templates.py`; the host could not run the comparison, because
`openpyxl` was missing, so I ran it on the received bundle. The differences are
the ones the [check of November 12](2026-11-12-excel-check-1badef1.md)
explains, and they come from the host that saved the file:
- the right margin column of each sheet is 11.53 wide in the committed template
  and 9.00 in Excel's;
- the Boolean cells of `Control` and `In_Curvature` are a number with a
  TRUE/FALSE format in LibreOffice and a Boolean in Excel.

The comparison does not look at data validation or protection.

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

- Typing into a validated cell, and into an unlocked cell of a protected
  sheet: the run reads the rules and the locks back, and a person still has to
  see Excel accept and refuse entries.
- Filtering, sorting and inserting rows on a protected input sheet, as a person
  does it, and a click on a navigation link or on a button of a protected
  sheet.
- The unlocking of an older workbook that the synchronization updated: the
  synchronization runs here on workbooks that already carry the unlocked cells.
- Printing: the page setup is set by the build, but nothing was printed.
- Other Excel versions, 32-bit Excel and other regional settings: this host uses
  the comma decimal separator only.
- Regulatory correctness and performance on large portfolios.
