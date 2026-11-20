# 🪟 Excel check — November 20, 2026

The fourth recorded session of K-FRTB-ASA in desktop Excel. It checks the data
validation of the input and parameter sheets (#348) and the 11 pt table
headers (#351). I ran the unattended check `tools\Check-FRTB-Excel.cmd`
(#274) on my Windows host, on November 20, 2026, from 10:55:15 to 11:09:33
(UTC+01:00). Every check passed: `RESULT=PASS`, 25 of 25. Its sheet pictures
showed four headers that did not fit; #356 fixes them, and a new run records
the result.

| | |
| --- | --- |
| Candidate | `release/v0.3.0` at `27eb55af76aa2ffce5d19f951bcaddb11b08c454` |
| Template | `src/workbook/FRTB_Template.xlsx`, SHA-256 `4e54ad4f69dc2726e56b3f0831f7eb3f580d57d5efdea81e3aa358b4799d08b0` |
| Validator outcome | **PASS**: `tools/check_excel_evidence.py` on the recorded run, on the host |
| Operator | Daniele Penza |

On this host, the workbook builds from the template and from source,
compiles, passes its regression and acceptance checks and round-trips. The
build also reports the Trust Center error, and the modules synchronize into an
existing workbook. The run is not regulatory validation, and it says nothing
about methodology (#29) or performance.

Like the [check of November 15](2026-11-15-excel-check-308cbc6.md), this
record does not list the digests of the bundle's logs: the bundle stays with
me, because it records the computer name and local paths.

## 🖥️ Host

The same host as the earlier runs: Excel 16.0, build 20527, comma decimal
separator. Microsoft Print to PDF was already the default printer.

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
| `FROZEN_PANES` | PASS: 22 sheets frozen as in the template |
| `AUTOFILTERS` | PASS: five input sheets filtered as in the template; `In_Delta` filtered on GIRR shows its 10 GIRR rows only |
| `CONTROL_DESCRIPTIONS` | PASS: the Description column is 125.14 wide; its texts need 121.43 |
| `DISPLAYED_VALUES` | PASS: 14,430 cells on 22 sheets after a sample run; none shows `####` |
| `NO_DATES` | PASS: none of these cells holds a date |
| `SHEET_VIEW` | PASS: 22 sheets without gridlines, each with the tab color of its role |
| `DATA_VALIDATION` | PASS: 92 rules on the input and parameter sheets, as in the template, also after **Load sample data** and **Reset parameters** (#348) |
| `INVALID_ROW_VARIANT` | PASS: total 3,917,301.09, 22 errors and 1 warning, as the independent reference |

`DATA_VALIDATION` is new. It reads every rule back in Excel by type, alert
style and error title, after it has run **Load sample data** and **Reset
parameters**, which clear their sheets: they now clear contents and formats and
leave the rules. It does not compare the formulas, which Excel returns in the
language of the host, and it does not type into a cell, so what Excel does with
a rejected entry is not recorded here.

## 📸 Sheet pictures (#206, #351)

`SHEET_PICTURES` passed: 44 of 44 pictures, every sheet after **Clear outputs**
and after a sample run. I read the pictures after the run, and they show that my
estimate for #351 was wrong: an 11 pt header on two lines did not fit four
columns.

| Header | Sheet and column | What the picture shows |
| --- | --- | --- |
| Option maturity (years) | `In_Vega` H | The second line runs under the filter arrow |
| Underlying maturity (years) | `In_Vega` I | The same |
| References reporting ccy | `In_Curvature` J | The same |
| Reference | `P_CSR`, `P_EQ`, `P_COM` E | The word breaks after "Referenc" |

Every other header, and the 10 pt data rows, fit. #356 widens the columns and
adds a test that estimates the wrap of every header; it was merged after this
run, so the next unattended check on `release/v0.3.0` shows whether the headers
now fit. The pictures of this run are not added to the gallery, because they
show those headers.

## 🏗️ Builds and the Trust Center

| Check | Result |
| --- | --- |
| `BUILD_FROM_SOURCE` | PASS: `-FromSource -RunTests`, exit code 0 |
| `BUILD_TEMPLATE` | PASS: `-Template` saved a macro-free template, SHA-256 `8372c7c500badfe395d70cc42d0edb365aee33950c7f235bd9b2c19568c99721` |
| `TRUST_CENTER_ERROR` | PASS: with the VBA project access off, the build exits 1 with a message that names the setting; the access was restored to 1 |

I compared the Excel-built template with the committed one using
`tools/compare_templates.py`; the host could not run the comparison, because
`openpyxl` was missing, so I ran it on the received bundle. The differences are
the ones the [check of November 12](2026-11-12-excel-check-1badef1.md)
explains, and they come from the host that saved the file:
- the right margin column of each sheet is 11.53 wide in the committed template
  and 9.00 in Excel's;
- √2 is stored with 15 significant digits by LibreOffice and as the full double
  by Excel;
- the Boolean cells of `Control` and `In_Curvature` are a number with a
  TRUE/FALSE format in LibreOffice and a Boolean in Excel.

The comparison does not look at data validation. The Excel-built template has
none of the 92 rules, because only the committed template carries them: a
workbook built with `-FromSource` has no rules, as the
[workbook guide](../workbook-guide.md#-data-validation) says.

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

- Typing into a validated cell: the run reads the rules back, and a person
  still has to see Excel accept and refuse entries.
- The header widths of #356, which were merged after the run.
- Printing: the page setup is set by the build, but nothing was printed.
- Other Excel versions, 32-bit Excel and other regional settings: this host uses
  the comma decimal separator only.
- Regulatory correctness and performance on large portfolios.
