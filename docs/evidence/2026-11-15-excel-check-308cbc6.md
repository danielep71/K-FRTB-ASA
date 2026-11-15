# 🪟 Excel check — November 15, 2026

The third recorded session of K-FRTB-ASA in desktop Excel. It checks the
review of every sheet (#206) and the fixes of the check itself (#324) and of
the build (#327). I ran the unattended check `tools\Check-FRTB-Excel.cmd`
(#274) on my Windows host, on November 15, 2026, from 18:11:34 to 18:15:40
(UTC+01:00). Every check passed: `RESULT=PASS`, 24 of 24.

| | |
| --- | --- |
| Candidate | `release/v0.2.0` at `308cbc6b144943ce17d5e369a4a7c35d5ea70290` |
| Template | `src/workbook/FRTB_Template.xlsx`, SHA-256 `ea430d540c13ec667b37d3fc6d8b5d8fd4eb98699ce707b12432f7c87b1dc3d8` |
| Validator outcome | **PASS**: `tools/check_excel_evidence.py` on the recorded run, on the host |
| Operator | Daniele Penza |

On this host, the workbook builds from the template and from source,
compiles, passes its regression and acceptance checks and round-trips. The
build also reports the Trust Center error, and the modules synchronize into
an existing workbook. The run is not regulatory validation, and it says
nothing about methodology (#29) or performance.

Unlike the [check of November 12](2026-11-12-excel-check-1badef1.md), this
record does not list the digests of the bundle's logs. The recorder validated
the bundle on the host, and the bundle stays with me, because it records the
computer name and local paths. This record rests on the check's summary,
its logs and its pictures.

## 🖥️ Host

The same host as the runs of
[October 31](2026-10-31-excel-run-323260b.md) and
[November 12](2026-11-12-excel-check-1badef1.md): Excel 16.0, build 20430,
64-bit, comma decimal separator. Microsoft Print to PDF was already the
default printer.

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

The two compiles are the evidence for #327: the build of `3c20005` had
stopped on a compile error in `M_Setup`.

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
| `DISPLAYED_VALUES` | PASS: 14,430 cells on 22 sheets after a sample run; none shows `####` (#316) |
| `NO_DATES` | PASS: none of these cells holds a date (#316) |
| `SHEET_VIEW` | PASS: 22 sheets without gridlines, each with the tab color of its role (#316) |
| `INVALID_ROW_VARIANT` | PASS: total 3,917,301.09, 22 errors and 1 warning, as the independent reference |

The run of November 15 on `7aaa47e` counted 26,369 cells. Fewer cells are in
the used range now, because **Clear outputs** formats whole columns rather
than single cells (#326).

## 📸 Sheet pictures (#206, #324)

`SHEET_PICTURES` passed: 44 of 44 pictures, every sheet after **Clear
outputs** and after a sample run. None is blank. On `7aaa47e`, four pictures
had come out blank and still counted; the check now retries an empty paste
(#324).

The pictures after the sample run show the fixes of #326 in Excel:
- every cell is Aptos and vertically centered;
- output headers are navy;
- input amounts show `5.000.000,00`, and tenors `1,0`;
- the `P_CSR` and `P_EQ` descriptions are complete;
- `Results` has its rule under "Capital requirement", the binding scenario
  right-aligned, and the binding total in navy.

They are in [`docs/assets/screenshots/`](../assets/screenshots/), named
`check-308cbc6-<sheet>.png`:
- [README](../assets/screenshots/check-308cbc6-readme.png),
  [Control](../assets/screenshots/check-308cbc6-control.png),
  [Results](../assets/screenshots/check-308cbc6-results.png);
- inputs:
  [In_Delta](../assets/screenshots/check-308cbc6-in-delta.png),
  [In_Vega](../assets/screenshots/check-308cbc6-in-vega.png),
  [In_Curvature](../assets/screenshots/check-308cbc6-in-curvature.png),
  [In_DRC](../assets/screenshots/check-308cbc6-in-drc.png),
  [In_RRAO](../assets/screenshots/check-308cbc6-in-rrao.png);
- outputs:
  [Out_Summary](../assets/screenshots/check-308cbc6-out-summary.png),
  [Out_RiskClass](../assets/screenshots/check-308cbc6-out-riskclass.png),
  [Out_Buckets](../assets/screenshots/check-308cbc6-out-buckets.png),
  [Out_DRC](../assets/screenshots/check-308cbc6-out-drc.png),
  [Log](../assets/screenshots/check-308cbc6-log.png);
- parameters:
  [P_General](../assets/screenshots/check-308cbc6-p-general.png),
  [P_GIRR](../assets/screenshots/check-308cbc6-p-girr.png),
  [P_CSR](../assets/screenshots/check-308cbc6-p-csr.png),
  [P_EQ](../assets/screenshots/check-308cbc6-p-eq.png),
  [P_COM](../assets/screenshots/check-308cbc6-p-com.png),
  [P_FX](../assets/screenshots/check-308cbc6-p-fx.png),
  [P_DRC](../assets/screenshots/check-308cbc6-p-drc.png),
  [P_RRAO](../assets/screenshots/check-308cbc6-p-rrao.png);
- [Test_Expected](../assets/screenshots/check-308cbc6-test-expected.png).

Each picture shows at most the first 60 rows and 26 columns of its sheet.

## 🏗️ Builds and the Trust Center

| Check | Result |
| --- | --- |
| `BUILD_FROM_SOURCE` | PASS: `-FromSource -RunTests`, exit code 0 |
| `BUILD_TEMPLATE` | PASS: `-Template` saved a macro-free template, SHA-256 `865e0ff43d3d89c6a00ba9f7984ee7b446397ce61f7601d26933dd563b454510` |
| `TRUST_CENTER_ERROR` | PASS: with the VBA project access off, the build exits 1 with a message that names the setting; the access was restored to 1 |

`TRUST_CENTER_ERROR` failed on `7aaa47e` only because the console had
wrapped the message in two lines. The check now matches it across line
breaks (#324).

I compared the Excel-built template from the check with the committed one
using `tools/compare_templates.py`. The differences are the ones the
[check of November 12](2026-11-12-excel-check-1badef1.md)
explains, and they come from the host that saved the file:
- the right margin column of each sheet is 11.53 wide in the committed
  template and 9.00 in Excel's;
- √2 is stored with 15 significant digits by LibreOffice and as the full
  double by Excel;
- the Boolean cells of `Control` and `In_Curvature` are a number with a
  TRUE/FALSE format in LibreOffice and a Boolean in Excel.

## 🔄 Synchronization (#72)

| Check | Result |
| --- | --- |
| `SYNC_PRODUCTION` | PASS: compiled; total 3,877,792.81 as expected; calculation mode kept |
| `SYNC_DEVELOPMENT` | PASS: compiled; total 3,877,792.81; regression 52 of 52; calculation mode kept |
| `SYNC_FOLDER` | PASS: `CORE_Config` replaced from a folder of downloaded modules |
| `SYNC_FAILURE` | PASS: the synchronization fails while the workbook is open elsewhere, keeps its safety copy and leaves the workbook unchanged |
| `LAUNCHER_Sync-FRTB-VBA.cmd` | PASS, exit code 0 |
| `LAUNCHER_Sync-FRTB-VBA-Downloads.cmd` | PASS, exit code 0 |

## 🔁 Earlier attempts the same day

- **On `7aaa47e`:** 23 of 24 passed. The check gave the pictures and the
  findings of #206, and it showed the two faults of the check fixed by #324.
- **On `3c20005`, with the #206 fixes:** the build stopped on a compile
  error in `M_Setup`, fixed by #327. The run was stopped there.

## ⏭️ Not covered by this run

- Printing: the page setup is set by the build, but nothing was printed.
- Other Excel versions, 32-bit Excel and other regional settings: this host
  uses the comma decimal separator only.
- Regulatory correctness and performance on large portfolios.
