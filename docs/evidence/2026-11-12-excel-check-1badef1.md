# 🪟 Excel check — November 12, 2026

The second recorded session of K-FRTB-ASA in desktop Excel, for #28, #72
and #85. I ran the unattended check `tools\Check-FRTB-Excel.cmd` (#274) on
my Windows host, on November 12, 2026, from 19:27:46 to 19:35:51
(UTC+01:00). Every check passed: `RESULT=PASS`, 21 of 21.

| | |
| --- | --- |
| Candidate | `release/v0.2.0` at `1badef14c469da9778ca34183eae90a371a59735` |
| Template | `src/workbook/FRTB_Template.xlsx`, SHA-256 `1b0c1df7ad5c53598a356cf175bee3210970bdd5586ed05b057fb8a7ea1b82aa` |
| Validator outcome | **PASS**: `tools/check_excel_evidence.py` on the recorded run, on the host and again on the received bundle |
| Operator | Daniele Penza |

The record binds the run to the candidate and to its logs by SHA-256. It is
evidence that, on this host, the workbook builds from the template and from
source, compiles, passes its regression and acceptance checks, round-trips,
reports the Trust Center error, and synchronizes into an existing workbook.
It is not regulatory validation, and it says nothing about methodology (#29)
or performance.

## 🖥️ Host

The same host as the [run of October 31, 2026](2026-10-31-excel-run-323260b.md).

| Item | Recorded |
| --- | --- |
| Windows | Windows 11 Enterprise 25H2 (build 26200.9711), AMD64 |
| Excel | 16.0, build 16.0.20430.20180, 64-bit |
| Separators | decimal `,`, list `;` |
| Macro setting | Disable VBA macros with notification, set by policy |
| VBA project access | On, as found; no Trust Center setting changed, except by the error-path check, which restored it |
| References | Visual Basic For Applications; Microsoft Excel 16.0 Object Library; OLE Automation; Microsoft Office 16.0 Object Library |

The recorder read these values from Excel and the registry into
`session.log`. The computer name, the user name and the local paths are in
the bundle but are left out here.

## 🚦 Recorded run

The check's first stage is `tools\Record-FRTB-ExcelRun.ps1`, from 19:27:53 to
19:29:15.

| Stage | Result | Log | Log SHA-256 |
| --- | --- | --- | --- |
| `build_development` | PASS, exit code 0 | `build-development.log` | `ba35cd026fb867082af94568e67fea40cc9bb830a7820281f2c5dd4726aa9a02` |
| `compile_development` | PASS, **Debug → Compile VBAProject** with no error | `session.log` | `ca77dbd42825932a7cbde5d2b25bc8ae22afb5fd4d0babe3a869a320250dcab1` |
| `regression` | PASS, 52 of 52 items | `build-development.log` | as above |
| `round_trip_development` | PASS: 13 modules, 0 differences, binding verified | `round-trip-development.log` | `3520c371fe5060ced6edbac4c6467444c9b0eb1ff0e63fc70249660cfa719ec6` |
| `build_production` | PASS, exit code 0 | `build-production.log` | `23c356d9c1559f7f638d665e940c1acdacfffeca5d02a8793532ffe003d6a31f` |
| `compile_production` | PASS, **Debug → Compile VBAProject** with no error | `session.log` | as above |
| `round_trip_production` | PASS: 9 modules, 0 differences, binding verified | `round-trip-production.log` | `f4a17bfac6b2a8d6cc1fa5b7f335dc5cfa96d9b96ca9b2ffeba5252ef425ee0d` |

| Workbook | SHA-256 |
| --- | --- |
| Development build | `b2d681e88ae02835f1a75db7ad729ee2425fcfce793d640ca23fbb13a896f5cb` |
| Production build | `2a9807a4c654af0e7f230634c70a261dd1741a050995d67fdebb7e5b26673b16` |

`host.json` has SHA-256
`9a175528127892cc82e1fc5917b8e859aee5243e561b4f894793f991621faa95`. The
comparison ignored 242 (development) and 180 (production) lines that differ
only in the letter case of code: the VBA editor re-cases identifiers on
import.

## ✅ Acceptance checks in Excel

`TEST_Acceptance.Test_RunAcceptanceInExcel_Silent`, in a copy of the
development workbook (#274, #276):

| Check | Result |
| --- | --- |
| `SAMPLE_RUN` | PASS: total 3,877,792.81, 0 errors, 0 warnings |
| `FAILED_RUN` | PASS: `Results` C6 shows "Run failed (see the Log sheet): Invalid numeric parameter in [GIRR_RW] data row 1, column 2 (P_GIRR): blank" (#85) |
| `CLEAR_OUTPUTS` | PASS: headers in row 5 and the note in row 6 on `Results`, the `Out_*` sheets and `Log` (#85) |
| `NUMERIC_TEXT` | PASS: rejected as text `0,5`, `1.000,5`, `1,5`; accepted `0.5`, `-1.25e2` and the number 0.5 |
| `UNDEFINED_SUB_TYPES` | PASS: the nine rows are rejected with their messages, and the results are the sample's (#245) |
| `FROZEN_PANES` | PASS: 22 sheets frozen as in the template (#203) |
| `AUTOFILTERS` | PASS: five input sheets filtered as in the template; `In_Delta` filtered on GIRR shows its 10 GIRR rows only |
| `CONTROL_DESCRIPTIONS` | PASS: the Description column is 125.14 wide; its texts need 121.43 (#203) |
| `INVALID_ROW_VARIANT` | PASS: total 3,917,301.09, 22 errors and 1 warning, as the independent reference (#28) |

The sample results and the variant's expected values come from the
independent Python reference, through `tests/modules/TEST_AcceptanceData.bas`.

## 🏗️ Builds and the Trust Center

| Check | Result |
| --- | --- |
| `BUILD_FROM_SOURCE` | PASS: `-FromSource -RunTests`, exit code 0, regression 52 of 52 |
| `BUILD_TEMPLATE` | PASS: `-Template` saved a macro-free template, SHA-256 `5785436215ac9582b4aead3501928199da2dc5ea7bd23075cff121103312f91d` |
| `TRUST_CENTER_ERROR` | PASS: with the VBA project access off, the build exits 1 with a message that names the setting; the access was restored to 1 |

The host had no `openpyxl`, so the check wrote "not compared" for the
templates. I compared the Excel-built template from the bundle with the
committed one using `tools/compare_templates.py`. The sheets, the values,
the layout and the frozen panes are the same; the differences come from the
host that saved the file:

* The right margin column of each of the 22 sheets is 11.53 wide in the
  committed template and 9.00 in Excel's: each host's default width.
* √2 in `P_GIRR!D28` and `P_FX!D9`: LibreOffice stores 15 significant digits
  (`1.4142135623731`), Excel the full double (`1.4142135623730951`).
* The Boolean cells `Control!D8:D11` and `In_Curvature!J13`: LibreOffice
  stores TRUE and FALSE as a number with the format `"TRUE";"TRUE";"FALSE"`,
  Excel as a Boolean with its own format. The values read are the same.

The committed template stays as it is.

## 🔄 Synchronization (#72)

| Check | Result |
| --- | --- |
| `SYNC_PRODUCTION` | PASS: 9 modules replaced; compiled; total 3,877,792.81 as expected; calculation mode kept |
| `SYNC_DEVELOPMENT` | PASS: 13 modules replaced; compiled; total 3,877,792.81; regression 52 of 52; calculation mode kept |
| `SYNC_FOLDER` | PASS: `CORE_Config` replaced from a folder of downloaded modules |
| `SYNC_FAILURE` | PASS: with the workbook open in another Excel, the synchronization fails, keeps its safety copy and leaves the workbook unchanged |
| `LAUNCHER_Sync-FRTB-VBA.cmd` | PASS, exit code 0 |
| `LAUNCHER_Sync-FRTB-VBA-Downloads.cmd` | PASS, exit code 0 |

## 🔁 Earlier attempts the same day

Two earlier runs of the check needed fixes before this one passed:

* **At `b0222de`:**
  * The `-Template` build hung at its save. The build macros turn Excel
    alerts back on, so the hidden Excel waited on the question about saving
    without the VB project. Fixed by #303 (#302). My first reading, the
    printer, was wrong; its change, #296 (#295), is harmless and stays.
  * Both launcher checks reported FAIL with an empty exit code, although
    their logs show success. Fixed by #298 (#297).
* **At `d1898f0`:** the launchers passed, and the template build hung again,
  before #303.

The other checks passed in all three runs. Excel also showed "Waiting for
printer connection" during this run; it did not stop the run. Since #305
(#304), the check uses Microsoft Print to PDF as the default printer for the
run.

## 📸 Screenshots

The check took pictures in the development workbook of this run:

* [Results after the sample run](../assets/screenshots/check-1badef1-results.png)
  and [Control](../assets/screenshots/check-1badef1-control.png), with the
  #203 layout;
* after a failed run:
  [Results](../assets/screenshots/check-1badef1-failed-results.png),
  [Out_Summary](../assets/screenshots/check-1badef1-failed-out-summary.png) and
  [Log](../assets/screenshots/check-1badef1-failed-log.png);
* [Results after Clear outputs](../assets/screenshots/check-1badef1-cleared-results.png).

## 🧾 Bundle

The bundle `excel-run-1badef1` holds `session.log`, the four logs above,
`host.json` and both exports with their `export-manifest.json`. The
validator's reports are in `excel-run-1badef1-validation`. The check's own
folder `excel-check-1badef1-20261112-192746` holds `summary.txt`,
`check.log`, `acceptance.txt`, the build, synchronization and launcher logs,
the pictures and the Excel-built template. I keep them, because they record
the computer name and local paths. The digests above let anyone holding them
check that they are the ones validated here.

## ⏭️ Not covered by this run

* The review of every sheet's formatting (#206).
* Other Excel versions, 32-bit Excel and other regional settings: this host
  uses the comma decimal separator only.
* Regulatory correctness and performance on large portfolios.
