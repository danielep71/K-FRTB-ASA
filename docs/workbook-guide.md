# 🗂️ Workbook Guide

[![Model: template plus source](https://img.shields.io/badge/model-template%20%2B%20source-0969da)](#source-of-truth)
[![Build: Excel COM](https://img.shields.io/badge/build-Excel%20COM-217346)](#building-the-workbook-windows-excel)
[![Layout: workbook frame](https://img.shields.io/badge/layout-workbook%20frame-1f3864)](#sheets)

This document is authoritative for **how the workbook is built, laid out and
used**. What it calculates is owned by the [methodology](methodology.md);
what is checked by the [validation record](validation.md).

## Numeric cells and numeric text

Use actual Excel numeric cells for amounts and parameters; values entered
with Excel's local decimal separator retain their numeric value. Text cells
have a locale-independent contract: optional sign, decimal point and optional
`E` exponent (for example `0.5`, `.5`, `-1.25e2`). The entire trimmed text
must match and fit a VBA Double. Commas, grouping, currency/type suffixes,
embedded whitespace, trailing junk and overflow are rejected rather than
partly converted. Thus text `0,5` is rejected on every host; a numeric cell
containing one half is supported on every host. Required key/value numeric
parameters reject blank, error and nonnumeric values with key/sheet context.
Every numeric value in a `P_*` table (risk weights, correlation matrices,
liquidity horizons, DRC and RRAO weights, bucket numbers) and every entry of
a `P_General` vertex list is required as well: a blank or nonnumeric value,
an empty list or a blank list entry (`0.25,,1`, a trailing comma) stops the
run and the `Log` names the sheet, table, data row and column, or the list.
Tagged parameter tables reject blank keys in populated rows and completely
blank rows followed by more data before the next tag; blank separator rows
between tables remain valid (#48).
Every run clears the output sheets first and again when it fails, so a
failed run shows "Run failed" on `Results` and leaves neither earlier nor
partial capital behind (#45). **Clear outputs** and a failed run keep the
column headers of every output sheet and put the note in the first data
row, as in a new workbook (#85).

Validation and conversion use the same parser. Blank/error optional values
still use the caller's explicit default; this does not relax required-key
validation. The runtime regression suite covers these contracts (#41).

<a id="source-of-truth"></a>

## 🧩 Source of truth

The exported VBA modules are the source of truth. The macro-free template
`src/workbook/FRTB_Template.xlsx` holds the sheets, styles, default
parameters, sample portfolio and expected results that `M_Setup.BuildWorkbook`
produces, without a VBA project, buttons or personal document properties.
The `.xlsm` workbook is a build output: it is not committed (`/dist/` is
ignored) and is built from the template plus the modules, following the
[template-based build](REPOSITORY_STRUCTURE.md#conditional-excelvba-policy).

The template is generated, never edited by hand: `python3
tools/build_template.py` rebuilds it headless in LibreOffice and normalizes
saved view/print settings and README guide borders through
`tools/template_presentation.py`, and
`tools/update_expected.py` rebuilds it whenever the expected values change.
`tests/test_template.py` fails when the committed template differs from a
fresh build of the modules, in values or layout, or carries VBA, controls or
personal properties.
`tools/Build-Workbook.ps1 -Template` produces the same template through Excel.

```text
src/
  modules/     VBA modules: engine (CORE_*, Option Private Module) and workbook macros (M_*)
  workbook/    macro-free template FRTB_Template.xlsx
tests/
  modules/     VBA test modules imported into development builds (TEST_*)
  *.py         independent reference, reconciliation and tooling tests
tools/         build and sync scripts with their launchers, template builder, LibreOffice runner, VBA lint, expected-value generator
docs/          methodology, workbook guide, validation, assets
```

| Module | Responsibility |
| --- | --- |
| `src/modules/CORE_Config` | Product name, sheet names, layout, palette, risk class codes, components, scenarios |
| `src/modules/CORE_Util` | Sheet access, tagged parameter tables, conversions, table formatting, output placeholders, run log |
| `src/modules/CORE_Params` | Control settings and `P_*` parameters; risk weight and correlation accessors |
| `src/modules/CORE_SBM` | Delta, vega and curvature: validation, netting, bucket and cross-bucket aggregation |
| `src/modules/CORE_DRC` | Default risk charge for non-securitizations |
| `src/modules/CORE_RRAO` | Residual risk add-on |
| `src/modules/M_Main` | `RunFRTB`, `RunFRTB_Silent`, `ClearOutputs`; Results and Out_Summary |
| `src/modules/M_Setup` | `BuildWorkbook`, `FinishWorkbook`, `AddNavigationBars`, `ResetParameters`, `LoadSampleData` |
| `src/modules/M_Dev` | `Dev_ExportModules`: export the modules back to their folders |
| `tests/modules/TEST_Regression` | `Test_RunRegression`, `Test_RunRegression_Silent` |
| `tests/modules/TEST_Numeric` | `Test_NumericContracts_Silent`: numeric validation and conversion contracts |
| `tests/modules/TEST_Acceptance` | `Test_RunAcceptanceInExcel_Silent`, `Test_RunAcceptance_Silent`: the acceptance checks of the unattended check (#274) |
| `tests/modules/TEST_AcceptanceData` | The acceptance checks' data, generated by `tools/build_acceptance_data.py` |

Each module opens with a banner (purpose, public surface, dependencies, state
ownership, error policy, compatibility); `tools/vba_lint.py` checks the banner,
the `CORE_`, `M_` and `TEST_` prefixes and `Option Private Module` in `CORE_`
modules.

<a id="building-the-workbook-windows-excel"></a>

## 🏗️ Building the workbook (Windows, Excel)

1. In Excel, enable **File > Options > Trust Center > Trust Center Settings >
   Macro Settings > Trust access to the VBA project object model**.
2. From the repository root:

   ```powershell
   powershell -ExecutionPolicy Bypass -File tools\Build-Workbook.ps1 -RunTests
   ```

   The script opens a copy of `src\workbook\FRTB_Template.xlsx`, imports
   `src\modules` and `tests\modules`, runs `FinishWorkbook` (sheet buttons
   and reapplication of saved view settings) and `AddNavigationBars` (the bar
   of links on every sheet), the calculation and the regression test, and saves
   `dist\K-FRTB-ASA.xlsm`. `-FromSource` builds every sheet from the modules
   (`BuildWorkbook`) instead of the template; `-Production` leaves out the
   test modules; `-Template` saves a source-built, macro-free template as
   `dist\FRTB_Template.xlsx` for comparison with the committed one; `-Force`
   overwrites an existing file.

   Without a PowerShell prompt, double-click `tools\Build-Workbook.cmd`. It
   asks for a production or a development workbook, and whether to
   overwrite an existing `dist\K-FRTB-ASA.xlsm`. It then runs the same script,
   with `-Production` or with `-RunTests`.
3. Open the workbook and run **Debug → Compile VBAProject**.

Manual alternative: open a copy of `src\workbook\FRTB_Template.xlsx`, open
the VBA editor (Alt+F11), use **File > Import File** for every `.bas` file of
`src\modules` and `tests\modules`, run `FinishWorkbook` and `AddNavigationBars` (Alt+F8), then save
as `.xlsm`.

<a id="updating-an-existing-workbook"></a>

## 🔄 Updating an existing workbook

When only the VBA changed, update a workbook you already built instead of
building a new one. `tools\Sync-FRTB-VBA.ps1` replaces its standard modules
through Excel and leaves the cells, parameters, inputs, buttons, sheets and
calculation settings as they are. It needs the same Trust Center setting as the build. Close the
workbook in Excel first.

* **From the repository checkout:** drag the `.xlsm` onto
  `tools\Sync-FRTB-VBA.cmd`, or pass its path. It asks which workbook you are
  updating, or takes the answer as a second argument (`P` or `D`):
  * **P**roduction imports `src\modules`;
  * **D**evelopment also imports `tests\modules`.
* **A fresh development workbook:** double-click `tools\Sync-FRTB-VBA.cmd`
  without a workbook. It builds one from the template, with the modules of
  `src\modules` and `tests\modules`, as `test-results\K-FRTB-ASA-Development.xlsm`,
  and replaces the one it built last. The folder is ignored by Git. Open it, run
  **Debug → Compile VBAProject**, then the regression test.
* **From a downloads folder:** drag the `.xlsm` onto
  `tools\Sync-FRTB-VBA-Downloads.cmd`. It asks for the folder, by default
  `%USERPROFILE%\Downloads`, and imports only the modules found there. The
  repository files are not changed.
  * Each file is identified by the single `Attribute VB_Name` it declares,
    never by its file name. When several files declare the same name, for
    example `CORE_SBM (1).bas`, the newest is taken.
  * A new module whose name starts with `CORE_`, `M_` or `TEST_` is listed
    separately.
  * The list is shown for confirmation before anything is imported. Any
    other module of the workbook stays as it is.
  * Line endings are converted to CRLF, so modules downloaded from GitHub's
    **Raw** view import correctly.

The script first saves a copy beside the workbook
(`<name>.pre-vba-sync.<time>.xlsm`). It removes the copy after a successful
save and keeps it when the synchronization fails. It stops without saving if
the sheet structure changes. A production synchronization never removes the
test modules of a workbook; to drop them, build a new production workbook.

It exits with 0 on success, 1 on failure (the workbook is not saved) and 2
when canceled (the workbook is not changed).

From PowerShell, the same options are available without the prompts:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Sync-FRTB-VBA.ps1 -WorkbookPath dist\K-FRTB-ASA.xlsm -Mode Development
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Sync-FRTB-VBA.ps1 -WorkbookPath dist\K-FRTB-ASA.xlsm -SourceDir "$env:USERPROFILE\Downloads" -Force
```

`-NoBackup` skips the safety copy and `-Visible` shows Excel while it works.
After a synchronization, open the workbook and run **Debug → Compile
VBAProject**; in a development workbook, run the regression test as well.

A template change needs a new build: a synchronization changes the VBA
only, so a workbook built from an older template keeps that template's
sheets and parameters. `git diff --stat <old-commit> HEAD --
src/workbook/FRTB_Template.xlsx` shows whether the template changed.

<a id="sheets"></a>

## 🎨 Sheets

Every sheet carries the same frame: the product name in gold (`FFC000`) on
a black bar, the title in white on a blue bar (`265A99`), an italic grey
subtitle and tables from column C. Table headers are navy (`1F3864`) on
every sheet, and input cells are cream (`FFF2CC`); PASS and FAIL use green
(`E2EFDA`) and red (`FFC7CE`). Every cell is Aptos and vertically centered,
also after a run or **Clear outputs** rewrites an output sheet, and
gridlines are hidden. Input amounts show two decimals (`#,##0.00`) and
tenors and maturities in years one to three (`0.0##`). A wrapped
description row is as high as its lines need (#206). The palette is defined
once in `CORE_Config`.

The XLSX saves colored tabs, hidden gridlines and frozen panes before any
macro runs. Data rows are 20.1 points high and 10 pt, table headers are 42
points high and 11 pt bold (#351), and input tables use thin gray borders. The README uses dashed internal
separators. Display titles and column captions use readable words; engine
sheet names, parameter codes and column positions remain unchanged.
Initial output sheets show navy headers and a no-results message, without
invented calculation values. Printing uses A4 landscape, one page wide with
unlimited height, repeated heading rows and an Aptos page-number footer.
Output print areas remain automatic so later runs are not clipped to an
empty template's extent.

| Tab | Sheet | Content |
| --- | --- | --- |
| Navy | `Home` | The start sheet, where the workbook opens: what it does, the steps of a run and every sheet with its role and a link |
| Navy | `README` | How to use and the input conventions; the last sheet |
| Gold | `Control` | Run settings (`[SETTINGS]`) and the buttons Run FRTB-SA, Clear outputs and Run regression test (development builds only: `-Production` leaves out `TEST_Regression`) |
| Gold | `In_Delta`, `In_Vega`, `In_Curvature` | Sensitivity inputs (see [methodology](methodology.md#-inputs)) |
| Gold | `In_DRC` | Positions: obligor, bucket (`CORP`, `SOV`, `LGM`), credit quality, seniority, notional, market value, maturity, optional LGD |
| Gold | `In_RRAO` | Category (`EXOTIC`, `OTHER`) and gross notional |
| Gold | `P_General`, `P_GIRR`, `P_CSR`, `P_EQ`, `P_COM`, `P_FX`, `P_DRC`, `P_RRAO` | Regulatory parameters in tagged tables (`[TAG]` heading, header row, data rows), with MAR references |
| Blue | `Results` | Capital by risk class, component and scenario; the binding correlation scenario, SBM, DRC, RRAO and total; the run note, and a note when scenarios tie (#138) |
| Blue | `Out_Summary` | Key/value results used by tests |
| Blue | `Out_RiskClass` | Capital per risk class, component and scenario; `Alternative Sb used` flag |
| Blue | `Out_Buckets` | `Kb`, `Sb`, risk factor count and notes (`ALT_SB`, `UP`/`DOWN`, `OTHER`) per bucket |
| Blue | `Out_DRC` | Net long/short per obligor; HBR and charge per bucket |
| Blue | `Log` | Every rejected input row with the reason, warnings and run status; `Row` and `Reference` name the sheet row and its ID |
| Green | `Test_Expected` | Expected sample results, actual values and PASS/FAIL per item, with the area, description, source and source status of each |

Input sheets are read from row 6 down to the first row whose `ID` and
`Risk class` (second) columns are both blank. Invalid rows are skipped and
logged on `Log`; check the error count on `Results` after every run.

`Log` has five columns: `Level`, `Source`, `Message`, `Row` and `Reference`.
`ERROR` is red and `WARN` amber in the `Level` column. For a rejected input row,
`Row` is its sheet row and `Reference` its `ID`, left blank when the row has
none; for a finding about an obligor, `Reference` is the obligor; for a key
that a parameter table repeats, it is the key. The messages and the `Source`
text are the same as before, so a filter on them keeps working. `Results` does
not say whether its figures still match the inputs: edit an input, then run
again. A line that said so would need a worksheet event or a fingerprint of
every input cell, which a changed value of the same sum would defeat, and I
decided against both (#350).

A key that repeats in the first column of a `Key | Value` table, which the run
ignores with a warning, is highlighted in red where it is typed or pasted, also in
the row directly below the table, which the run reads too. The
rule is conditional formatting on the ten tables that `CORE_Params` reads that
way (`SETTINGS`, `CORR_SCENARIOS`, `VEGA`, `VERTICES`, `GIRR_PARAMS`,
`CSR_PARAMS`, `EQ_PARAMS`, `COM_PARAMS`, `FX_PARAMS` and `DRC_PARAMS`). The keys
are locked on a protected sheet, so a repeat appears only after the sheet is
unprotected or a block is pasted over it. `tools/key_highlights.py` writes the
rules into the template and `M_Setup` applies them when Excel builds or resets
a sheet, with the formula in the language of the Excel that runs it; a workbook
built with `-FromSource` gets them from `M_Setup`. The Excel check
`KEY_HIGHLIGHTS` types a repeat into every table after **Reset parameters** and
checks that it turns red.

### 🧭 Navigation

The workbook opens on `Home`, which lists every sheet once with its role and a
link to it. On every sheet, column A (18 wide, black) holds a bar of dark
entries. START and Home sit at the top, in the rows that stay in view when you
scroll; below them come INPUTS, PARAMETERS, OUTPUTS and TESTS & REFERENCE. A
grey triangle at the left of B1 hides the bar (it points left) and, with the bar
hidden, shows it again (it points right); the **Navigation bar** button of the
ribbon does the same. A strip at the left of each entry has the color of the
sheet's tab; the current sheet is the blue entry with white bold text and an
amber strip, and the others are links. An entry is `262626` with `D9D9D9` Aptos
9 pt text, a section label is amber 7 pt bold and a strip is 2.25 pt wide. The
bar is made of shapes,
which the macro-free template cannot hold: `tools/Build-Workbook.ps1` adds it
through `M_Setup.AddNavigationBars` after `FinishWorkbook`, and the template
save drops it. A workbook you update with `tools\Sync-FRTB-VBA.ps1` keeps its
bar, and one built before the bar existed has none until you run
`AddNavigationBars`. The unattended Excel check `NAVIGATION` reads the markers,
the links and the list on `Home` back in Excel.

The command buttons follow the same pattern, in the title bar of the sheet they
act on. Each is a rounded rectangle with its macro assigned, named `btn_` and
the macro, in Aptos 11 pt bold: navy with white text for **Run FRTB-SA**, light
blue with navy text for the others. It moves with its cells, so it follows the
navigation bar when you hide it. They are **Run FRTB-SA**, **Clear outputs** and **Run
regression test** on `Control`; **Run FRTB-SA** and **Clear outputs** on
`Results`; **Clear outputs** on `Out_Summary`, `Out_RiskClass`, `Out_Buckets`,
`Out_DRC` and `Log`; **Run regression test** on `Test_Expected`. `Load sample
data` and `Reset parameters` have no button, because each overwrites what you
typed; run them from the macro list. `NAVIGATION` also counts the buttons of
every sheet and toggles the bar to read where they sit.

### 🎀 Ribbon tab

`tools\Build-Workbook.ps1 -Ribbon` adds an **FRTB-ASA** tab to the workbook it
saves, with **Run FRTB-SA**, **Clear outputs**, **Run regression test** and a
toggle for the navigation bar, each with a screen tip. The three commands are
disabled while a run is busy. The sheet buttons stay, so the workbook works
where a policy removes custom UI. [The contract](ribbon.md) lists every control,
callback and state, and `tests/test_ribbon.py` checks that the part, the
contract and `M_Main` agree. Without `-Ribbon` the workbook has no tab.

### ✅ Data validation

The template carries data validation rules (#348) that check what you type
into a cell. The run's own validation stays the authority: a pasted value
bypasses a rule, and the run rejects the same rows with the same messages
whatever the rules say.

| Cells | Rule |
| --- | --- |
| `Control` switches | A drop-down of `TRUE` and `FALSE`; blank is accepted |
| `Control` reporting currency | A warning unless the text has three letters |
| `In_*` risk class | A drop-down of `GIRR`, `CSR_NS`, `EQ`, `COM` and `FX` |
| `In_Delta`, `In_Vega`, `In_Curvature` bucket | `GIRR` and `FX`: a three-letter currency. `CSR_NS`, `EQ` and `COM`: a number of 1 or more |
| `In_Delta`, `In_Vega`, `In_Curvature` sub-type | The sub-types the run accepts for the risk class: delta `GIRR` `YIELD`, `INFLATION` or `XCCY`, `CSR_NS` `BOND` or `CDS`, `EQ` `SPOT` or `REPO`; vega the same without `REPO`; curvature `GIRR` `YIELD` only. `COM` takes free text, `FX` none, and a blank takes the default |
| Tenors, maturities, amounts, CVRs, LGD, notionals | Numbers only; tenors, maturities and the RRAO gross notional also zero or more |
| `In_Curvature` references reporting ccy | A drop-down of `TRUE` and `FALSE` |
| `In_DRC` bucket, credit quality, seniority and `In_RRAO` category | A drop-down of the keys in the tables of `P_DRC` and `P_RRAO`. A value that is not there yet only warns, so you can add the key to the table afterwards |
| `P_*` numeric cells | Numbers only. A risk weight below zero and a correlation outside -1 to 1 only warn |

The `ID`, `Qualifier`, `Obligor` and `Comment` columns have no rule: the run
does not read the ID and the comment, and it reports a blank qualifier or
obligor itself. Text lists are case-sensitive in Excel while the run ignores
case, and a bucket number must be typed as a number, not as text.

`tools/input_validation.py` defines the rules once. `tools/build_template.py`
writes them into the template after the headless build, and
`tests/test_input_validation.py` pins them, evaluates their formulas in
LibreOffice and compares their vocabularies with what the run accepts. The
unattended Excel check `DATA_VALIDATION` reads every rule back in Excel. A
workbook built with `-FromSource` has no rules, because it is not made from
the template: build from the template for the workbook you use. **Load sample
data** and **Reset parameters** rewrite their sheets but keep the rules. The rules
reference cells directly, so the workbook defines no names for them, and the
engine keeps reading the tagged tables.

### 🔒 Protection

The sheets and the workbook structure are protected **without a password**
(#349), so that a header, a tag, an output or a sheet name is not overwritten
or changed by accident. It is not a security control: anyone can remove it.

| What | Protection |
| --- | --- |
| Cells in the input fill (cream) | Unlocked on every sheet: the `Control` settings, the five input sheets and the values of the `P_*` sheets |
| Headers, tags, outputs, `Log`, `Test_Expected`, `Home` and `README` | Locked |
| `In_*` sheets | Also allow inserted rows, sorting and the autofilter |
| Workbook structure | Locked: a sheet cannot be added, deleted, renamed or moved |

To change something locked, use **Review > Unprotect Sheet** or **Protect
Workbook**; no password is asked. The commands (**Run FRTB-SA**, **Clear
outputs**, **Run regression test**, **Load sample data**, **Reset
parameters**, the build and the sync) lift the protection while they write and
put it back afterwards, also when they fail, so a protected workbook runs
without unprotecting anything. A workbook made before the protection, or updated by the sync,
has every cell locked: its first command unlocks the cells in the input fill
on `Control`, the input sheets and the `P_*` sheets and then protects it. If you protect a sheet or the workbook with a
password of your own, a command stops with a message that names the sheets:
remove the password and run again.

The committed template carries the protection: `M_Setup` and `CORE_Util`
apply it, `tools/template_presentation.py` writes what the headless build
cannot (the actions of the input sheets and the locked structure), and
`tests/test_template.py` asserts the locked and unlocked cells. A workbook
built with `-FromSource` gets it from the same code. The unattended Excel check
`PROTECTION` reads it back in Excel and tries a write to a header and to an
input cell; typing and filtering by a person are not part of a run.

## ▶️ Running

* **Run FRTB-SA** (`RunFRTB`) recalculates everything and rewrites the output
  sheets and the log.
* **Run regression test** (`Test_RunRegression`) compares the last run with
  `Test_Expected`. It is meaningful only with the sample data and default
  parameters and settings: run `LoadSampleData` and `RunFRTB` first.
* `ResetParameters` rewrites the `P_*` sheets with the defaults.

## 🔁 Changing the VBA

1. Edit the modules in the VBA editor of a development build.
2. Run `Dev_ExportModules` to write them back to `src\modules` and
   `tests\modules`. It finds the folders only when the workbook sits in the
   repository root or in `dist\`. A workbook kept elsewhere, for example one
   updated from a downloads folder, cannot export: move or copy it into
   `dist\` of the checkout first.
3. On a machine with LibreOffice and Python, run the checks in
   [validation](validation.md). If the sample portfolio, parameters or engine
   logic changed, regenerate the expected values with
   `python3 tools/update_expected.py` (which also rebuilds the template) and
   review the diff. A layout change needs `python3 tools/build_template.py`.

<a id="verifying-the-round-trip"></a>

## 🔍 Verifying the round trip

The round trip shows that a workbook holds exactly the source of one commit
(#30). It never writes into the repository, so it can check any workbook,
including one updated from a downloads folder.

1. Build the workbook from the template at the commit to check, in
   development or production mode (see
   [building the workbook](#building-the-workbook-windows-excel)), and record
   the commit.
2. Export its VBA to a new, empty folder outside the checkout:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File tools\Export-FRTB-VBA.ps1 -WorkbookPath dist\K-FRTB-ASA.xlsm -OutDir C:\FRTB\export
   ```

   The exporter opens the workbook read-only, with macros disabled, in a
   separate Excel instance. It exports every component with the VBA editor's
   own Export and closes the workbook without saving. It then checks that the
   workbook's SHA-256 has not changed and writes `export-manifest.json`. The
   manifest records the SHA-256 of the workbook and of every exported file,
   and the type of every component: the comparison takes a document module
   (type 100) from it, because the export of `ThisWorkbook` and the sheets
   carries no `VB_Base` line.
3. Compare the export with the source of that commit, on any machine with
   Python and the checkout:

   ```sh
   python3 tools/compare_vba_export.py --export C:\FRTB\export --mode development --candidate-sha <commit> --workbook dist\K-FRTB-ASA.xlsm
   ```

   The source is read from Git at `<commit>`, not from the working tree.
   * **Production** expects every module of `src/modules`.
   * **Development** expects those and every module of `tests/modules`.

   What the comparison checks:
   * Missing or unexpected modules, and attribute or code differences, fail.
   * Line endings, trailing blank lines and the letter case of code outside
     strings and comments are ignored, because the VBA editor re-cases
     identifiers on import. The re-cased lines are counted.
   * `ThisWorkbook` and the sheet modules are not part of the source; they
     must hold no code beyond `Option` lines.
   * With `--workbook`, the export must match its manifest and the workbook's
     bytes: the binding is `VERIFIED`. A changed file or another workbook
     gives `FAIL`; an export without a manifest stays `UNVERIFIED`.

   The output names the commit, the template SHA-256 and the workbook
   SHA-256, and ends with `RESULT=PASS` or `RESULT=FAIL`. Exit code 0 is
   PASS, 1 is FAIL, 2 means the inputs could not be read. Record the full
   output with the evidence.
