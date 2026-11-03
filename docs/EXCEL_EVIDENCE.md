# 🪟 Excel evidence

[![Execution: manual](https://img.shields.io/badge/execution-manual-1D76DB)](#manual-procedure)
[![Binding: exact SHA](https://img.shields.io/badge/binding-exact%20SHA-217346)](#what-a-record-binds)
[![Validator: offline](https://img.shields.io/badge/validator-offline-6f42c1)](#validating-a-bundle)

This document defines how a run of K-FRTB-SA in desktop Excel is recorded, so
that it can be tied to one exact commit and checked (#82). It is
authoritative for:

* the evidence record;
* the manual Windows/Excel procedure;
* the validator `tools/check_excel_evidence.py`.

The [workbook guide](workbook-guide.md) owns building, updating and the round
trip. [Validation](validation.md) owns what the other checks show. Issues #28
and #30 say which Excel results the SA prototype (v0.2.0) needs.

> [!IMPORTANT]
> The validator never runs Excel. It checks three things:
> * a record is complete and consistent;
> * its sources are the ones stored in Git at the candidate commit;
> * its logs say what the record says.
>
> It cannot prove what a person did. A `NOT_RUN` or `UNAVAILABLE` stage, a
> LibreOffice run, a static check or a synthetic bundle is never an Excel
> pass.

Execution is **manual**: an operator runs Excel interactively on Windows.
There is no automated Excel job.

<a id="what-a-record-binds"></a>

## 🔗 What a record binds

An evidence bundle is a folder outside the checkout holding `host.json` and
the logs it cites.

| Item | Bound to |
| --- | --- |
| Candidate | The full 40-character commit SHA. Every build log, every manual log and every comparison log must name it |
| Sources | Every module of `src/modules` and `tests/modules`, and the template `src/workbook/FRTB_Template.xlsx`, at that commit, with the SHA-256 of the bytes stored in Git |
| Policy | `.github/excel-evidence-policy.json` **at that commit**: the repository, the required references, the regression item count and the ordered stages with their dependencies |
| Environment | Excel version, build and bitness; the Windows version and architecture; the decimal and list separators; references; macro policy and VBA-project access as found; any Trust Center change |
| Logs | The SHA-256 of each cited log; a log changed after the record was written fails |

Source digests are of the bytes stored in Git (LF), not of the CRLF files in
a Windows checkout, so do not compute them with `Get-FileHash`. Print them
with:

```sh
python tools/check_excel_evidence.py --candidate-sha FULL_SHA --inventory
```

The policy changes together with the source. When the number of expected
values that `M_Setup` embeds changes, `regression_items` changes in the same
pull request. `python tools/check_excel_evidence.py --check-policy`, run by
`tests/test_excel_evidence.py`, fails otherwise.

<a id="stages"></a>

## 🧭 Stages

| Stage | Kind | Depends on | Log | A pass needs |
| --- | --- | --- | --- | --- |
| `build_development` | build | none | `build-development.log` | one `CANDIDATE=` line and the `Saved ...` line of `Build-Workbook.ps1` |
| `compile_development` | manual | `build_development` | `session.log` | one `CANDIDATE=` line; **Debug → Compile VBAProject** with no error |
| `regression` | regression | `compile_development` | `build-development.log` | one `Regression test: Result: PASS (52 items)` line; `items` equal to it; 52 is `regression_items` in the policy |
| `round_trip_development` | compare | `build_development` | `round-trip-development.log` | the output of `compare_vba_export.py`: the candidate, `development` mode, the template digest of the inventory, `RESULT=PASS` and `binding=VERIFIED`; `workbook_sha256` equal to its workbook line |
| `build_production` | build | none | `build-production.log` | as for the development build |
| `compile_production` | manual | `build_production` | `session.log` | as for the development compile |
| `round_trip_production` | compare | `build_production` | `round-trip-production.log` | as for the development round trip, in `production` mode |

The other checks of #28 have no machine-readable log. These are the Excel-only
presentation, the invalid-row variant and the regional-settings cases. Record
them in `session.log`, and in the issue with their screenshots; the validator
does not judge them.

<a id="recorded-run"></a>

## ▶️ Recorded run

`tools\Record-FRTB-ExcelRun.cmd` records the run in one step. Use desktop
Excel for Windows, with Git and Python 3 on PATH, and turn on *Trust access to
the VBA project object model* (**File → Options → Trust Center → Trust Center
Settings → Macro Settings**).

1. Check out the candidate and confirm the tree is clean:

   ```powershell
   git fetch origin
   git switch --detach FULL_SHA
   ```

2. Close Excel and double-click `tools\Record-FRTB-ExcelRun.cmd`. It writes
   the bundle `test-results\excel-run-SHORTSHA\` (ignored by Git):
   * `session.log`: the candidate, the start and finish times, the operator
     and the environment, read from Excel and the registry;
   * the build, compile, export and comparison of both workbooks, with the
     logs of the [stages](#stages); the compile runs **Debug → Compile
     VBAProject** through the VBA editor and records the references;
   * `host.json`, written by `tools/record_excel_run.py` from the logs only.

   It then runs the validator and writes the reports next to the bundle, in
   `test-results\excel-run-SHORTSHA-validation\`. A failing stage keeps its
   log, and the stages that depend on it are `NOT_RUN`. If a compile error is
   found, Excel shows it in a message box: note it, click **OK**, and the run
   goes on.
3. Do the remaining #28 checks by hand, in `dist\K-FRTB-SA-development.xlsm`,
   with their screenshots.
4. Zip the bundle and the validation folder, attach them to #28 with the
   results of step 3, and add the dated record under `docs/evidence/`.

If you turned on the VBA project access only for the run, start it from
PowerShell and say so:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Record-FRTB-ExcelRun.ps1 -TrustChange "VBA project access turned on for the run and off after it"
```

The launcher exits 0 when the record is a valid pass, 1 for any other
outcome and 2 when the record could not be written. The manual procedure
below gives the same bundle by hand.

<a id="manual-procedure"></a>

## 🧑‍💻 Manual procedure

Use desktop Excel for Windows, with Git, Python 3 and PowerShell on the same
machine. Run the commands in PowerShell from the checkout.

1. **Pick the candidate** and confirm the tree is clean:

   ```powershell
   git fetch origin
   git switch --detach FULL_SHA
   git status
   $sha = git rev-parse HEAD
   ```

2. **Create the bundle folder** outside the checkout, for example
   `..\frtb-evidence-SHORTSHA\`. Start `session.log` there in Notepad and save
   it as UTF-8. Its first line is `CANDIDATE=FULL_SHA`, then the start time
   with its UTC offset.
3. **Record the environment as found.** In `session.log`, note:
   * **File → Account → About Excel**: version, build and bitness;
   * **winver** and **Settings → System → About**: the Windows edition,
     version and architecture;
   * **Control Panel → Region → Additional settings**: the decimal and list
     separators;
   * **File → Options → Trust Center → Trust Center Settings**: the macro
     setting, and whether *Trust access to the VBA project object model* is
     on.

   The build script needs that access. If you turn it on for the run, write
   that in `session.log`, turn it off again at the end, and record it as
   `trust_changes: true` with `trust_changes_detail`. Change no other setting.
4. **Build the development workbook** with its log:

   ```powershell
   & { "CANDIDATE=$sha"; powershell -NoProfile -ExecutionPolicy Bypass -File tools\Build-Workbook.ps1 -RunTests -Force } *>&1 | Out-File -Encoding utf8 ..\frtb-evidence-SHORTSHA\build-development.log
   ```

   The log shows every imported module and the regression result
   (`Regression test: Result: ...`). It ends with `Saved ...dist\K-FRTB-SA.xlsm`.
5. **Compile.** Open `dist\K-FRTB-SA.xlsm` and run **Debug → Compile
   VBAProject**. Then check **Tools → References**: it must list only the
   four default references of the policy. Write both results in
   `session.log`, then close the workbook **without saving**.
6. **Round trip the development workbook:**

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File tools\Export-FRTB-VBA.ps1 -WorkbookPath dist\K-FRTB-SA.xlsm -OutDir ..\frtb-evidence-SHORTSHA\export-development
   python tools\compare_vba_export.py --export ..\frtb-evidence-SHORTSHA\export-development --mode development --candidate-sha $sha --workbook dist\K-FRTB-SA.xlsm | Out-File -Encoding utf8 ..\frtb-evidence-SHORTSHA\round-trip-development.log
   ```

7. **Build, compile and round trip the production workbook** the same way:
   * build with `-Production -OutputPath dist\K-FRTB-SA-production.xlsm -Force`
     and log to `build-production.log`;
   * compile it and record the result in `session.log`;
   * export it to `export-production` and compare it with `--mode production`,
     logging to `round-trip-production.log`.
8. **Do the remaining #28 checks** by hand and note them in `session.log`.
   Then write the finish time and close Excel.
9. **Hash the logs**, after their last edit:

   ```sh
   python -c "import glob,hashlib;[print(hashlib.sha256(open(f,'rb').read()).hexdigest(), f) for f in sorted(glob.glob('*.log'))]"
   ```

   Run this in the bundle folder.
10. **Write `host.json`** in the bundle from the [record template](#record-template),
    with the `--inventory` output and the observed values.
11. **Validate** as in [validating a bundle](#validating-a-bundle).
    * Attach the zipped bundle (the logs, `host.json` and the exports) and the
      summary to #28.
    * Add a dated Markdown record under `docs/evidence/` that cites them.
    * Never commit a workbook.

Keep a failing log as it is: never replace it with a later passing run. If a
step was not done, mark its stage `NOT_RUN`. If Excel could not be run at
all, write an [unavailable record](#unavailable-record). Never fill a field
with a value that was not observed.

<a id="record-template"></a>

## 🧾 Record template

Replace every value with what was observed. This shape is not evidence.

```json
{
  "schema_version": 1,
  "repository": "danielep71/K-FRTB-SA",
  "candidate_sha": "FULL_40_CHARACTER_SHA",
  "execution": "manual",
  "availability_reason": null,
  "started_at": "2026-10-20T09:00:00+02:00",
  "finished_at": "2026-10-20T10:30:00+02:00",
  "operator": "Name, workstation",
  "environment": {
    "excel_version": "16.0",
    "excel_build": "Version 2609 (Build 16.0.00000.00000)",
    "office_bitness": "64-bit",
    "os": "Windows 11 Pro 24H2 (build 26100.0000)",
    "os_architecture": "x64",
    "decimal_separator": ",",
    "list_separator": ";",
    "references": [
      "Visual Basic For Applications",
      "Microsoft Excel 16.0 Object Library",
      "OLE Automation",
      "Microsoft Office 16.0 Object Library"
    ],
    "macro_policy": "As found, for example: Disable VBA macros with notification",
    "vba_project_access": "As found, for example: On",
    "trust_changes": false
  },
  "sources": "PASTE THE --inventory OUTPUT HERE (a JSON array)",
  "stages": {
    "build_development": {"status": "PASS", "detail": "Build-Workbook.ps1 -RunTests -Force", "log": {"path": "build-development.log", "sha256": "DIGEST"}},
    "compile_development": {"status": "PASS", "detail": "Debug > Compile VBAProject: no error", "log": {"path": "session.log", "sha256": "DIGEST"}},
    "regression": {"status": "PASS", "detail": "Test_RunRegression_Silent through -RunTests", "log": {"path": "build-development.log", "sha256": "DIGEST"}, "items": {"total": 52, "failures": 0}},
    "round_trip_development": {"status": "PASS", "detail": "Export and comparison", "log": {"path": "round-trip-development.log", "sha256": "DIGEST"}, "workbook_sha256": "WORKBOOK_DIGEST_FROM_THE_LOG"},
    "build_production": {"status": "PASS", "detail": "Build-Workbook.ps1 -Production", "log": {"path": "build-production.log", "sha256": "DIGEST"}},
    "compile_production": {"status": "PASS", "detail": "Debug > Compile VBAProject: no error", "log": {"path": "session.log", "sha256": "DIGEST"}},
    "round_trip_production": {"status": "PASS", "detail": "Export and comparison", "log": {"path": "round-trip-production.log", "sha256": "DIGEST"}, "workbook_sha256": "WORKBOOK_DIGEST_FROM_THE_LOG"}
  }
}
```

Rules the validator applies:

**Record**
* `candidate_sha` must equal the commit given to the validator.
* `sources` must equal the `--inventory` of that commit.
* `repository` must equal the policy's repository.
* The stages must be exactly the policy's stages.

**Stage statuses**
* A stage is `PASS`, `FAIL`, `NOT_RUN` or `TIMEOUT`, with a nonempty
  `detail`.
* A `NOT_RUN` stage has `"log": null`, and a regression that did not run has
  `"items": null`.
* A stage may run only after the stages it depends on passed.

**Logs**
* Build, manual and regression logs hold exactly one `CANDIDATE=` line, and
  it names the candidate.
* Comparison logs name the candidate on their `Candidate:` line.
* Each build and each comparison has its own log. The same run cannot be
  cited twice, even with blank lines or trailing spaces added. The
  regression may cite the development build log, which prints its result.
* A passing build log shows `Saved ...`; a failed or timed-out one does not.
* A regression's `items` must equal its result line. That line's item count
  must equal `regression_items` of the policy at the candidate, so counts of
  an older commit fail. A `TIMEOUT` regression shows no result and claims no
  items.
* A comparison must name its mode and the candidate's template digest. A
  pass needs `RESULT=PASS` and `binding=VERIFIED`; a failure needs
  `RESULT=FAIL`. A `TIMEOUT` comparison shows no `RESULT=` line. It may
  have stopped before printing its `Candidate:`, `Mode:` and `Template:`
  lines, but any of them it did print must match.
* Log paths must be relative, stay inside the bundle and not pass through a
  symbolic link. Logs may be UTF-8, with or without a byte-order mark, or
  UTF-16, as PowerShell 5.1 redirection writes them. A log that cannot be
  decoded is invalid, and so is any malformed record: the validator reports
  `EVIDENCE_INVALID` rather than stopping.

**Environment and time**
* `environment.os` must start with `Windows`.
* `office_bitness` must be `32-bit` or `64-bit`.
* `references` must be exactly the policy's references, in any order.
* Every other environment string must be nonempty.
* `trust_changes: true` needs a nonempty `trust_changes_detail`.
* `started_at` and `finished_at` must carry a UTC offset, and the finish
  cannot precede the start.

<a id="unavailable-record"></a>

When Excel was not run, use the same identity and time fields with
`"execution": "unavailable"` and a nonempty `availability_reason`. Set
`operator` and `environment` to `null`, `"sources"` to `[]` and `"stages"` to
`{}`.

<a id="validating-a-bundle"></a>

## 🚦 Validating a bundle

From the checkout:

```sh
python tools/check_excel_evidence.py --candidate-sha FULL_SHA \
  --evidence ../frtb-evidence-SHORTSHA/host.json \
  --output test-results/excel-evidence.json --summary test-results/excel-evidence.md
```

The candidate commit must be present in the clone; the checkout does not need
to be on it. Reports must be written outside the bundle. Exit code 0 means
the record is a consistent pass. Exit code 1 means a non-passing or invalid
record. Exit code 2 means a command or output error.

| Outcome | Meaning |
| --- | --- |
| `PASS` | A complete, consistent record in which every stage passed |
| `BUILD_FAILED` | A build failed |
| `COMPILE_FAILED` | A compile failed |
| `TEST_FAILED` | The regression test reported a failure |
| `ROUND_TRIP_FAILED` | An export differed from the source |
| `EXECUTION_TIMEOUT` | A stage was abandoned as hung |
| `INCOMPLETE` | At least one stage did not run |
| `UNAVAILABLE` | No Excel run was attempted |
| `EVIDENCE_INVALID` | A missing, malformed, contradictory or stale-source record or log |

A record can carry several outcomes, for example `COMPILE_FAILED,
INCOMPLETE`. Keep non-passing bundles too, until the finding is resolved.

## 🧱 Limits

* Source and log digests are identity controls. They do not prove what a
  workbook contained, and they do not authenticate the operator.
* The `CANDIDATE=` line is written by the recorder, or typed in a manual run. It stops a log of one
  commit from being cited, by mistake, for another; it does not stop a
  deliberate forgery.
* A pass covers the recorded host only: its Excel version, bitness and
  regional settings. It does not validate methodology (#29), regulatory
  correctness or performance.
* The manual checks of #28 recorded in `session.log` are not judged by the
  validator.
* Keep logs free of credentials, client data and unredacted screenshots; see
  [Security](../SECURITY.md).

`tests/test_excel_evidence.py` tests the validator with synthetic bundles in
a temporary Git repository. Those tests are not Excel evidence.
