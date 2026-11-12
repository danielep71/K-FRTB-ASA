# 🛡️ Security policy

[![Stage: prototype](https://img.shields.io/badge/stage-prototype-1D76DB)](README.md)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

## 📦 Supported scope

| Material | Security maintenance status |
| --- | --- |
| Current `main`, the active `release/v0.2.0` and the milestone 3 integration branch `release/v0.3.0`: source and maintenance tooling | Reports accepted, and triaged by me |
| Historical commits, including the deleted setup branch `release/v0.0.1` and the superseded `release/v0.1.0`, or unmaintained task branches | No separate backport commitment; report issues that also affect current maintained source |
| Calculation engine, workbook, or packaged releases | None published or supported yet |

I maintain this prototype alone; it is not a managed security service.
There is no guaranteed response or remediation time and no bug-bounty program.

## 🧭 What the code reads, writes and deletes

| Entry point | Reads | Writes or deletes |
| --- | --- | --- |
| Workbook macros: **Run FRTB-SA**, **Clear outputs**, **Run regression test** | `Control`, the `In_*` and `P_*` sheets, `Test_Expected` | Clear and rewrite `Results`, the `Out_*` sheets, `Log` and the result cells of `Test_Expected`. No file, network or registry access |
| `M_Dev.Dev_ExportModules` (development workbook only) | The workbook's standard modules | Overwrites the `.bas` files of `src/modules` and `tests/modules` next to or above the workbook |
| `tools/Build-Workbook.ps1` | The template and the modules of the checkout | Writes the output workbook (`dist\` by default), replacing it with `-Force` |
| `tools/Sync-FRTB-VBA.ps1` and its launchers | The repository modules, or the `.bas` files of a chosen folder | Replaces the standard modules of the target workbook and saves it. A safety copy `*.pre-vba-sync.<time>.xlsm` is deleted after a successful save and kept after a failure |
| `tools/Export-FRTB-VBA.ps1` | A workbook, opened read-only with macros disabled | Writes the exported components and `export-manifest.json` into a new or empty folder |
| `tools/Record-FRTB-ExcelRun.cmd` | The checkout; Excel, Office and Windows settings from Excel and the registry; the user and computer name | Builds both workbooks in `dist\` and writes the bundle under `test-results\` |
| `tools/Check-FRTB-Excel.cmd` and `.ps1` | The checkout and the branch; Excel, Office and Windows settings, as the recorder reads them | Fetches the branch and switches the checkout to its commit; can turn the VBA project access on, and off for the Trust Center check, restoring your setting; makes Microsoft Print to PDF the default printer for the run and restores yours; builds in `dist\` and in its folder; writes `test-results\excel-check-SHORTSHA-TIME` and a zip; copies the summary to the clipboard; closes the Excel instances it started |
| Python tools under `tools/` | Tracked files and given paths. `check_issues.py` and `check_external_links.py` also read GitHub or the linked sites over the network | The output files they are given, except: `update_expected.py` rewrites `src/modules/M_Setup.bas` and rebuilds the template; `build_template.py` rewrites `src/workbook/FRTB_Template.xlsx` and `build_acceptance_data.py` rewrites `tests/modules/TEST_AcceptanceData.bas` (neither with `--check`); `build_market_diagrams.py` writes `docs/assets/market-risk/` by default; `traffic_history.py` rewrites the CSV files and `badges/` in the data folder it is given |
| `tools/traffic_publish.py` (traffic workflow, on main only) | The traffic history folder; `GITHUB_TOKEN` | Creates one tree and one commit through the GitHub API and moves the remote `traffic-history` branch to it, never forced |
| GitHub workflows | The repository | Only the label, PR-milestone and traffic workflows write: labels, PR milestones, the `traffic-history` branch and traffic alert issues. Each write permission is limited by `tools/check_repository.py` |

**Host settings.** The build, sync, export, recorder and unattended-check scripts need Excel's
*Trust access to the VBA project object model*. Turn it on only for those
runs if your policy allows, and record it. The workbook's buttons need macros
enabled. The export and the recorder's compile open workbooks with macros
disabled.

**Data.** The sample portfolio and every test fixture are synthetic. A
workbook can hold data that is not visible: hidden rows and sheets, names,
document properties and VBA. The committed template carries no VBA project
and no personal properties. Before you share a workbook or an evidence
bundle, check it for such content. A recorder bundle records the user and
computer name, so keep it private; the public evidence record leaves them
out. If a credential or token leaks, revoke or rotate it at its issuing
service first, then report it privately as below.

## 🔒 Report privately

Use [GitHub private vulnerability reporting](https://github.com/danielep71/K-FRTB-ASA/security/advisories/new).
This channel was verified enabled in repository security settings on
October 18, 2026 (Europe/Rome). It permits private discussion with me;
do not publish the report as an advisory yourself.

Include the affected commit/branch, relevant component, prerequisites,
minimal synthetic reproduction, expected/observed behavior, potential impact,
and any proposed mitigation. Provide only the sensitive detail necessary
to reproduce the issue; never submit live credentials, client data, or
unrelated private information.

Do not post undisclosed vulnerabilities in public issues, PRs, comments, or
attachments. For ordinary non-security defects, use the bug issue form.

If the private-report link is unavailable, do not fall back to a public issue.
For an exposed credential or an account/platform security incident, contact
[GitHub Support](https://support.github.com/) and revoke/rotate the affected
credential through its issuing service. Support handles platform/account
incidents; it is not a substitute project vulnerability inbox.

## 🔎 Handling and disclosure

I assess the report, ask for missing reproduction detail, and coordinate
a fix and disclosure when warranted. Keep discussion private
until a coordinated disclosure decision. Reports may require further evidence
or be found outside the maintained scope; the outcome should be explained.

## 🤝 Good-faith research

I will not pursue action against research that stays within these limits:

- it uses only your own data, accounts, forks and local copies;
- it does not degrade the repository, its workflows or GitHub for others,
  access data that is not yours, or keep such data;
- it is reported privately as above, and kept private until a coordinated
  disclosure decision.

This covers my own project and my own claims only. It cannot authorize
anything on behalf of GitHub or any third party, whose own terms apply.

Repository scanning, passing workflows, or a lack of reported vulnerabilities
does not establish that the project is secure or that future capital
calculations are correct. Static checks do not run Excel or a calculation
engine. Security reports and calculation-validation evidence serve different
purposes.
