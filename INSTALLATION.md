# Repository setup and local verification

This guide covers repository maintenance and the separate ASA prototype path.
The universal maintenance check requires Git, Node.js with built-in fetch, Python
3.12 or newer, the pinned packages of the three requirements files
(`.github/requirements.txt`, `.github/requirements-quality.txt` and
`tests/requirements.txt`) and [actionlint](docs/CI.md#quality-gates-219). No npm
packages, Excel or PowerShell are required for maintenance. Local verification used Linux,
Git 2.51.1, Node.js 24.19.0 and Python 3.12.14; hosted CI uses Ubuntu 24.04,
with Python 3.13.18 and Node.js 24.22.2 for the integrity gate.

## Check out the active work

```sh
git clone --branch release/v0.2.0 https://github.com/danielep71/K-FRTB-ASA.git
cd K-FRTB-ASA
git branch --show-current
git rev-parse HEAD
git --version
node --version
```

Current prototype and maintenance work uses `release/v0.2.0`, milestone 2,
v0.2.0 — ASA prototype (#117). It supersedes `release/v0.1.0`; the setup
branch `release/v0.0.1` was deleted after the setup closeout (#62). Record the source commit when reporting results.
Main remains the stable default. Create focused task
branches from the integration branch following [Governance](docs/GOVERNANCE.md).

Use a Git clone for development and checks. The inherited archive attributes
exclude repository-control files and `.github/` from Git-produced source
archives, so such an archive is not a complete maintenance checkout.

## Verify the setup

From the repository root:

```sh
python3 -m venv .venv-ci
.venv-ci/bin/python -m pip install -r .github/requirements.txt -r .github/requirements-quality.txt -r tests/requirements.txt
.venv-ci/bin/python tools/check_repository.py
git diff --check
git diff --cached --check
git status --short
```

Use `.venv-ci\Scripts\python` on Windows. The universal command must exit 0;
it includes the quality gates (Ruff, mypy and actionlint), the existing label
commands, maintenance fixtures, JSON/YAML, tracked whitespace and local
Markdown target and anchor checks. A gate whose tool is missing prints `SKIP`;
set `REQUIRE_QUALITY_TOOLS=1` to fail instead, as CI does. Stage new files first
so they appear in its tracked-file inventory. No GitHub credentials are used.
The additional diff checks inspect local changes; the universal command also
inspects unchanged tracked files. A fresh checkout should be clean.

Read the purpose of `src/`, `tests/`, `docs/`, and `tools/` in
[Repository structure](docs/REPOSITORY_STRUCTURE.md). Track new source,
fixtures, manifests, and reviewed documentation deliberately. Keep caches and
generated output local. For policy changes, probe representative paths with
`git check-ignore --no-index` and `git check-attr`.

See [Repository checks](docs/CI.md) for hosted branch coverage, Dependabot,
failure behavior and the #2 protection handoff. These checks do not execute
an engine or Excel, test a workbook, or establish verified capital results.

## ASA prototype (v0.2.0)

The [v0.2.0 — ASA prototype](https://github.com/danielep71/K-FRTB-ASA/milestone/2)
work integrates on `release/v0.2.0` (issue #20). Its approved Excel/VBA
evaluation scope does not resolve the long-term platform decision or the
remaining methodology questions in [Scope](docs/SCOPE.md).

- Checks on Linux: Python 3.12 or newer, LibreOffice Calc with the Python UNO bridge
  (`python3-uno`) and the pinned packages in `tests/requirements.txt` plus
  `.github/requirements.txt`; run
  `python3 tools/vba_lint.py` and `REQUIRE_LIBREOFFICE=1 python3 -m pytest tests`
  as described in [Validation](docs/validation.md).
- Workbook on Windows: desktop Excel and PowerShell; run
  `tools\Build-Workbook.ps1 -RunTests` (or double-click
  `tools\Build-Workbook.cmd`) as described in the
  [workbook guide](docs/workbook-guide.md#building-the-workbook-windows-excel).
  To update an existing workbook from the checkout or a downloads folder, use
  `tools\Sync-FRTB-VBA.cmd` or `tools\Sync-FRTB-VBA-Downloads.cmd`; see
  [updating an existing workbook](docs/workbook-guide.md#updating-an-existing-workbook).

The macro-free template `src/workbook/FRTB_Template.xlsx` is generated from
`src/modules` (`python3 tools/build_template.py`); see the
[workbook guide](docs/workbook-guide.md#source-of-truth).
