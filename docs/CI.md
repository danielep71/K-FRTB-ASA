# 🧪 Repository checks and dependency updates

[![Scope: maintenance](https://img.shields.io/badge/checks-repository_maintenance-1D76DB)](../tools/check_repository.py)

## Run the same check locally

Use a Git clone, Python 3.12 or newer, Node.js with built-in fetch, and Git.
Install the three hashed requirements files in a local environment:

```sh
python3 -m venv .venv-ci
.venv-ci/bin/python -m pip install -r .github/requirements.txt -r .github/requirements-quality.txt -r tests/requirements.txt
.venv-ci/bin/python tools/check_repository.py
```

The quality gates need the second and third files and
[actionlint](#quality-gates-219). Without a tool, a local run skips its gate
and prints `SKIP`; set `REQUIRE_QUALITY_TOOLS=1` to fail instead, as CI does.

On Windows use `.venv-ci\Scripts\python` instead of `.venv-ci/bin/python`.
Run from the repository root. Stage new source and documentation before running
the command: it inspects paths returned by `git ls-files`, using their current
working-tree contents. It checks committed files even when the diff is empty.
It does not stage, commit, change labels, read secrets, or contact GitHub.
Installing the package requires access to the package registry; validation
after installation and all maintenance fixtures are offline.

## What the gate checks

- All tracked JSON and YAML parse successfully, with duplicate keys rejected.
  GitHub's YAML `on` key remains text. Workflow files have basic structural
  checks and external step actions use full commit SHAs. This is not complete
  GitHub Actions or issue-form schema validation.
- Every workflow declares top-level `permissions`, every job sets
  `timeout-minutes`, and every `actions/checkout` step sets
  `persist-credentials: false`. A job may hold a write permission only if it is
  listed in `WRITE_ALLOWED` of `tools/check_repository.py`, for that scope:
  `labels-sync.yml` `reconcile` (`issues`), `traffic-export.yml` `export`
  (`contents`) and `alert` (`issues`), and `pr-milestones.yml` `reconcile`
  (`pull-requests`).
  Workflow-level permissions and `write-all` count too. The label and PR
  milestone reconciliations must keep their default-branch conditions.
- Tracked text has a final newline and no accidental trailing whitespace or
  extra EOF blank lines. Intentional Markdown hard breaks are allowed. VBA
  exports may use CRLF and Windows-1252. Binary assets are excluded according
  to Git's resolved `text` attribute from `.gitattributes`, including its
  macros and path overrides; the checker has no separate extension list.
- Common inline Markdown links/images and reference definitions outside code
  fences point to tracked repository files or directories. A `#fragment`, in
  the same file or in another tracked Markdown file, must name one of its
  anchors: a GitHub heading slug (with `-1`, `-2` for repeats) or an HTML `id`
  or `name`; a file with a duplicated explicit anchor fails (#218). Remote URLs
  are not fetched here; this is not a complete Markdown parser.
- The universal workflow keeps its unconditional stable check and branch
  coverage. Dependabot covers the inspected ecosystems and intended branches.
- The integrity workflow sets exact Python and Node.js versions, installs its
  Python packages by hash from wheels only, writes its report and always
  uploads it, and cancels only superseded PR runs. The three requirements
  files pin every package, including dependencies, with at least one hash. The PR
  milestone workflow keeps one concurrency group per PR. Every workflow step
  that runs `pip install` must use `--require-hashes --only-binary=:all:`.
- The existing label self-tests, manifest validation and read-only drift
  fixtures run once through their authoritative scripts. The path-filtered
  label workflows remain responsible for their own live evidence and trusted
  reconciliation; this runner does not duplicate their implementations.
- Offline maintenance tests under `tests/tooling` cover malformed/duplicate
  configuration (including non-standard JSON numeric constants), binary policy,
  broken links, whitespace, unpinned actions, narrowed trigger
  coverage, incorrect update targets, Windows-only test dependencies missing
  from `tests/requirements.txt`, and the traffic collector's safeguards.

CodeQL runs as GitHub's **default setup** for Actions and Python: the
default query suite, weekly and on pull requests into main (#214 ran it; a
release-branch PR such as #224 gets no CodeQL run). It is configured in the
repository's security settings, not in a workflow file. Its configuration therefore cannot
be rebuilt from a commit; [settings evidence](REPOSITORY_SETTINGS_EVIDENCE.md)
records its last readback. A versioned CodeQL workflow would duplicate the
scans, so none is added (#223).

[External link check](../.github/workflows/link-check.yml) runs weekly on
Mondays at 04:41 UTC and on manual dispatch, read-only and outside the required
checks. After installing the hashed `.github/requirements.txt` (the checker
reuses the gate's Markdown helpers, which import PyYAML),
`tools/check_external_links.py` requests each `http(s)` link of the
tracked Markdown files once, without its fragment, following redirects, with a
25-second timeout, up to two retries and three requests at a time. A 2xx answer
is verified; 403, 429, 5xx and network errors are unverifiable and reported as
warnings, never as a pass; any other answer is broken and fails the run. Links
to this repository on github.com are left to the local gate (#218).

A failed prerequisite, validation error or failed child command exits nonzero.
The fixture suite includes clean input as well as deliberate failures. It
never deletes a live label or exports real traffic.

## Hosted check and protection handoff

Workflow: [Repository integrity](../.github/workflows/integrity.yml).
The job ID is `integrity`; the exact check name is **Repository integrity**.
The producer is **GitHub Actions** (app slug `github-actions`, app ID `15368`).
There is no matrix, job condition, path filter or changed-file shortcut.
PRs targeting `main` or `release/**`, and pushes to those branches, run the
same command on Ubuntu 24.04. Manual dispatch is also available after
default-branch activation. Checkout is pinned to the existing reviewed SHA,
credentials are not persisted, permissions are `contents: read` and the
timeout is 10 minutes.

### Reproducible runs (#220)

- **Runtimes:** `actions/setup-python` and `actions/setup-node`, pinned by
  SHA, install Python 3.13.18 and Node.js 24.22.2, so a runner image update
  cannot change them. Dependabot updates the actions; I change the versions
  by hand.
- **Dependencies:** the three requirements files, `.github/requirements.txt`,
  `.github/requirements-quality.txt` and `tests/requirements.txt`, are
  compiled from their `.in` files with `pip-compile --allow-unsafe
  --generate-hashes --no-emit-index-url --output-file=<file>.txt
  --strip-extras <file>.in`, the command each records in its header. They pin
  every package, dependencies included, by version and hash. CI installs them
  with `--require-hashes --only-binary=:all:`. A local `pip install -r`
  checks the hashes too, because pip turns hash checking on when a file
  has hashes. After changing an `.in` file, compile it again and commit both.
  pip-compile evaluates environment markers where it runs, so a compile on
  Linux or macOS drops colorama, the Windows-only dependency of pytest that
  `tests/requirements.in` pins with its marker; I compile that file on
  Windows or restore the entry, and
  `tests/tooling/test_requirements_platforms.py` fails without it (#246).
- **Report:** `tools/check_repository.py --report PATH` writes a JSON report
  of every check it ran: the commit, run id and attempt, the runtimes, and
  each check's result and duration. It writes the report whatever the result.
  On a PR, `commit` is the merge commit GitHub tests and `head_commit` the PR
  head; on a push they are the same. The job uploads the report as the
  artifact `integrity-report-<head sha>-<run id>-<attempt>`, kept 30 days,
  even when a check failed; a missing report fails the job.
- **Concurrency:** a PR run has one group per PR, and a newer head cancels
  the older run. A push or dispatch run has a group of its own and is never
  cancelled, so each pushed commit keeps its result. The PR milestone
  workflow has one group per PR for PR events, and one for its scheduled,
  manual and push scans, so one PR's pending event no longer replaces
  another's.
- **Prototype checks:** this job keeps the OS Python, because the
  LibreOffice UNO bridge is built for it, so its Python version follows the
  runner image. It installs the hashed requirements with `--ignore-installed`
  as well, so every package comes from its hashed wheel.
- **External link check:** it installs the parser the same way, on the
  runner's Python; it is not a required check.

### Quality gates (#219)

The required check runs these before the label commands and the tooling
tests, with the configuration of [`pyproject.toml`](../pyproject.toml):

| Gate | Scope | Rules |
| --- | --- | --- |
| Ruff 0.17.3 | `tools/`, `tests/`, `.github/` | `E4`, `E7`, `E9`, `F`, `I001` (import order) and `C90` with a maximum complexity of 15 |
| mypy 2.4.1 | `tools/` | `strict`, `warn_unused_configs` and `warn_unreachable` |
| mypy 2.4.1 | `tests/` | the same, with only the annotation requirement relaxed: function bodies are still type-checked |
| actionlint 1.7.12 | every workflow | its own syntax, expression and type checks |

- **Exceptions:** a function above the complexity limit carries a
  `# noqa: C901` with its reason. Six do: the independent FRTB reference's
  `delta_risk_factors` and `sbm_component`, which follow the MAR21 steps in
  order, the engine reconciliation, the evidence validator, the LibreOffice
  runner and the VBA lint parser. No stubs exist for the LibreOffice UNO
  bridge or CairoSVG, so mypy does not follow those imports.
- **Versions:** Ruff, mypy and the PyYAML and openpyxl stubs are pinned with
  hashes in `.github/requirements-quality.txt`, compiled from
  `.github/requirements-quality.in`. The job also installs
  `tests/requirements.txt`, so mypy checks the tests against NumPy, openpyxl
  and pytest. It downloads the actionlint release archive and checks its
  SHA-256 before use.
- **actionlint integrations:** actionlint runs without ShellCheck and
  pyflakes. Those tools would come unpinned from the runner image, so the
  shell and Python inside `run:` steps are not linted.
- **Missing tools:** the job sets `REQUIRE_QUALITY_TOOLS=1`, so a missing
  tool fails the check. `tests/tooling/test_quality_gates.py` proves that each
  gate fails: an unused import, a complexity breach, a type error, an
  unreachable branch and an invalid workflow expression. It also shows that a
  missing tool fails only when tools are required. The gate refuses an
  integrity workflow that drops `REQUIRE_QUALITY_TOOLS` or the actionlint
  checksum.

The check name and app identity are required by the main and release
rulesets, and both are strict: a PR branch must be up to date with its base
before merging. Issue #2 records the readback and a representative blocked
PR (October 18, 2026); the release ruleset has been strict since the
readback of October 31, 2026
([settings evidence](REPOSITORY_SETTINGS_EVIDENCE.md)). A task PR proves
the hosted command; release-to-main integration and synchronization into the
prototype branch (#31) install the workflow on the respective branch lines.
The prototype workflow installs both requirements manifests in its own test
environment before pytest discovers the shared maintenance tests. Its
LibreOffice installation step tries the runner's package mirror first, with
short apt timeouts and limits of 2 minutes for the index and 5 for the
packages. When that mirror does not answer, it installs again from
`mirrors.edge.kernel.org/ubuntu` over HTTPS and prints a warning (#263): the
runner could not resolve any ubuntu.com host, so neither another ubuntu.com
mirror nor HTTPS helped. Apt still checks the Ubuntu archive signatures, so
the other mirror changes where packages come from, not what is trusted. The
step stops after 12 minutes, within the job's 25-minute limit, so a hung mirror still fails
early (#124). The
universal job remains a separate, runtime-independent check.
Do not require a check on a line before installing it there.

## 🔄 Dependabot policy

The configuration must reach the default branch before GitHub schedules
version updates. The temporary maintenance successor was approved in
[#51](https://github.com/danielep71/K-FRTB-SA/issues/51#issuecomment-6254417218)
on October 18, 2026, and the entries moved to `release/v0.2.0` with the
integration branch on October 27 (#118); the configuration was activated
on main through #120.
PRs have no milestone under the issue-only policy (#104).
The setup branch `release/v0.0.1` receives no version-update
PRs; it was deleted after the closeout (#62):

| Ecosystem and manifest location | Version-update PR target |
| --- | --- |
| GitHub Actions, `/` | `release/v0.2.0` |
| pip, `/.github` (`requirements.txt`: PyYAML configuration parser; `requirements-quality.txt`: Ruff, mypy and stubs) | `release/v0.2.0` |
| pip, `/tests/requirements.txt` (prototype NumPy, openpyxl, pytest) | `release/v0.2.0` |

The two previous Actions entries are consolidated into one. Both pip
manifest locations remain covered; no dependency is copied into the historical
setup branch. The integrity gate rejects duplicate entries, retired or wrong
targets, lost ecosystem/directory coverage and any `milestone` setting.
There is no npm package manifest: the label scripts use Node built-ins. No npm updater or
other absent ecosystem is configured. Each entry runs weekly, limits open
PRs to five, assigns `danielep71`, and sets exactly P3 plus type labels.
I still verify all metadata during triage and put any owning issue in the
appropriate milestone.

Non-default `target-branch` settings govern version updates, not security
update PRs. Dependabot alerts and automatic security updates are on
([settings](REPOSITORY_SETTINGS.md#security)). A security update PR
opens against main; I retarget it to the active release branch before
review. See the [official options reference](https://docs.github.com/en/code-security/reference/supply-chain-security/dependabot-options-reference).

## Issue-only milestones

[Keep milestones issue-only](../.github/workflows/pr-milestones.yml) removes
milestone assignments from PRs only (#104). It handles `opened`, `reopened`,
`edited` and `milestoned` PR activity, plus a daily full reconciliation at
03:17 UTC, manual dispatch and reviewed workflow changes pushed to main.
PR events and the schedule always run the reconciliation; a push or manual
dispatch runs it only from the default branch, so no other branch's copy of
the workflow gets the write token. `tools/check_repository.py` requires this
condition. The full scan includes open, closed and merged
PRs. It uses unfiltered creation-order pagination so removing assignments does
not shift the collection. A 100-page bound, malformed response, API error or
failed removal fails the run rather than claiming success.

The workflow runs trusted inline Python without checkout, third-party actions,
PR code, PR text interpolation or custom secrets. Its token has only
`pull-requests: write` for PR metadata updates and listing; all other
permissions are absent. Each mutation first reads the live issues endpoint and
refuses to modify an object without a `pull_request` marker. An unset milestone
is a no-op. Offline fixtures exercise pagination, actual issue refusal,
idempotence and error propagation against the workflow's inline implementation.

Default-branch activation is required. My October 27 instruction
to apply this rule to past and future PRs is recorded in #104; focused policy
PRs activated it on main and synchronized release/v0.1.0. This activation does
not integrate unrelated release work. The activation push performs the historical
cleanup; daily runs also cover changes by bots whose events may not trigger
another workflow. Event delivery and scheduled runs can be delayed or blocked
by GitHub Actions policies or outages. If a run fails, inspect it and rerun the
trusted workflow; do not claim synchronous assignment prevention. The query
`repo:danielep71/K-FRTB-SA is:pr has:milestone` should return zero.

## Issue triage

[Issue triage](../.github/workflows/issue-triage.yml) runs
[`tools/check_issues.py`](../tools/check_issues.py) on issue events, on
milestone events and on manual dispatch. Issue and milestone events run the
workflow from the default branch only, so it is active after the change
reaches main. Its token has `contents`, `issues` and `pull-requests` read
access only: it never changes an issue, and a failed run reports the findings
in the job summary and a 30-day artifact. The rules are in
[labels and triage](LABELS.md).

## Evidence boundaries

This gate checks repository maintenance. It does not execute Excel, run the
FRTB engine, validate formulas, certify regulation, or substitute for runtime
and independent numerical evidence. The prototype's separate checks already
cover VBA lint, Python reference tests and LibreOffice reconciliation on its
branch. One desktop Excel run is recorded under #28
([Excel run](evidence/2026-10-31-excel-run-323260b.md)), which keeps the
remaining manual Excel checks; the methodology readings decided in #29 are
reviewed for the certified scope in v0.3.0; prototype synchronization was
completed in #31. Future engine gates
must retain independently derived expectations and name the actual host and
commit, following [Scope](SCOPE.md) and the [structure policy](REPOSITORY_STRUCTURE.md).
