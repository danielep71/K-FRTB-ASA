# 🤝 Contributing to FRTB-SA

[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)
[![Stage: prototype](https://img.shields.io/badge/stage-prototype-1D76DB)](docs/ROADMAP.md)

Participation follows the [Code of Conduct](CODE_OF_CONDUCT.md).
Potential vulnerabilities follow [SECURITY.md](SECURITY.md). The editorial
rules and the rules for automated assistants are in [AGENTS.md](AGENTS.md).

Repository setup (milestone 1, v0.1.0 — Repo setup) is closed, and its branch
`release/v0.0.1` was deleted (#62); its history is on main. The Excel/VBA SA
prototype and repository maintenance use `release/v0.2.0`, milestone 2,
v0.2.0 — SA prototype (#117); `release/v0.1.0` is superseded.
Read [scope and decisions](docs/SCOPE.md)
before functional changes. The prototype is on main and release/v0.2.0 with a
committed macro-free template; Excel host evidence remains pending in #28.
Versions and tags follow [Releasing](RELEASING.md).

## Start with an issue

Describe the objective, acceptance criteria, dependencies, affected contracts,
and evidence needed to finish. For a defect, include the source commit,
environment, steps to reproduce, expected behavior, and observed behavior.
For documentation, identify the authoritative document and proposed correction.
For a future calculation change, identify its regulatory source, applicability,
and independently specified expected results.

[Governance](docs/GOVERNANCE.md) owns branch, metadata, review, and completion
rules. [Labels and triage](docs/LABELS.md) owns the label catalogue and priority
definitions. Use the bug, feature, or documentation form in the issue chooser,
which GitHub serves from main. Use the
[PR template](.github/PULL_REQUEST_TEMPLATE.md) for scope, evidence, and the
single merge checklist. API-created issues and PRs must supply the same
information and metadata explicitly: issues have milestones; PRs do not.

## Make a focused change

1. Fetch the repository and branch from `origin/release/v0.2.0` for current
   prototype or maintenance work (milestone 2), or from
   `origin/release/v0.3.0` for milestone 3 work (#132). Later milestones need
   an explicitly approved integration branch. Confirm the PR base explicitly.
2. Reproduce the problem where applicable and make a focused change.
3. Follow the root editing, ignore, and attribute policies. Do not include
   credentials, client data, local caches, or generated output.
4. Update affected documentation and [CHANGELOG.md](CHANGELOG.md). Write
   repository content in US English.
5. Run checks appropriate to the changed files. Record the exact commands,
   source commit, results, and anything unavailable.
6. Open the PR against that milestone's integration branch, link the issue, explain scope and
   compatibility impact, and include validation and unresolved findings.
   Apply the issue's assignee and one priority label to the PR. Keep the
   milestone on the linked issue only; leave the PR milestone unset.
7. Follow the review and completion procedure in governance.

For example:

```sh
git fetch origin
git switch -c docs/117-release-branch origin/release/v0.2.0
```

Use an imperative commit subject describing the change. Add a body when useful
to explain its reason and validation; distinguish local file checks from
runtime execution. Reference the issue without including private data or
unsupported claims.

## Validation available today

The label maintenance tools use Node.js with built-in fetch support, without
npm dependencies. For label configuration or tooling changes, run from the
repository root:

```sh
node .github/scripts/labels-sync.mjs --self-test
node .github/scripts/labels-sync.mjs --mode validate
node .github/scripts/labels-drift.mjs --self-test
```

For staged changes, check whitespace with `git diff --cached --check`.
For documentation, verify local paths, headings, and commands against the
candidate tree. For ignore or attribute changes, exercise representative
paths with `git check-ignore --no-index` and `git check-attr`.

Run `python tools/check_repository.py` after installing the pinned
`.github/requirements.txt`, `.github/requirements-quality.txt` and
`tests/requirements.txt`; [Repository checks](docs/CI.md) documents its
tracked-file inventory, its quality gates (Ruff, mypy and actionlint),
prerequisites and limits. A gate whose tool is missing is skipped unless
`REQUIRE_QUALITY_TOOLS=1`, as in CI. Run `python -m ruff check --fix tools tests` to fix
import order. Inspect applicable
hosted checks and reviews before merging; no emitted job is not a passing job.

The prototype has a separate suite: `python3 tools/vba_lint.py` and
`REQUIRE_LIBREOFFICE=1 python3 -m pytest tests`, using the pinned dependencies
from both `.github/requirements.txt` and `tests/requirements.txt`. See
[validation](docs/validation.md) for LibreOffice prerequisites and the
separate outstanding Excel evidence.

## Compatibility and evidence

State any change to documented inputs, outputs, defaults, errors, file formats,
supported environments, or maintenance behavior. Update callers and relevant
documentation together. Synthetic fixtures must identify their provenance.
Expected calculation results must be independently derived, not generated by
the implementation being tested.

Static checks, runtime tests, independent numerical comparisons, and
methodology review establish different things. The current maintenance
commands do not execute Excel or validate FRTB calculations. A future engine
must define its runtime and evidence contract before claiming support.

Do not disclose suspected vulnerabilities in public issues or PRs. Use
[private vulnerability reporting](https://github.com/danielep71/K-FRTB-SA/security/advisories/new),
enabled in the [settings baseline](docs/REPOSITORY_SETTINGS.md).
See the dedicated security and conduct policies for reporting and moderation.

## Rights and licenses

Submit only material you have the right to contribute. This project and
contributions to it use the [MIT License](LICENSE). Preserve applicable
copyright and permission notices when redistributing project material.
