# Issue labels and triage

The versioned catalogue in [labels.json](../.github/labels.json) is the source
of truth for the names, descriptions and colors of the 20 labels. It uses
schema 2: a single `labels` list, with no profile or domain overlays.

## Metadata rules

Every development issue, including issues opened through the API and closed
issues, must have:

- Assignee `danielep71`.
- Exactly one priority: `P1`, `P2` or `P3`.
- The milestone that owns its work; current prototype and repository maintenance
  belong to **v0.2.0 — SA prototype** (milestone 2, approved in #51).
  Historical setup issues retain the closed **v0.1.0 — Repo setup** milestone.
  An open issue's milestone must be open.
- A target date, `YYYY-MM-DD`, no later than the milestone's due date: the
  **Target end date** of the issue forms, or a `**Target completion date:**`
  planning line in the body, which takes precedence (#221).
- An objective, acceptance checklist, dependencies and required evidence.

PRs retain the assignee and exactly one priority, and link to their owning
issue. **PRs must have no milestone**, including closed, merged and bot PRs.
Milestone ownership is recorded on the issue only (#104). The metadata
workflow removes PR assignments automatically; issue triage remains manual,
with the read-only audit described below.

**P1 is reserved for a confirmed calculation bug: wrong results for valid
inputs.** It is used only together with `bug`. An input that should have been
rejected but was accepted, such as a blank parameter, is a P2 defect, not P1.
P2 covers material defects in correctness, robustness or maintainability. P3
covers setup, documentation, testing and other non-blocking quality work. The
catalogue description of P1 states the same rule.

| Priority | Use |
| --- | --- |
| `P1` | Confirmed calculation bug: wrong results for valid inputs; only with `bug` |
| `P2` | Material correctness, robustness or maintainability defect |
| `P3` | Setup, documentation, testing and other non-blocking quality work |

| Label | Use |
| --- | --- |
| `bug` | Confirmed defect in behavior or a documented contract |
| `enhancement` | New capability or meaningful improvement |
| `documentation` | Documentation for users or for my repository maintenance |
| `tests` | Test coverage, fixtures or manual validation |
| `refactor` | Internal restructuring that preserves behavior |
| `behavior-change` | Intentional externally observable behavior change |
| `ci` | Automated checks, workflows or certification |
| `repository` | Structure, governance, metadata or tooling |
| `release` | Release preparation, evidence or packaging |
| `security` | Disclosed security work or supply-chain hardening |
| `blocked` | Work cannot proceed until a named dependency is resolved |
| `question` | I need further information or need to make a decision |
| `help wanted` | External contributions or investigation are welcome |
| `good first issue` | A bounded task suitable for a first contribution |
| `duplicate` | Another issue or PR already covers the work |
| `invalid` | The report is unsupported, incomplete or not actionable |
| `wontfix` | I have decided not to proceed |

Type and status labels may be combined; priority labels may not. Setup issues
normally carry `repository` and `P3`, plus applicable type labels. Every
open issue has a section whose heading contains "Dependencies", normally
**Dependencies and blocking status**. It names each prerequisite issue
(`#123` or its URL), or an `External condition:` line for a blocker outside the
tracker. `blocked` is set exactly while that section names an open
prerequisite or an external condition; a named pull request is a reference,
never a prerequisite. Before closing, link the merged PR,
applicable green checks and any required runtime evidence; only mark acceptance
checkboxes complete when that evidence exists.

**Issue triage** ([workflow](../.github/workflows/issue-triage.yml),
[check](../tools/check_issues.py)) audits every open issue on issue events, on
milestone events (a new due date or a closed milestone changes the rules of
its issues) and on manual dispatch: assignee `danielep71`, exactly one
priority, P1 only with `bug`, an open milestone, a real target date no later
than its due date, labels from the catalogue, and the dependency rule above. It is read-only: a failed run is the triage signal, and I resolve each
finding by hand. `python tools/check_issues.py --check-config`, run by
`tools/check_repository.py`, checks offline that the forms, this guide and
Dependabot agree with the catalogue, and that the three work forms require
the target end date. For an offline audit of exported issues:
`python tools/check_issues.py --input issues.json --repo-name danielep71/K-FRTB-SA`.
The label workflows manage the catalogue only: they do not assign issues,
choose priorities or milestones, or close issues. The bug form defaults to
P2, and feature/documentation forms to P3; all assign danielep71. Milestones
and metadata confirmation remain manual, and API-created issues bypass form
defaults. [Governance](GOVERNANCE.md) explains default-branch activation and
the contribution workflow; live activation was verified through #21/#6.

## Reusable triage searches

Run these repository-wide, without `is:open` or a milestone filter, so closed
issues and orphaned issues remain visible. Every search should return no issues.
The PR milestone search should also return zero results.

| Check | GitHub search query |
| --- | --- |
| Missing assignee | `repo:danielep71/K-FRTB-SA is:issue no:assignee` |
| Wrong assignee | `repo:danielep71/K-FRTB-SA is:issue -assignee:danielep71` |
| Missing milestone | `repo:danielep71/K-FRTB-SA is:issue no:milestone` |
| PR incorrectly assigned a milestone | `repo:danielep71/K-FRTB-SA is:pr has:milestone` |
| Missing priority | `repo:danielep71/K-FRTB-SA is:issue -label:P1 -label:P2 -label:P3` |
| Conflicting priorities | `repo:danielep71/K-FRTB-SA is:issue label:P1 label:P2` |
| Conflicting priorities | `repo:danielep71/K-FRTB-SA is:issue label:P1 label:P3` |
| Conflicting priorities | `repo:danielep71/K-FRTB-SA is:issue label:P2 label:P3` |

Review P1 issues individually to confirm a calculation bug is established;
a search cannot verify the substance of that decision.

## Validation and reconciliation

Requires Node.js with built-in `fetch` support; no npm dependencies are needed.
Run from the repository root:

```sh
node .github/scripts/labels-sync.mjs --self-test
node .github/scripts/labels-sync.mjs --mode validate
node .github/scripts/labels-drift.mjs --self-test
```

**Sync issue labels** validates PRs into `main` and `release/**` without write
permissions. A relevant push to `main`, or a manual dispatch on the default
branch, runs trusted reconciliation with `issues: write`; a dispatch on any
other branch skips it, so no branch other than the default can apply its
catalogue. It records the live catalogue before
changes, a plan, the resulting catalogue and an independent drift report.
Reconciliation verifies the exact final match; it is not best effort.

**Detect issue-label drift** runs offline fixtures on relevant PRs. Scheduled
runs (05:17 UTC daily) and manual runs read the live catalogue without mutation
and retain JSON/Markdown evidence for 30 days. Drift fails the run and requires
review; detection does not silently repair it. Scheduled workflows use the
default branch; they are active on `main` (drift run 39105486786 passed).

The manifest uses `prune: true`: labels outside the selected catalogue are
deleted, which also removes their issue/PR associations. Before any change to
the catalogue or manual reconciliation, review the plan and inspect affected
issue/PR associations; migrate those associations before removing a used label.
The reconciler refuses all mutations if any planned deletion still has an
open or closed issue/PR association, or if that lookup fails. This guard uses
the same validated plan as the mutations; updates and creations never trigger
it, and `prune: false` performs no deletion association checks. The workflow
retains its read-only preview; the script rechecks current live labels and
associations before applying its actual plan. When the final check fails, the
script writes one summary of the failed run. The offline self-test covers
updates, disabled pruning, used and unused deletion candidates, and failed or
malformed association lookups.
For comma-containing label names, the script scans actual issue/PR labels
instead of GitHub's comma-separated filter. The scan includes closed items,
matches complete names case-insensitively, and fails closed on malformed data
or after 100 pages of 100 items; it never assumes a truncated scan is empty.

Keep deliberate catalogue changes in a reviewed PR. Do not run unreviewed
workflow code with a write token.

For a read-only preview against a saved JSON label array:

```sh
node .github/scripts/labels-sync.mjs --mode plan --live labels-before.json
```

With `GITHUB_TOKEN` set securely in the environment and an explicit reviewed
checkout, I can run live verification or reconciliation:

```sh
node .github/scripts/labels-drift.mjs --repository danielep71/K-FRTB-SA
node .github/scripts/labels-sync.mjs --mode reconcile --dry-run --repository danielep71/K-FRTB-SA
node .github/scripts/labels-sync.mjs --mode reconcile --repository danielep71/K-FRTB-SA
```

The scripts use bounded pagination and retries. Workflow actions are pinned to
full commit SHAs, checkout credentials are not persisted, and PR validation
never receives a write token.
