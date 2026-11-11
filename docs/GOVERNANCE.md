# Repository governance

I own and maintain this repository (**@danielep71**).

Explicit approvals and requests for integration, publication or settings
changes in these documents must come from me.

This document owns branch, issue, review, and milestone procedures.
[CONTRIBUTING.md](../CONTRIBUTING.md) provides the contributor walkthrough;
[labels and triage](LABELS.md) owns label definitions,
[AGENTS.md](../AGENTS.md) owns the editorial and assistant rules, and
[repository settings](REPOSITORY_SETTINGS.md) records effective GitHub controls.

## Branches

The active integration branch is **`release/v0.2.0`** for the SA prototype
and repository maintenance, milestone 2. I approved it on
October 27, 2026, after the milestone renumbering, and it was created from
`0eb819b92e6f627099ceb1400c15ebb4404e385b`, then the head of both main and
`release/v0.1.0` (#117).

**`release/v0.3.0`** is the integration branch for milestone 3, v0.3.0 —
Regulatory baseline and certification scope. Its work includes the review,
for the certified scope, of the methodology readings I decided under #29 on
October 27, 2026. I approved the branch on October 28, 2026, and it was
created from `69e3df35781fc2e8a13bb523059d596d9da08150`, then the head of
both main and `release/v0.2.0` (#132). It runs alongside `release/v0.2.0`,
which keeps the prototype closeout and repository maintenance; bring
`release/v0.2.0` or main into it with a merge commit when it needs their
changes.

**`release/v0.1.0`** is superseded: no PR targets it. It was approved for the
prototype on October 17, 2026 and created from main at
`d150d3822d5d50f69aabfca153bf6d302e80f2e4` (the integration of #18);
maintenance joined it on October 18
([decision in #51](https://github.com/danielep71/K-FRTB-ASA/issues/51#issuecomment-6254417218)).
Its whole history is in `release/v0.2.0` and on main. It is deleted under the
[release-branch lifecycle](#release-branch-lifecycle) once milestone 2 is
closed. Dependabot version updates target `release/v0.2.0` since #118,
activated on main through #120.

The setup milestone, v0.1.0 — Repo setup (first named v0.0.1), is closed. The integration branch
**`release/v0.0.1`** was deleted on October 18, 2026 (#62). Its closeout head
`b4d6c02` is on main, so no history was lost; the closed milestone and the
[setup closeout record](evidence/2026-10-18-setup-closeout.md) remain the
setup baseline. Do not recreate or target the branch.

| Branch | Purpose and allowed changes |
| --- | --- |
| `release/v0.2.0` | Active integration branch for the SA prototype and repository maintenance, milestone 2 |
| `release/v0.3.0` | Integration branch for milestone 3 work, approved October 28, 2026 (#132) |
| `release/v0.1.0` | Superseded on October 27, 2026; receives no PRs; deleted under the lifecycle once milestone 2 is closed |
| `main` | Stable default branch; receives an integration branch only through a separate explicitly requested PR |
| Short-lived task branches | Branch from the integration branch of the issue's milestone; use focused names with one of the prefixes `feat/`, `fix/`, `docs/`, `chore/` or `test/`, such as `fix/<issue>-<slug>` |

Current prototype and maintenance task PRs target **`release/v0.2.0`**,
including Dependabot version updates. Milestone 3 task PRs target
**`release/v0.3.0`**. Any later milestone needs its own explicitly approved
integration branch. Task PRs are squash-merged after applicable
checks pass and review findings are resolved, all on the PR's final head:
record that head's SHA, its CI result and each review's kind, result and
link, and merge with the expected head SHA. A push that changes content after
a review needs a new review of the new head; a review of an earlier commit
does not carry over (#223). Do not commit development work
directly to main or the release branch. Check the base explicitly when opening
a PR: GitHub's default branch remains main.

Do not open or merge a release-to-main integration PR merely to report progress.
When integration is explicitly requested, open a PR from the applicable integration branch into
`main`, verify its diff, applicable checks and review findings, then merge with
a **merge commit** to preserve ancestry. Integration does not by itself publish
a version tag, release, workbook or website.

After integration, synchronize the active release branch through a reviewed
merge-commit PR from main, with strict required checks and resolved conversations.
No administrator or app bypass exists. Do not directly fast-forward the protected
branch. At a milestone transition, create the next explicitly approved integration
branch and update this document, contribution guidance and automated PR targets together.

### Release-branch lifecycle

A release branch is active while its milestone accepts work. Both release
rulesets protect it, including a separate deletion rule with no bypass. Once
it is **obsolete**, it is deleted (decision of October 18, 2026, #66).
A release branch is obsolete when all of these hold:

1. its milestone is closed;
2. its head is an ancestor of main, integrated through the explicitly requested
   merge-commit PR;
3. no open PR, Dependabot entry or workflow targets it;
4. a repository-wide search finds no current document that names it as
   active or tells contributors to use it. The search covers README,
   INSTALLATION, CONTRIBUTING, SECURITY, SCOPE, this document and the
   other `docs/` pages, the PR template and the issue forms. Dated evidence
   records and historical CHANGELOG entries are excluded.

To delete an obsolete branch:

1. The milestone closeout issue, or a follow-up issue, records the branch
   head, the check of each condition and the search used for the fourth.
2. With my explicit authorization for the deletion, I temporarily exclude
   only the obsolete branch from both release rulesets (25319621 and
   25319604), delete it, and immediately remove both exclusions.
3. The issue records the readback: the branch is absent, its head is on
   main, and both release rulesets again have no exclusions or bypass actors.

`release/v0.0.1` was deleted this way (#62). Never delete an active release
branch, force-push or reset shared history, or recreate a deleted release
branch name; later work uses a new explicitly approved branch.

### Integration verification

Before integration, record the release head and main head. After the
explicitly requested merge, inspect the resulting commit: its first parent must
be the previous main head and its second parent the reviewed release head.
Synchronize main back into the active release branch through a separate reviewed
merge-commit PR. Record both heads and verify the merged tree and ancestry;
if either branch moves, reassess the new work before proceeding. Never reset
or force an update.

Record the PR, merge commit, branch readback, and applicable post-merge checks.
PR [#242](https://github.com/danielep71/K-FRTB-ASA/pull/242) shows the current
procedure: it integrated `release/v0.2.0` into main with a merge commit, and
PRs [#243](https://github.com/danielep71/K-FRTB-ASA/pull/243) and
[#244](https://github.com/danielep71/K-FRTB-ASA/pull/244), from the task
branches `chore/sync-release-v0.2.0-242` and `chore/sync-release-v0.3.0-242`,
brought main back into `release/v0.2.0` and `release/v0.3.0` with merge
commits that keep both parents. Later task PRs advance the release branches
again; branches need match only at the integration checkpoint.

Historical example, before the synchronization rule of #151: after PR
[#13](https://github.com/danielep71/K-FRTB-ASA/pull/13), merge commit
`c01f2ad6c71ea2d6f23eac2b5a07fcdd430456a4` preserved both parents and the
retained release branch was fast-forwarded to it.
[CodeQL passed](https://github.com/danielep71/K-FRTB-ASA/actions/runs/39071612077).
Do not repeat that fast-forward.

## Issue metadata and completion

Every issue has assignee **danielep71**, exactly one priority label and a
milestone. P1 is reserved for confirmed calculation bugs, meaning wrong
results for valid inputs, and is used only with `bug`; setup improvements
normally use P3. Record the objective, acceptance checklist, dependencies and
required evidence. The [label guide](LABELS.md) owns the detailed taxonomy and
triage queries; these metadata requirements are manual triage rules.

Close an issue only when its criteria have linked evidence. Mark checkboxes
complete only when the corresponding work is verified. Keep blocked work
open with named dependencies. Static repository checks, actual runtime tests,
independent expected-value comparisons and methodology review are distinct
forms of evidence; do not claim one establishes the others.

### Synchronizing integration branches

Bring shared main changes into another active release through a task branch
and reviewed task PR. Resolve shared documents against both source heads,
preserving current requirements and replacing superseded statements. Record
both heads, the resulting commit and a content comparison. A squash merge
preserves the release's existing history but does not make the synchronized
main head an ancestor; never use that merge to claim ancestry advancement.

### Issue forms and template activation

The bug form defaults to `bug` and `P2`; feature and documentation forms
default to their type label and `P3`. All forms assign `danielep71`.
The bug and feature forms ask for the affected area, and the documentation
form for the document and section. Every form asks for the validation
evidence the change needs, and the dependency status in the form Issue triage
checks: issue links, an `External condition:` line, or None. The bug form
also asks for the full
commit SHA, a sourced expected value and the host environment; the feature
form asks for the compatibility impact. Forms cannot assign a milestone, so I
select it and confirm exactly one priority before completing triage. P1 remains limited to confirmed
calculation bugs. API/CLI-created issues and PRs bypass form defaults and need
explicit metadata: assign milestones to issues only, never to PRs.
Disabling ordinary blank issues does not enforce metadata
or prevent me from creating blank issues.

GitHub serves the issue chooser and PR template from the default branch.
Files merged into a release branch need a separate explicitly requested main
integration before default-branch activation. Initial forms and policies were
activated and verified through #21/#6; the prototype-aware updates
reached main through #58. The
[PR template](../.github/PULL_REQUEST_TEMPLATE.md) contains the single merge
checklist; this document owns its detailed procedures.

### Triage and closeout procedure

1. Inspect the title, objective, acceptance criteria, dependencies, labels,
   milestone, assignee, and discussion. Use the repository-wide searches in
   [Labels](LABELS.md), including closed issues and issues without milestones.
2. Assign `danielep71`, exactly one priority, and the milestone that owns the
   work. Apply the assignee and priority to development PRs, link their issue,
   and leave the PR milestone unset. Do not create a new
   milestone or infer authorization for functional scope from a version number.
3. Keep blocked criteria unchecked and name their dependency. Split unrelated
   scope into linked issues while retaining traceability.
4. Before merging, inspect the candidate diff, applicable hosted checks,
   review submissions, inline threads, and automated-review progress. Wait for
   any running review, address valid findings, and explain dismissals with
   evidence before resolving threads. Resolve stale findings against the
   current source. An author's self-check is not independent review.
5. After merging, link the PR, commit, successful applicable checks, and any
   required runtime or settings evidence in the issue. Mark only verified
   criteria complete. If validation failed or could not run, record it and
   keep the affected work open.
6. Close as completed only when the full acceptance scope is verified.
   For a duplicate, link its authoritative replacement and close as duplicate
   where supported, otherwise as not planned with the reason recorded.
   Declined or invalid work closes as not planned with a concrete explanation;
   its implementation checkboxes remain unchecked. Preserve required metadata.

### Milestone management

[v0.1.0 — Repo setup](https://github.com/danielep71/K-FRTB-ASA/milestone/1)
closed on October 18, 2026; its [detailed guide](Milestones/v0.0.1.md) and
[closeout record](evidence/2026-10-18-setup-closeout.md) remain the setup
baseline. [v0.2.0 — SA prototype](https://github.com/danielep71/K-FRTB-ASA/milestone/2)
is active: its [guide](Milestones/v0.2.0.md),
[interim evaluation](evidence/2026-10-27-prototype-evaluation.md) and
[Excel run](evidence/2026-10-31-excel-run-323260b.md) exist, prototype
closeout is tracked in #32, and #28 owns the remaining Excel checks, with
#72, #85 and #206. The methodology readings of #29 are decided and
implemented; v0.3.0 reviews them for the certified scope in the
[milestone plan](ROADMAP.md). The live issue tracker
owns progress; a guide is not a substitute for completed acceptance criteria.

**Milestones contain issues only**, including closed milestones. PRs link to
their owning issues and carry no milestone, whether open, closed, merged,
human-created or bot-created. My decision of October 27, 2026
(#104) supersedes earlier requirements to copy the issue milestone onto a PR.
Dated closeout records retain the policy and names in effect at that time.
The [metadata workflow](../.github/workflows/pr-milestones.yml) removes PR
milestones on PR events and reconciles all PRs daily or by manual dispatch.
It never changes an issue's milestone. See [Repository checks](CI.md) for
permissions, activation and operating limits.

Milestone scope and any next integration branch require explicit approval. Before a
milestone closes, its closeout issue must reconcile all associated work,
including closed issues and their linked PR evidence, confirm metadata and evidence, and record
adopted and deferred maintenance practices. For the setup milestone, issue #8 owns that record.
A genuine gap stays open or I explicitly respecify it; moving it
elsewhere does not make the original acceptance criterion pass.

Closing a milestone does not authorize integration into main, a tag, a GitHub
release, a workbook, a website, or a visibility change. Those decisions are
separate. [Releasing](../RELEASING.md) defines the release procedure and
version convention (#8).

## Current controls and maintenance

[Repository settings](REPOSITORY_SETTINGS.md) records the portable configuration.
[Local evidence](REPOSITORY_SETTINGS_EVIDENCE.md) records dated readbacks,
security-feature history and limitations. PRs and resolved
conversations are required; main allows merge commits; release branches allow
squash task merges and merge-commit synchronization PRs. No bypass actors exist. Required approvals are zero. The universal integrity workflow
(#7) is a required check on main and release branches (#2). The ChatGPT Codex
Connector reviews every PR automatically when it is opened; later pushes are
re-reviewed only on request (`@codex review`).

| Requirement | Enforcement or working procedure |
| --- | --- |
| PRs, merge methods, conversation resolution, shared-history and deletion protection | Active rulesets without bypass actors, recorded in the [settings](REPOSITORY_SETTINGS.md) and their [evidence](REPOSITORY_SETTINGS_EVIDENCE.md) |
| Release-branch deletion protection | Separate active ruleset with no bypass actors; an obsolete branch is deleted only through the lifecycle procedure above (#66) |
| Independent approval | No required approvals configured; my procedure must not claim that an independent review occurred |
| Successful universal integrity check | Required `Repository integrity` check, strict on main and release branches (#7, #2) |
| Review coverage for every PR | Automatic Codex review when a PR is opened, on main and release targets (#2); my review of later pushes |
| Issue/PR assignee, priority, acceptance evidence; issue milestone | Manual triage and completion procedure; the read-only **Issue triage** workflow audits open issues ([labels](LABELS.md)) |
| Label catalogue | Versioned `labels.json` (schema 2), validated on PRs; live reconciliation only from the default branch, never pruning a label still in use ([labels](LABELS.md)) |
| No milestone on any PR | Trusted metadata workflow on PR events, daily reconciliation and manual dispatch; Dependabot configuration is checked offline (#104) |
| Explicitly requested main integration and reviewed release-branch synchronization | Reviewed merge-commit PR with strict required checks and recorded source heads |

Issue #6 recorded initial default-branch activation of forms and the PR
template; later changes follow the same activation procedure. [Conduct](../CODE_OF_CONDUCT.md) and [security](../SECURITY.md)
policies define the reporting routes.
Dependabot targets Actions, the CI parser and prototype test packages at
`release/v0.2.0` without PR milestones, as detailed in [Repository checks](CI.md).
Issue #8 completed the setup closeout; its record is
[Setup closeout](evidence/2026-10-18-setup-closeout.md).

The procedures fit this project's prototype stage and release-branch
protections. This document records maintenance procedures, not Excel execution or
regulatory certification. Prototype source and LibreOffice evidence are
tracked separately from setup closeout.
