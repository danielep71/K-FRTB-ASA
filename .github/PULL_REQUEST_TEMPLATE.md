## 🎯 Objective and scope

Related issue:
Explain the problem and resulting behavior. State non-goals and compatibility
impact (inputs, outputs, defaults, errors, formats, supported environments).

## 🧪 Validation

Record the exact source commit, commands, results, and relevant run/evidence
links. Distinguish static checks, runtime execution, independent numerical
comparison, and methodology review. State anything unavailable or skipped.

## 🔍 Review findings and limitations

Link findings and explain their resolution. Identify unresolved findings,
risks, dependencies, or missing evidence; write None when applicable.

## ✅ Merge checklist

- [ ] Base is `release/v0.2.0` for current prototype/maintenance tasks (milestone 2, v0.2.0 — ASA prototype), `release/v0.3.0` for milestone 3 tasks, another explicitly approved branch for later milestones, or this is an explicitly requested release-to-main integration. Do not target the superseded `release/v0.1.0` or recreate the deleted setup branch `release/v0.0.1`.
- [ ] Linked issue and PR have assignee `danielep71` and exactly one priority; P1 is only for a confirmed calculation bug (wrong results for valid inputs), with `bug`. The linked issue has the appropriate milestone; the PR has no milestone.
- [ ] Scope, compatibility impact, affected documentation, and CHANGELOG.md are current.
- [ ] Applicable checks pass and the automated review is finished on the final head SHA, with every finding addressed or explained and resolved; a push that changes content after a review gets a new review.
- [ ] Evidence and limitations are accurate; no credentials, confidential data, or undisclosed vulnerability details are included.

Task PRs squash-merge into the integration branch. Explicitly requested main
integrations use a merge commit. If work continues, synchronize the release
branch through a reviewed merge-commit PR from main; after the
milestone-final integration, delete the obsolete branch under the
[release-branch lifecycle](https://github.com/danielep71/K-FRTB-ASA/blob/main/docs/GOVERNANCE.md#release-branch-lifecycle).
