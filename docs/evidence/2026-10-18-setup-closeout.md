# ✅ v0.0.1 setup closeout record

Closeout record for issue #8 and the
[v0.0.1 — Repo setup](https://github.com/danielep71/K-FRTB-ASA/milestone/1)
milestone. Readback on October 18, 2026 (Europe/Rome), at main,
`release/v0.0.1` and `release/v0.1.0` = `d487ee293f6762f38041ed06dc055ac02b43adf9`.
`release/v0.0.1` was fast-forwarded to main by the owner before this record.

v0.0.1 is a repository setup baseline. It does not implement, run in Excel or
validate an FRTB-SA calculator; the prototype on main belongs to v0.1.0.

## 🧩 Workstreams

| Issue | Result | Merged PRs (merge commit) |
| --- | --- | --- |
| #1 Labels | 20 canonical labels; reconciliation and drift detection | #9 (`7d514b0`); pruning guard #33 (`8e7e710`) for #25 |
| #2 Settings | Rulesets, security features, required check, automatic review | #10 (`a03cbac`); settings changed and read back by the owner |
| #3 Identity and scope | README, scope, roadmap | #11 (`96a3b44`) |
| #4 Governance | Integration branch, contribution and milestone procedure | #12 (`4f6ee85`), #13 (`c01f2ad`), #16 (`78bf66f`) |
| #5 Structure | Directories, file policies, workbook-template path, MIT configuration | #14 (`a00f11e`), #15 (`00e1c30`), #17 (`f631670`) |
| #6 Contribution policies | Issue forms, PR template, conduct and security policies | #19 (`26312bd`), integrated by #21 (`359a1da`) |
| #7 CI | Universal `Repository integrity` check and Dependabot configuration | #35 (`d96f081`), integrated by #36 (`7ee0ced`) |
| #23 Traffic history | Protected data branch, daily collection | #24 (`588029d`), #27 (`9efa093`), integrated by #26 (`71d3e7a`) |
| #25 Label pruning defect | Pruning checks the resolved plan | #33 (`8e7e710`) |

Issues #1–#7, #23 and #25 are closed as completed.

## 🔎 Verification

| Check | Result |
| --- | --- |
| Labels | Live catalogue equals `.github/labels.json`: 20 of 20 names, colors and descriptions, none extra. [Drift run 39105486786](https://github.com/danielep71/K-FRTB-ASA/actions/runs/39105486786) on main passed. `labels-sync.mjs --mode validate` plans 0 changes |
| Metadata | All 20 issues and all 29 PRs are assigned to danielep71 and carry exactly one priority and a milestone; P1 is used only for the confirmed calculation bugs #41, #45 and #48 |
| Fresh checkout | Clone of `release/v0.0.1` at `d487ee2`. Each command passed: `tools/check_repository.py` (79 files plus label and traffic self-tests), `tools/vba_lint.py` (11 modules, 0 problems), the label self-tests and `REQUIRE_LIBREOFFICE=1 pytest` (97 passed, 21 subtests passed). Live drift detection needs a token locally; the hosted run covers it |
| Required check | Main ruleset 25319557 and release ruleset 25319604 require `Repository integrity` (GitHub Actions, app 15368); main also requires up-to-date branches. Dependabot PR #37 was `blocked` only by the missing check |
| Review | The ChatGPT Codex Connector reviewed #42 and #46 automatically when they were opened |
| Security | CodeQL default setup covers GitHub Actions and Python (default suite, weekly). Secret scanning and push protection are enabled, private vulnerability reporting is enabled and Dependabot security updates are off, as documented. Dependabot version updates opened #37. Owner readback of the Security tab: 0 open code-scanning alerts on main, 0 Dependabot alerts, 0 secret-scanning alerts and no security advisories |
| Repository | Description and topics updated by the owner to describe the Excel/VBA prototype with its pending Excel evidence; public; default branch main; Pages and Discussions off; automatic deletion of merged branches on; merge methods enabled globally and narrowed by rulesets |
| Forms and policies on main | Three issue forms, `config.yml` with blank issues disabled and the private-reporting link, PR template, `SECURITY.md` and `CODE_OF_CONDUCT.md` |
| License | `LICENSE` is MIT with copyright Daniele Penza; GitHub detects MIT. README, contribution and policy documents link it, and the maintenance scripts carry `SPDX-License-Identifier: MIT` |
| Last setup-only commit `7ee0ced` | The #36 merge, 45 files, no VBA modules or workbook. It predates the closeout (no release procedure or closeout record; its documents describe completed controls as pending), so it is not a tag candidate. Fresh checkout of that commit: `tools/check_repository.py` passed (45 files plus label and traffic self-tests), and `labels-sync.mjs --mode validate` plans 0 changes. Hosted [Repository integrity run 39077036863](https://github.com/danielep71/K-FRTB-ASA/actions/runs/39077036863) passed on it |
| Wiki | Enabled for collaborators. The owner published a Home page (wiki commit `3725aa9`) that links to the versioned repository documents; it holds no other content |
| Traffic history | [Run 39105087826](https://github.com/danielep71/K-FRTB-ASA/actions/runs/39105087826) published data commit `af6a68a`; the first window has zero counts |

## 🧭 Adopted and deferred practices

| Practice | Decision |
| --- | --- |
| Label catalogue, reconciliation and drift detection | Adopted |
| Branch and tag rulesets, required universal check | Adopted |
| Automatic Codex review on PR open | Adopted; later pushes are re-reviewed on request |
| CodeQL default setup, secret scanning, push protection, private vulnerability reporting | Adopted |
| Dependabot version updates | Adopted, targeted at the integration branches |
| Issue forms, PR template, conduct and security policies | Adopted |
| Traffic history | Adopted; daily collection active |
| Release procedure and version convention | Adopted in [Releasing](../../RELEASING.md) |
| Required approvals | Not adopted: one maintainer; the procedure must not claim independent review |
| Copilot code review | Not available for this account |
| Dependabot security update PRs | Deferred; alerts remain enabled |
| Auto-merge | Off |
| Pages, SEO and repository social preview | Deferred unless separately authorized |
| Tag `v0.0.1` and GitHub release | Deferred by the owner on October 18, 2026. No main commit is both setup-only and complete, so there is no tag candidate; see [Releasing](../../RELEASING.md) |

## ⚠️ Limitations

- Codex does not re-review later pushes automatically.
- The traffic token expires; collection fails until it is renewed.
- Security alert lists are not readable with the automation token used for
  this record (HTTP 403); the zero-alert result is the owner's readback.
- The wiki is a separate repository that the owner publishes. It is an entry
  point only; the versioned documents remain authoritative.
- Main also carries the v0.1.0 prototype. Its LibreOffice evidence is not
  Excel evidence (#28) and its round trip (#30) are open under v0.1.0; its
  methodology decisions (#29) move to v0.2.0 under the approved milestone
  plan (#55).

## ➡️ Next decisions

1. Integrate the #8 closeout PR into main (owner request), then fast-forward
   `release/v0.0.1` and recheck the issue forms from main.
2. Retire the setup line: Dependabot still targets `release/v0.0.1` and
   milestone v0.0.1 for GitHub Actions and the CI parser, and governance and
   agent instructions list the branch as active (#51).
3. Tag `v0.0.1`: deferred by the owner. No main commit is both setup-only
   and complete, so a tag would first need an owner decision on preparing a
   complete candidate (see [Releasing](../../RELEASING.md)). Until then the
   closed milestone and this record are the setup baseline.
4. v0.1.0 closeout (#32), after #28 and #30; the methodology decisions
   (#29) belong to v0.2.0 ([roadmap](../ROADMAP.md)).

## 🗒️ Addendum — October 18, 2026, after closeout

This addendum records later outcomes; the readback above is unchanged.
Readback at main and `release/v0.1.0` = `a9ee4bbe251ce7ebdaf771a4aab42b850c4584b2`,
`release/v0.0.1` = `b4d6c02edcd60bb412a06999d1909e6eaba63837`.

| Next decision | Outcome |
| --- | --- |
| 1. Integrate the #8 closeout | #53 merged into main (`b4d6c02`); the owner fast-forwarded `release/v0.0.1` to it. The issue forms are served from main |
| 2. Retire the setup line | #51 closed: #57 (`b40ba5c`) retargeted all Dependabot entries to `release/v0.1.0`, milestone 2, and aligned the branch guidance; #58 (`a9ee4bb`) integrated it into main and the owner fast-forwarded `release/v0.1.0`. `release/v0.0.1` is historical, protected and closed to new work |
| 3. Tag `v0.0.1` | Still deferred by the owner; no tag exists |
| 4. v0.1.0 closeout | Open (#32), after #28 and #30 |

The approved milestone plan was recorded in the [roadmap](../ROADMAP.md)
through #56 (`b0f5c73`), closing #55. Milestone v0.2.0 exists and holds the
methodology decisions (#29).

Later on October 18, 2026 the owner deleted `release/v0.0.1` (#62). Its
head `b4d6c02` remains on main, so the setup history is unchanged.
