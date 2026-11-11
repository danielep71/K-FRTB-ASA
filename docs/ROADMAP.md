# Roadmap

This roadmap records the milestone plan approved on October 18, 2026,
from the repository setup to the first certified release. A milestone's due
date on GitHub is a planning target, not a release promise; issue triage
checks each issue's target date against it ([labels and triage](LABELS.md)).
Each milestone after v0.3.0 is created on GitHub
when its predecessor has fixed its scope, as [governance](GOVERNANCE.md)
requires; the issue tracker is authoritative for current progress.

On October 27, 2026 the milestones were renumbered one step up: the setup
milestone, first named v0.0.1, became v0.1.0, and every later milestone moved
with it. The integration branch `release/v0.2.0` superseded `release/v0.1.0`
the same day (#117). Dated records
keep the names in effect when they were written.

## 🗺️ Milestone plan

| Milestone | Objective | Exit condition |
| --- | --- | --- |
| ✅ v0.1.0 — Repo setup | Repository foundations: governance, protections, CI, issue management and the release procedure | Required checks active and release procedure verified; analytics kept separate from the certification path. Closed October 18, 2026 |
| 🧪 v0.2.0 — SA prototype | Consolidate the existing prototype: SBM for the five implemented risk classes, non-securitization DRC and RRAO | Excel build and run documented, sample reconciliation, source round trip verified and limits stated explicitly |
| 📐 v0.3.0 — Regulatory baseline and certification scope | Fix exactly what will be certified: standard and version, any CRR profile, included products, exclusions and interpretations | Approved requirement → source → formula → parameter → test matrix; the methodology decisions the scope needs are resolved; the meaning of certification, the validator's role, the target platform and the supported Excel environments are recorded |
| ⚙️ v0.4.0 — Calculation engine completion | Complete and correct the engine for the approved scope, progressively separating calculation from worksheet access | Every included requirement implemented with independently derived expected values, input/output contracts defined and no known calculation defect open |
| 🗂️ v0.5.0 — Workbook and user workflow | Guided workbook use: inputs, parameters, default restore, run, diagnostics and results | Full flow verified in Excel; understandable errors, no stale results, safe application state and documented handling of confidential portfolio data |
| 🔍 v0.6.0 — Independent validation | Validate formulas, parameters and results through a review independent of the implementation | Intermediate and final results reconciled, official cases where available, independently derived cases, and anomalies resolved or formally excluded from scope |
| 🛡️ v0.7.0 — Operational hardening | Robustness, performance and compatibility on the supported Excel environments | Malformed inputs, representative portfolios, repeated runs, error recovery and agreed run times all pass |
| 📦 v0.9.0 — Release candidate and certification dossier | Freeze the candidate and prepare the certification dossier | Rebuildable package, updated manual and limits, full regression suite on the candidate and final review complete; any later change needs the relevant checks again |
| 🏁 v1.0.0 — First certified release | Approve and publish the first certified version | Methodological, numerical and technical approval on the same commit and workbook; tag, package (rebuildable from template and sources, with recorded digests), evidence and release notes consistent |

v0.8.0 is reserved for an iteration the hardening or validation work may
require. Independent expected values accompany every calculation increment
from v0.4.0 onward; v0.6.0 is the formal review by someone other than the
implementer, not the first check. Every milestone names its own Excel
evidence, which only a Windows desktop Excel host can produce. Versions,
tags and release evidence follow [Releasing](../RELEASING.md).

## ✅ v0.1.0 — Repo setup (closed)

[Milestone](https://github.com/danielep71/K-FRTB-ASA/milestone/1) ·
[milestone guide](Milestones/v0.0.1.md) ·
[closeout record](evidence/2026-10-18-setup-closeout.md)

| Issue | Deliverable |
| --- | --- |
| [#1](https://github.com/danielep71/K-FRTB-ASA/issues/1) | Canonical labels, safe reconciliation, drift detection, and triage rules |
| [#2](https://github.com/danielep71/K-FRTB-ASA/issues/2) | Effective GitHub settings, protection, security, required check and review verification |
| [#3](https://github.com/danielep71/K-FRTB-ASA/issues/3) | Project identity, accurate scope, navigation, and starting instructions |
| [#4](https://github.com/danielep71/K-FRTB-ASA/issues/4) | Governance, integration branches, contribution rules, and milestone guide |
| [#5](https://github.com/danielep71/K-FRTB-ASA/issues/5) | Minimal source, test, documentation, and tooling structure |
| [#6](https://github.com/danielep71/K-FRTB-ASA/issues/6) | Issue forms, PR template, and contribution/security policies |
| [#7](https://github.com/danielep71/K-FRTB-ASA/issues/7) | Universal repository-integrity checks and dependency updates |
| [#8](https://github.com/danielep71/K-FRTB-ASA/issues/8) | Release procedure and evidence-based setup closeout |
| [#23](https://github.com/danielep71/K-FRTB-ASA/issues/23) | Traffic history: protected data branch and daily collection |
| [#25](https://github.com/danielep71/K-FRTB-ASA/issues/25) | Label pruning guard |

The milestone closed without a tag or release. No main commit is both
setup-only and complete, so the closed milestone and the closeout record are
the setup baseline.

## 🧪 v0.2.0 — SA prototype

[Live milestone](https://github.com/danielep71/K-FRTB-ASA/milestone/2) ·
[milestone guide](Milestones/v0.2.0.md) ·
[interim evaluation](evidence/2026-10-27-prototype-evaluation.md) ·
[Excel run](evidence/2026-10-31-excel-run-323260b.md),
integration branch `release/v0.2.0` (`release/v0.1.0` until October 27, 2026).

The table below is the readback of October 27, 2026, kept as it was. Since
then, #20, #30, #79 and #82 have closed with the
[Excel run of October 31, 2026](evidence/2026-10-31-excel-run-323260b.md) on
`323260b`; the remaining manual Excel checks are in #28, #72, #85 and #206.

| Issue | State on October 27, 2026 | Deliverable |
| --- | --- | --- |
| [#20](https://github.com/danielep71/K-FRTB-ASA/issues/20) | Open | Excel VBA workbook prototype: SBM (delta, vega, curvature) for five risk classes, DRC non-securitization, RRAO, independent Python reference and reconciliation |
| [#28](https://github.com/danielep71/K-FRTB-ASA/issues/28) | Open | Excel host build, compile, execution and evidence |
| [#30](https://github.com/danielep71/K-FRTB-ASA/issues/30) | Open | Committed template and modules layout; remaining Excel round trip |
| [#32](https://github.com/danielep71/K-FRTB-ASA/issues/32) | Open | Milestone guide, evaluation and evidence closeout |
| [#31](https://github.com/danielep71/K-FRTB-ASA/issues/31) | Closed | Bring `release/v0.1.0` up to date with main and reconcile governance |
| [#40](https://github.com/danielep71/K-FRTB-ASA/issues/40) | Closed | Remaining prototype documentation inconsistencies |
| [#41](https://github.com/danielep71/K-FRTB-ASA/issues/41) | Closed | Reject invalid required parameters; consistent numeric-text parsing |
| [#45](https://github.com/danielep71/K-FRTB-ASA/issues/45) | Closed | Reject blank or invalid parameter-table values and vertex lists; clear outputs on failure |
| [#48](https://github.com/danielep71/K-FRTB-ASA/issues/48) | Closed | Reject blank keys and gaps in required parameter tables |
| [#51](https://github.com/danielep71/K-FRTB-ASA/issues/51) | Closed | Retire the setup line: Dependabot targets and branch guidance |
| [#55](https://github.com/danielep71/K-FRTB-ASA/issues/55) | Closed | This milestone plan |
| [#59](https://github.com/danielep71/K-FRTB-ASA/issues/59) | Closed | Post-closeout documentation consistency |
| [#62](https://github.com/danielep71/K-FRTB-ASA/issues/62) | Closed | Historical setup-branch deletion and corrected guidance |
| [#66](https://github.com/danielep71/K-FRTB-ASA/issues/66) | Closed | Obsolete release-branch lifecycle |
| [#68](https://github.com/danielep71/K-FRTB-ASA/issues/68) | Closed | Remove duplicated AGENTS.md guidance |
| [#72](https://github.com/danielep71/K-FRTB-ASA/issues/72) | Open | VBA synchronization implemented; desktop Excel and Windows verification pending #28 |
| [#76](https://github.com/danielep71/K-FRTB-ASA/issues/76) | Closed | Saved template style and source-build print areas |
| [#77](https://github.com/danielep71/K-FRTB-ASA/issues/77) | Closed | Synchronization-tool documentation |
| [#79](https://github.com/danielep71/K-FRTB-ASA/issues/79) | Open | Read-only export and comparison implemented; actual Excel round trip pending #28/#30 |
| [#82](https://github.com/danielep71/K-FRTB-ASA/issues/82) | Open | Evidence contract and validator implemented; actual Excel bundle pending #28 |
| [#85](https://github.com/danielep71/K-FRTB-ASA/issues/85) | Open | Output-placeholder fix verified in LibreOffice; desktop Excel verification pending #28 |
| [#87](https://github.com/danielep71/K-FRTB-ASA/issues/87) | Closed | Repository-specific comments and supported file policies |
| [#93](https://github.com/danielep71/K-FRTB-ASA/issues/93) | Closed | Maintenance-audit documentation reconciliation |
| [#95](https://github.com/danielep71/K-FRTB-ASA/issues/95) | Closed | Repository references after the rename to K-FRTB-ASA |
| [#98](https://github.com/danielep71/K-FRTB-ASA/issues/98) | Closed | Market risk and CVA method maps |
| [#100](https://github.com/danielep71/K-FRTB-ASA/issues/100) | Closed | Project website, live on GitHub Pages |
| [#102](https://github.com/danielep71/K-FRTB-ASA/issues/102) | Closed | Market risk framework diagram as a map of approaches |
| [#104](https://github.com/danielep71/K-FRTB-ASA/issues/104) | Closed | Milestones for issues only |
| [#110](https://github.com/danielep71/K-FRTB-ASA/issues/110) | Closed | Root README brought up to date |
| [#112](https://github.com/danielep71/K-FRTB-ASA/issues/112) | Closed | Documentation aligned with the renumbered milestones |
| [#113](https://github.com/danielep71/K-FRTB-ASA/issues/113) | Closed | Prototype version `0.2.0-prototype` |
| [#115](https://github.com/danielep71/K-FRTB-ASA/issues/115) | Closed | MAR paragraph references corrected |
| [#117](https://github.com/danielep71/K-FRTB-ASA/issues/117) | Closed | Integration moved to `release/v0.2.0` |
| [#118](https://github.com/danielep71/K-FRTB-ASA/issues/118) | Closed | Dependabot targets `release/v0.2.0` |
| [#124](https://github.com/danielep71/K-FRTB-ASA/issues/124) | Closed | Timeout on the LibreOffice install step |
| [#126](https://github.com/danielep71/K-FRTB-ASA/issues/126) | Closed | Preview image version 0.2.0 and README wiki link |

States were read back on October 27, 2026; the live milestone is authoritative.
At that checkpoint main and `release/v0.2.0` both pointed to `46ae228`, after
integrations #123 (version 0.2.0, the branch move and the website) and #128
(install-step timeout, preview image and wiki link). These merges do not
supply desktop Excel evidence. The [milestone guide](Milestones/v0.2.0.md)
and an [interim evaluation](evidence/2026-10-27-prototype-evaluation.md)
exist; the final closeout record follows the Excel run of #28 (#32).

I decided the six methodology interpretations of
[#29](https://github.com/danielep71/K-FRTB-ASA/issues/29) on October 27, 2026;
they are implemented (#134, #136, #138, #140, #142, #144) and recorded in the
[methodology](methodology.md#interpretations-and-simplifications-to-review).
They are my readings, not independently reviewed; v0.3.0 reviews them for
the certified scope. The prototype informs the v0.3.0 scope and platform
decisions; it does not complete the later milestones' exit conditions.

## 📐 v0.3.0 — Regulatory baseline and certification scope

[Live milestone](https://github.com/danielep71/K-FRTB-ASA/milestone/3),
integration branch `release/v0.3.0` (#132). It holds the methodology
decisions of
[#29](https://github.com/danielep71/K-FRTB-ASA/issues/29), decided on October
27, 2026 and implemented; its exit condition, including their review for the
certified scope, is in the plan above.
