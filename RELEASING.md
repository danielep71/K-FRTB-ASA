# 🚀 Releasing

[![Versioning: SemVer](https://img.shields.io/badge/versioning-SemVer-1f3a5f)](#version-convention)
[![Published releases: none](https://img.shields.io/badge/published%20releases-none-6E7781)](#current-stage)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

This procedure covers versions, the changelog, integration into main, tags
and the evidence each stage needs. [Governance](docs/GOVERNANCE.md) owns
branches and milestones; this document owns what turns a milestone into a
version.

<a id="current-stage"></a>

## 📍 Current stage

No version has been tagged or published. v0.1.0, first named v0.0.1, is the
**repository setup baseline**: labels, governance, settings, forms, CI and
traffic analytics. It is not a working or regulatory-validated calculator.
v0.2.0 is the Excel/VBA **ASA prototype**; its LibreOffice evidence exists,
and one desktop Excel run, with both round trips, is recorded on `323260b`, on
one host ([Excel run](docs/evidence/2026-10-31-excel-run-323260b.md)). The
remaining Excel checks are in #28, #72, #85 and #206; `tools\Check-FRTB-Excel.cmd`
(#274) runs those of #28, #72 and #85 unattended, but it has not yet run in
desktop Excel. I decided its six
methodology interpretations (#29) on October 27, 2026, and they are
implemented; they are not independently reviewed, and v0.3.0 reviews them for
the certified scope under the [roadmap](docs/ROADMAP.md).

The milestones were renumbered one step up on October 27, 2026, and the
workbook reports the matching version `0.2.0-prototype` (#113). Dated records
keep the names and versions in effect when they were written.

<a id="version-convention"></a>

## 🔢 Version convention

| Item | Convention | Example |
| --- | --- | --- |
| Version number | [Semantic Versioning](https://semver.org/) `MAJOR.MINOR.PATCH`; before 1.0.0, MINOR marks a functional stage and PATCH a maintenance increment | `0.2.0`, `0.2.1` |
| Pre-release | SemVer pre-release identifier | `0.2.0-prototype` |
| Milestone | Version plus a short title | `v0.2.0 — ASA prototype` |
| Integration branch | `release/` plus the tag name; `release/v0.2.0` superseded `release/v0.1.0`, named before the renumbering (#117) | `release/v0.2.0` |
| Tag | `v` plus the version, annotated, on a main commit | `v0.2.0` |
| Changelog section | Version and release date | `## [0.2.0] - YYYY-MM-DD` |

One convention applies to every line. Do not tag without the `v` prefix,
reuse a number, or tag a release branch.

## 📌 Version source

- **The tag** is the authoritative record of a released version once it
  exists. Tags `v*` are protected against update, deletion and force pushes
  ([settings](docs/REPOSITORY_SETTINGS.md)); a wrong tag is corrected by a new
  PATCH version, never by moving it.
- **The changelog** section heading must name the same version and date.
- **The workbook** reports its own version in `APP_VERSION`
  (`src/modules/CORE_Config.bas`); it must equal the tagged version, or carry
  the same version with a pre-release identifier, when a workbook version is
  tagged.

## 📝 Changelog

[CHANGELOG.md](CHANGELOG.md) collects integrated changes under
`[Unreleased]`. When a release is explicitly requested, a task PR moves the
entries that belong to that version into a dated version section. Main can
carry work from more than one line: move only the entries of the released
milestone and leave the others under `[Unreleased]`.

## 🔀 Integration and tagging

1. **Task PRs** squash-merge into the integration branch of their issue's
   milestone after the required `Repository integrity` check and the
   applicable checks pass and review findings are resolved.
2. **Milestone closeout**: the closeout issue reconciles every issue and PR
   and links the evidence below. Closing the milestone does not merge main,
   create a tag or publish a release.
3. **Integration into main** happens only on explicit request,
   through a release-to-main PR with a merge commit. While work continues,
   main is then brought back into the release branch through a reviewed
   merge-commit PR; the branch is never fast-forwarded directly. Once its
   milestone is closed and fully integrated, the branch is obsolete and is
   deleted under the [release-branch lifecycle](docs/GOVERNANCE.md#release-branch-lifecycle).
4. **Tag candidate**: the main commit whose whole tree both satisfies the
   version's contents rule below and carries the version's completion
   record and dated changelog section. On explicit request, a task PR
   dates the section and it reaches main through step 3. A tag exposes the
   entire tree, so release notes can neither exclude files from it nor add
   the missing closeout to it. Git source archives of the tag omit only the
   `export-ignore` paths in `.gitattributes` (`.gitattributes`, `.gitignore`,
   `.editorconfig` and `.github/`), as [Installation](INSTALLATION.md) notes.
   - If no main commit meets both conditions, there is no tag candidate.
     The version is then recorded by its closed milestone and closeout
     record. Tagging it needs a separate explicit decision on how a complete
     candidate is prepared and validated.
5. **Validate the exact candidate**: run the stage's checks on a fresh
   checkout of that commit and confirm that the hosted `Repository integrity`
   run passed on it. Record the commands, results and run URL in the closeout
   or release notes. Evidence from another commit does not count.
6. **Tag**: on explicit request, create an annotated tag on the
   validated candidate, then push it:

   ```sh
   git fetch origin
   git tag -a v0.2.0 <candidate-commit> -m "v0.2.0 — ASA prototype"
   git push origin v0.2.0
   ```

7. **GitHub release** (optional, on request): notes that link the closeout
   issue and evidence. Attach no workbook binary unless the version's
   evidence covers it.

**The setup milestone has no tag candidate.** Main carried the prototype
(#39) before the setup closeout:
- `7ee0ced` (#36), the last setup-only main commit, predates the closeout:
  it has no release procedure or closeout record, and its documents still
  describe completed controls as pending.
- Every later main commit carries the prototype source and template.

The setup baseline is recorded by the closed v0.1.0 — Repo setup milestone and the
[setup closeout record](docs/evidence/2026-10-18-setup-closeout.md). The
tag was deferred on October 18, 2026.

## 📦 Contents and evidence by stage

| Stage | Release contents | Evidence required before tagging |
| --- | --- | --- |
| Setup (v0.1.0) | Source tree only; no binaries or workbooks | Closeout issue with merged PRs, green `Repository integrity` and label-drift runs, live settings readback, fresh-checkout verification and metadata audit |
| Prototype (v0.2.x) | Source tree and the committed macro-free template; no built `.xlsm` | Setup evidence, plus the LibreOffice reconciliation suite, Excel host build, compile and regression evidence (#28), the Excel round trip (#30) and the methodology interpretations decided and documented (#29; reviewed for the certified scope in v0.3.0) |
| Functional releases | Defined by the milestone that introduces them | Prototype evidence, plus independent benchmarks, the supported-host matrix and any packaging approved in [Scope](docs/SCOPE.md) |

Evidence is bound to a commit: it names the PR or commit, commands, results,
hosted run URLs, the settings readback date and anything not run. Static
checks do not prove runtime or numerical correctness, and LibreOffice runs
are not Excel evidence. Never fabricate workbook or Excel results.

## 🚫 Not part of a release

A tag or release does not change repository visibility, publish a website or
certify regulatory compliance. Each of those needs a separate explicit decision.
