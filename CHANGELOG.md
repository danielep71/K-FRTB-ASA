# Changelog

Record user-visible and maintainer-visible changes here. Entries under
Unreleased describe integrated work, not a published version.

## [Unreleased]

### Changed

- Prepared a reviewed wiki update for the first recorded desktop Excel run,
  with its exact candidate, one-host scope and remaining manual checks. The
  patch is under `docs/wiki-updates/`; wiki publication remains separate
  (#32).
- Dependabot security updates are on. The settings document, its evidence and
  CI record the change and how a security update PR, which opens against
  main, is retargeted to the active release branch (#191).
- GOVERNANCE, LABELS, CI, SECURITY, the Code of Conduct and the v0.2.0
  milestone guide state my roles and decisions in the first person, and the
  label table matches the catalogue descriptions. Dated evidence and closed
  records are unchanged (#187).
- The traffic history moves to CSV files: views and clones per day, one
  snapshot per collection day with the repository's stars, forks, watchers
  and open issues, and the top referrers and paths. The workflow adds the
  repository request to the App-token curl step, merges with
  `tools/traffic_history.py`, publishes atomically with
  `tools/traffic_publish.py`, writes three README badges, and opens a
  traffic alert issue when traffic spikes, stars or forks grow, or a new
  referrer appears. `import-legacy` replays the former JSON snapshots (#185).
- `.gitattributes`, `.gitignore` and `.editorconfig` are the complete file
  baseline: Git attributes for PDFs, all Office packages, fonts, media and
  archives; ignore rules for Office and Access locks, coverage, packaging and
  Node caches; editor indentation for every text type. Every tracked file
  keeps the same Git attributes, nothing is renormalized, and the template
  stays the only tracked Office package. No regulatory text is stored; the
  files state the rule for a byte-for-byte copy (#179).
- Migrate traffic collection to an isolated GitHub App token and curl-only
  API step; retain validated atomic publication with the workflow token (#181).

- The issue forms ask for the affected area, the validation evidence and the
  dependency status in the form Issue triage checks; the bug form also asks
  for the full commit SHA, a sourced expected value and the Excel or
  LibreOffice environment, and the feature form for the compatibility impact.
  The outdated branch history is removed from their opening note. The labels
  and the assignee are unchanged (#177).
- The settings evidence, the Excel evidence guide and the market-risk diagram
  notes describe only this repository, and the #168 entry below links to
  `AGENTS.md` instead of restating its rules (#174).
- The settings document names the default-branch ruleset "Protect main", as it
  is live, states that the traffic-history rule blocks deletion and force
  pushes while direct pushes stay allowed for the traffic workflow, and
  records the security features: Dependabot alerts, private vulnerability
  reporting, secret scanning and push protection on, Dependabot security
  updates off (#167).
- Six label descriptions are reworded in the first person, and P1 now reads
  "Confirmed calculation bug: wrong results for valid inputs", used only with
  `bug`, consistently in the catalogue, the PR template, the issue forms,
  LABELS and GOVERNANCE. Closed issues #45 and #48 and their PRs #46 and #49,
  about invalid inputs, move from P1 to P2 (#162).
- The PR milestone reconciliation runs only for PR events, the schedule, or a
  push or dispatch on the default branch, so a dispatch on another branch no
  longer runs that branch's workflow with `pull-requests: write`;
  `tools/check_repository.py` requires the condition (#160).
- The label catalogue uses schema 2, a single `labels` list without overlays,
  and the unused label policy file is removed. The label scripts are replaced
  by schema-2 versions that keep the guard against pruning a used label and
  write one summary when the final check fails. Live reconciliation runs only
  from the default branch, and the validation job has a 5-minute timeout
  (#158).
- `FX_CURVATURE_DIVISOR` divides only the FX CVRs of options that do not
  reference the reporting currency, which a new `RefReportingCcy` column on
  `In_Curvature` marks (MAR21.98). The new `FX_CURVATURE_DIVISOR_ALL` divides
  every FX CVR: the alternative that needs supervisory approval, labeled as
  such. An unrecognized flag rejects the row. The sample's USD/EUR row is
  flagged; with both switches off, the sample results are unchanged.
  Compatibility: the column moves `Comment` on `In_Curvature` one column to
  the right, so a workbook built earlier needs a rebuild or an inserted
  column; until then its comments are rejected as flags, never applied
  (#144, from the #29 decisions).
- The methodology states the curvature input contract the upstream pricing
  must follow: shock sizes (MAR21.98, MAR21.99), the delta deduction, the
  loss-positive sign, entry before the FX divisor, and no position on the
  sqrt(2) discretions for curvature shocks. It also records that a `P_CSR`
  override of the covered-bond risk weight applies to the whole bucket. The
  `In_Curvature` and `P_CSR` bucket 8 descriptions follow; no calculation
  changed (#140, from the #29 decisions).
- Import the shared repository-settings target and separate local/historical
  evidence, recording verified live alignment and PR-based synchronization (#149).
- The PR template, the release procedure, the v0.2.0 milestone guide and the
  governance controls table describe release-branch synchronization through a
  reviewed merge-commit PR from main, without a direct fast-forward or a
  ruleset bypass (#151).
- `release/v0.3.0`, created on October 28, 2026 from `69e3df3`, is the
  integration branch for milestone 3, starting with the methodology
  decisions of #29. `release/v0.2.0` keeps the prototype closeout,
  repository maintenance and Dependabot. Governance, contributing and the
  roadmap name both lines (#132).
- The prototype milestone guide names its assignee directly and limits the
  input-validation statement to the required-parameter cases covered by the
  recorded regression evidence (#32).
- The social preview image shows "PROTOTYPE 0.2.0", rendered again from its
  HTML source with the figures unchanged, and the root README navigation links
  the wiki after the website (#126).
- Dependabot version updates for GitHub Actions and both pip directories target
  `release/v0.2.0`; activation requires this configuration on `main` (#118).

- Milestones contain issues only. Historical PR assignments are removed;
  trusted metadata automation reconciles future PRs and bot changes. The PR
  checklist, contribution rules and Dependabot entries follow the same policy,
  with offline regression coverage (#104).

### Fixed

- A DRC obligor whose positions name different buckets is logged as an error
  and skipped with all its positions, instead of taking the first position's
  bucket (MAR22.22). When its positions differ in credit quality, the worst
  one, with the highest risk weight, applies with a warning, instead of the
  first one (MAR22.24). The reference, the methodology and the tests follow;
  the sample results are unchanged (#142, from the #29 decisions).
- With `GIRR_SPECIFIED_CCY_DISCRETION` on, the sqrt(2) division of MAR21.44
  also applies to the inflation and cross-currency basis risk weights of a
  specified currency or the reporting currency, not only to the tenor
  weights. The Control description, the methodology and the reference follow;
  the sample results are unchanged because the discretion is off by default
  (#136, from the #29 decisions).
- A cash-equity DRC position (`EQUITY` seniority) with a maturity between
  three months and one year is rejected and logged instead of being scaled:
  MAR22.16 and its FAQ allow only three months or at least one year. The
  `In_DRC` description, the methodology and the reference follow; the
  sample results are unchanged (#134, from the #29 decisions).
- The LibreOffice installation step of the prototype checks stops after 8
  minutes, so a hung package mirror fails the job early instead of holding it
  for the whole 20-minute limit (#124).
- MAR paragraph references now cite the paragraph that holds each rule
  (#115): netting of sensitivities MAR21.4(2) (In_Delta subtitle, CORE_SBM,
  methodology); commodity tenors MAR21.13(1) and vega maturities MAR21.8(4)
  (P_General); the DRC maturity floor MAR22.18 and one-year horizon MAR22.15
  (P_DRC); the RRAO weights MAR23.8(2)(a) and (b) (P_RRAO); the DRC steps in
  CORE_DRC; sensitivity definitions MAR21.19-26 (README sheet, methodology).
  The template is rebuilt; no parameter value or result changes.
- Workbook presentation now survives template generation: the XLSX saves
  hidden gridlines, role-colored tabs, frozen headers and identifiers, and
  landscape print settings. Data-row spacing, table borders, README
  separators, readable display labels and initial output headers follow the
  shared toolkit style. VBA formatting and the headless build are updated
  together, with saved-presentation and regeneration checks (#76).
- **Clear outputs** and a failed run no longer remove the output sheet
  headers. Results, the Out_* sheets and Log return to the template's empty
  state: headers in row 5 and the note in row 6, with "Run failed" on
  Results after a failure. The build, Clear outputs and the failure path
  share one placeholder routine, `CORE_Util.WriteOutputPlaceholder` (#85).

- Tagged parameter tables reject blank keys and internal blank rows rather
  than truncating later matrix data and leaving correlations at zero (#48).

- Required numeric parameters fail with key/source context instead of using
  zero for invalid text. Validation and conversion share a full invariant
  dot-decimal/scientific grammar; comma/grouped/junk/overflow text is rejected
  while actual numeric cells retain their value (#41, PR #39 review).
- Parameter-table values and vertex lists are required numbers: a blank or
  nonnumeric risk weight, correlation, liquidity horizon or DRC/RRAO weight,
  or an invalid, blank or missing `P_General` vertex-list entry, now stops
  the run with the sheet, table and row or list named, instead of becoming
  zero, being dropped or being truncated by `Val`. Every run clears the
  output sheets first and again on failure, so a failed run leaves neither
  earlier nor partial capital and notes the failure on `Results`. The
  reference applies the same numeric-text grammar to inputs (#45).
- Label pruning now checks issue/PR associations against the validated,
  policy-resolved reconciliation plan, respecting selected overlays and
  `prune: false`. Used-label protection also applies to direct CLI runs;
  offline regression fixtures cover the guard without live deletions, including
  exact matching and bounded association scans for comma-containing names (#25).
- Prototype documentation consistency: the v0.1.0 roadmap table renders as
  one table, the repository settings describe both integration branches, and
  the source, template, validation and workbook guides match the lint rules,
  CI triggers, build-time window settings and core worksheet access. Module
  comments in `M_Main`, `M_Setup` and `CORE_Util` were corrected; no code
  changed (#40).
- The structure, validation and README pages describe the VBA sync tools.
  The workbook guide says that `Dev_ExportModules` exports only from a
  workbook in the repository root or `dist\` (#77).
- Post-closeout documentation consistency: milestone names use their live
  titles with the em dash; completed work (#2, #6, #7, #8, #23, #31) is no
  longer described as pending; source archives are described as omitting the
  `export-ignore` paths; the test module and runtime test tables list
  `TEST_Numeric` and `test_numeric_contract.py`; the roadmap shows each
  v0.1.0 issue's state; the structure policy reflects the approved Excel/VBA
  prototype and the existing evidence and milestone directories; and the setup
  closeout record gains a dated addendum on the completed next decisions. No
  code changed (#59).

### Added

- [AGENTS.md](AGENTS.md) holds the editorial and assistant rules. Governance
  and contributing link to it (#168).
- `tools/check_repository.py` fails when a workflow lacks top-level
  permissions, a job lacks a timeout, a checkout keeps its credentials, a job
  outside an explicit list gets a write permission, or the label
  reconciliation is not restricted to the default branch. The current
  workflows already comply; no exception is needed (#165).
- **Issue triage** audits every open issue on issue events: assignee, one
  priority with P1 only on `bug`, a milestone, catalogue labels and a
  dependency section consistent with `blocked`. It is read-only.
  `tools/check_issues.py --check-config`, run by `tools/check_repository.py`,
  checks the issue forms, the label guide and Dependabot against the
  catalogue (#163).
- `Results` shows a note when two or three correlation scenarios give the
  same SBM total, naming the tied scenarios and the one reported. Ties still
  resolve in the order low, medium, high, now recorded as a reporting
  convention; the capital and the sample results are unchanged (#138, from
  the #29 decisions).
- The v0.2.0 — SA prototype milestone guide (`docs/Milestones/v0.2.0.md`)
  and an interim prototype evaluation written before the Excel run
  (`docs/evidence/2026-10-27-prototype-evaluation.md`). The roadmap lists
  the milestone's issues as read back on October 27, 2026; the
  documentation index and governance link both documents (#32).
- Project website with capabilities, workbook setup, methodology and validation
  links, shared K-Risk-Engines styling, and a root README website link (#100).
- Market risk and CVA method maps (#98): five diagrams in
  `docs/assets/market-risk/`, as SVG masters and PNGs, also published on the
  K-Risk-Engines site.
  - They cover the FRTB-SA and CVA frameworks, perimeter to RWA: SBM, DRC and
    RRAO; BA-CVA and SA-CVA, with the MAR50.8 carve-out and the MAR50.9
    alternative. The internal models approach and the simplified standardized
    approach are named.
  - `tools/build_market_diagrams.py` regenerates them reproducibly with
    CairoSVG 2.9.1.
  - They are method maps, not the prototype's implemented scope.

- Excel evidence record and validator (#82), adapted from the
  VBA-SACCR-Toolkit:
  - `docs/EXCEL_EVIDENCE.md` defines the manual procedure, the `host.json`
    record, the rules, the outcomes and the limits.
  - `.github/excel-evidence-policy.json` lists the required references, the
    regression item count and the ordered stages: build, compile and
    regression, and the round trip in both modes.
  - `tools/check_excel_evidence.py` checks a bundle, offline, against the
    sources and the policy in Git at the candidate commit. It also checks
    the policy against `M_Setup`.
  - Offline tests use synthetic bundles. No Excel run is recorded yet (#28).
- Excel round-trip tools (#79), from the VBA-SACCR-Toolkit:
  - `tools/Export-FRTB-VBA.ps1` exports a workbook's VBA read-only, with
    macros disabled, to an empty folder with `export-manifest.json`.
  - `tools/compare_vba_export.py` compares the export with the source of a
    commit, read from Git, in production or development mode. Re-cased
    identifiers are ignored, and document modules must hold no code. It
    reports the export's binding to the workbook as VERIFIED, FAIL or
    UNVERIFIED.
  - The workbook guide describes the round trip that #30 needs. Offline
    tests cover the comparison; no Excel export is recorded yet.

- VBA synchronization for an existing workbook (#72), modeled on the
  VBA-SACCR-Toolkit tools:
  - `tools/Sync-FRTB-VBA.ps1` replaces the standard modules of an `.xlsm`
    through Excel. A production workbook takes `src/modules`; a development
    workbook adds `tests/modules`. With `-SourceDir`, only the modules found
    in a folder are imported, identified by their `Attribute VB_Name`.
  - The script keeps a safety copy if the synchronization fails, and never
    writes cells or sheets.
  - Launchers: `Sync-FRTB-VBA.cmd` for the checkout,
    `Sync-FRTB-VBA-Downloads.cmd` for a downloads folder, and
    `Build-Workbook.cmd` for a double-click build. They print paths with
    delayed expansion, so `cmd.exe` never runs characters such as `&` that a
    path contains.
  - Offline PowerShell 7 tests cover the script's selection logic. The Excel
    run of the scripts is not yet recorded (#28).

- Release procedure and version convention in `RELEASING.md` (SemVer, `v`
  tags on main, changelog cut on request, evidence by stage) and the v0.0.1
  setup closeout record with the verification readback, including the
  owner's zero-alert security readback, the wiki entry page and the deferred
  `v0.0.1` tag, which has no complete setup-only candidate on main (#8).
- Excel/VBA SA prototype (#20/#22): SBM for five risk classes, DRC
  non-securitization and RRAO; independent Python reference, LibreOffice
  reconciliation and methodology/workbook/validation documentation. Excel
  host evidence remains pending (#28).
- Committed macro-free template and template-based builds; production and
  workbook modules under `src/modules`, with test modules under
  `tests/modules` (#30/#34). Generated `.xlsm` workbooks remain local.
- Owner-approved prototype integration branch `release/v0.1.0` alongside
  `release/v0.0.1` for setup.


- Universal Repository integrity check for ordinary PRs and pushes to main
  and release branches; one local entry point for configuration, tracked
  whitespace, local documentation targets and existing maintenance tests.
  Weekly Dependabot version updates cover Actions on both integration lines,
  the CI parser and the existing prototype test packages (#7).

- Owner-requested persistent traffic-history branch, daily/manual main-only
  collection, offline tooling checks, operational documentation, and an
  explicit analytics credential/first-run prerequisite (#23). The collector
  rejects malformed rows and existing empty history before publication.

- Bug, feature, and documentation issue forms with owner/type/priority defaults;
  one PR merge checklist; conduct and security policies with reporting routes;
  default-branch activation and manual milestone triage guidance (#6).
- Navigation glyphs and descriptive MIT/setup/development badges on the
  project entry points and new participation policies.

- Purposeful source, tests, documentation, tooling, and workbook-reservation
  READMEs; repository structure, source/generated-file and evidence policies,
  conditional VBA round-trip requirements, and local setup instructions (#5).

- Contribution walkthrough, detailed v0.0.1 milestone guide, triage and
  closeout procedures, integration verification, and a distinction between
  enforced controls and maintainer procedures (#4).

- Repository editing, ignore, and Git attribute policies, including the approved
  `src/workbook/FRTB_Template.xlsx` path (#5).

- Owner-selected integration branch `release/v0.0.1`, branch governance and
  agent instructions: all task PRs target the release branch; integration
  into main uses a separate owner-requested merge-commit PR (#4).

- Canonical 20-label catalogue, trusted synchronization, read-only drift
  detection, retained evidence, and issue-triage guidance (#1, PR #9).
- Documented repository settings baseline and pending setup dependencies
  (#2, PR #10).
- Project purpose, implementation-status table, scope and decision register,
  source/validation policy, roadmap, and current getting-started commands (#3).

### Changed

- The workbook reports version `0.2.0-prototype`, matching milestone v0.2.0 —
  SA prototype: `APP_VERSION`, the README sheet, the Log start line, the
  committed template and the README badge and status callout (#113).
- Prototype and maintenance work integrates on `release/v0.2.0`, created on
  October 27, 2026 from `0eb819b`; `release/v0.1.0` is superseded and is
  deleted under the release-branch lifecycle once milestone 2 is closed.
  Governance, contributing, installation, the README, security, scope,
  roadmap, releasing, CI, settings, the PR template and the issue forms
  follow; the Dependabot targets moved under #118 (#117).
- The documentation names the milestones as renumbered on October 27, 2026:
  v0.1.0 — Repo setup (first v0.0.1), v0.2.0 — SA prototype and v0.3.0 —
  Regulatory baseline and certification scope (#112).
  - The roadmap moves every later milestone one step: engine v0.4.0,
    workbook v0.5.0, independent validation v0.6.0, hardening v0.7.0, one
    reserved v0.8.0, release candidate v0.9.0, first certified release
    v1.0.0.
  - Releasing, scope, installation, contributing, governance, CI, settings,
    Excel evidence, the directory READMEs, the PR template and the issue
    forms follow; the methodology decisions (#29) belong to v0.3.0.
  - `release/v0.1.0` keeps its name. Dated records and earlier changelog
    entries keep the names in effect when they were written; the setup guide
    notes the renaming.
- The README names the milestones by their titles after the renumbering of
  October 27, 2026 (v0.1.0 — Repo setup, v0.2.0 — SA prototype, v0.3.0 —
  Regulatory baseline and certification scope), states that pull requests
  carry no milestone (#104), and lists the export, comparison and evidence
  tooling, the Excel evidence document, the market risk diagrams and the
  wiki (#110).
- The market risk framework diagram is a map of approaches, without project
  names (#102).
  - Market risk: the IMA for approved desks (expected shortfall,
    non-modellable risk factors, default risk charge) and the SA for the
    other desks (SBM, DRC, RRAO).
  - CVA risk: BA-CVA, reduced or full; SA-CVA, delta and vega; the MAR50.9
    alternative.
  - Below each approach, the earlier Basel 2.5 (2009) or Basel III (2011)
    rule; the total is RWA = 12.5 × K for each requirement.
  - The other four diagrams are unchanged.
- The repository is renamed from MKT-FTRB-SA, whose name transposed the
  letters of FRTB, to K-FRTB-SA, the workbook's name (#95).
  - Every link, clone command, label search and label command uses
    `danielep71/K-FRTB-SA`, dated records included; their content is
    unchanged.
  - The Excel evidence policy requires `danielep71/K-FRTB-SA` as the record's
    repository.
  - The label scripts identify themselves as `k-frtb-sa-…`.
  - GitHub redirects the old URLs.

- Maintenance audit: refresh the roadmap's dated issue inventory and remove
  unnecessary third-person wording from current guidance, forms and the PR
  template. Approval requirements and historical evidence are preserved (#93).

- Comments in the non-Markdown files now describe this repository (#87).
  - The sync script no longer claims that the tests change the input
    sheets: the regression test writes only to `Test_Expected`.
  - Module headers drop the unmaintained `UPDATED` date. `TEST_Numeric`
    follows the common header layout, and every COMPATIBILITY section names
    Excel and LibreOffice, except `M_Dev`, which needs Excel.
  - The prototype-checks workflow describes the whole suite it runs.
  - `.editorconfig`, `.gitattributes` and `.gitignore` cover only this
    repository's file types, with the same effect on every tracked file.
  - The label scripts identify themselves by this repository's name.
  - The social preview says "Excel / VBA workbook" instead of "toolkit".

- Removed `AGENTS.md` at the owner's request. Its rules duplicated the
  governance, contribution, scope and CI documents, which remain
  authoritative; CONTRIBUTING now also states the US English convention
  (#68).

- Release branches are deleted once obsolete: milestone closed, head
  integrated into main, no PR, updater or workflow targeting them and a
  repository-wide search finding no current guidance that names them as
  active. Governance defines the conditions and the
  owner's exclusion-based deletion procedure; the rulesets are unchanged and
  keep protecting active branches (#66).

- The owner deleted the setup branch `release/v0.0.1` after the v0.0.1
  closeout. Its head `b4d6c02` is on main, and the closed milestone and the
  closeout record remain the setup baseline. The release deletion ruleset
  excluded the branch only for the deletion and was then restored. Branch
  guidance, the PR template and the issue forms no longer describe the branch
  as retained (#62).

- After setup closeout, repository maintenance and all Dependabot version
  updates use `release/v0.1.0`, milestone 2, under the owner decision in #51.
  Consolidate the duplicate Actions coverage and retain both CI-parser and
  prototype pip manifests. Align current branch guidance, forms and templates;
  reject duplicate, retired, missing or incorrectly assigned updater entries
  in the integrity gate. `release/v0.0.1` stayed historical and protected,
  closed to new work, until the owner deleted it (#62).

- The roadmap records the owner-approved milestone plan from the closed
  v0.0.1 setup to v1.0.0, the first certified release, with each milestone's
  objective and exit condition; the old unversioned phases are replaced, and
  the methodology decisions (#29) move to v0.2.0 (#55).
- Settings, governance, CI, roadmap, traffic and contribution documents and
  the issue forms now record the required `Repository integrity` check,
  automatic Codex review on PR open and active traffic collection, and no
  longer describe them as pending (#2, #23, #8).
- Synchronized main setup controls into the prototype line, preserving its
  template and module layout. Shared guidance and forms describe both active
  milestones; prototype tests install the separate CI parser dependency (#31).


- Repository files consistently use the MIT license; removed external-project
  references and superseded license files at the owner's request (#5).

- GitHub About description and topics identify the FRTB-SA project and its
  setup-stage maturity (#3); the owner later updated them to describe the
  Excel/VBA prototype (#8).
- Main, release branches, and version tags have active protection rules;
  workflow action SHA pinning and the available public-repository security
  controls are enabled. The required integrity gate was added later under #2.
- Documentation uses the correct display name FRTB-SA while retaining the
  existing MKT-FTRB-SA repository identifier and URLs (#3).

No functional release has been published. Prototype LibreOffice results and
independent reconciliation are recorded; Excel execution and regulatory
certification are not claimed.
