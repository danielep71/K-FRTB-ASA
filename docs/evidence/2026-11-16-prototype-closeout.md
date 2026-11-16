# ✅ v0.2.0 prototype closeout — November 16, 2026

The closeout record of issue #32 and the
[v0.2.0 — ASA prototype](https://github.com/danielep71/K-FRTB-ASA/milestone/2)
milestone. It replaces the
[interim evaluation](2026-10-27-prototype-evaluation.md) of October 27, 2026,
where they differ. I read the milestone and the branches back on
November 16, 2026, after integration #336 and syncs #337 and #338.

| | |
| --- | --- |
| Main | `be7386760664c5bef0a5a17be65a8cc3b64fd954` (#336) |
| `release/v0.2.0` | `1a9dd8f84e28ce94544032a64dd1f816658dc5c8` (#337), same tree as main |
| `release/v0.3.0` | `8fed135bb67a37e88a9601fad62b328e5979bdd7` (#338), same tree as main |
| Excel candidate | `308cbc6b144943ce17d5e369a4a7c35d5ea70290`, [check of November 15](2026-11-15-excel-check-308cbc6.md): 24 of 24 |
| Template | `src/workbook/FRTB_Template.xlsx`, SHA-256 `ea430d540c13ec667b37d3fc6d8b5d8fd4eb98699ce707b12432f7c87b1dc3d8`, the same blob on `308cbc6` and on main |

The prototype is evaluated: it builds, compiles and runs in desktop Excel on
one host, reconciles with the independent reference and round-trips its
sources, and its limits are written down below. Nothing here is regulatory
validation, and no tag, release or workbook is published.

## 🧮 Coverage reached

| Component | Implemented | Not implemented |
| --- | --- | --- |
| SBM | Delta, vega and curvature for GIRR, CSR non-securitization, equity, commodity and FX; low, medium and high correlation scenarios; SBM is the largest scenario total | CSR securitization (non-CTP and CTP), whose rows are rejected; sensitivities and CVRs are inputs, not priced |
| DRC | Non-securitization: gross JTD by seniority, maturity scaling with the three-month floor, netting by obligor and seniority, bucket charge with the hedge benefit ratio | Securitization and CTP DRC |
| RRAO | Exotic and other residual risks on gross notional | Instrument classification: eligibility is an input |
| Parameters | Editable `P_*` sheets with MAR paragraph references, checked against the BIS text | Any jurisdiction profile |
| Workbook | Input, parameter, output and log sheets; **Run FRTB-SA**, **Clear outputs** and **Run regression test**; an in-workbook regression test of 52 items; the reviewed formatting of every sheet (#206) | — |

On the synthetic sample portfolio, with the default parameters and every
discretion off, the total capital requirement is 3,877,792.81 EUR: SBM
3,163,703.42 (low scenario), DRC 636,589.40 and RRAO 77,500.00. Excel
reproduced the total on `308cbc6`.

## 🔎 Evidence

| Kind | Result | Commit |
| --- | --- | --- |
| Runtime, Excel | Build, **Debug → Compile** and round trip in both modes, regression 52 of 52, twelve acceptance checks, 44 sheet pictures, template and source builds, the Trust Center error and six synchronization checks: 24 of 24 ([record](2026-11-15-excel-check-308cbc6.md)) | `308cbc6` |
| Earlier Excel runs | The first run, with build, compile, regression and round trip ([record](2026-10-31-excel-run-323260b.md)); the unattended check, 21 of 21 ([record](2026-11-12-excel-check-1badef1.md)) | `323260b`, `1badef1` |
| Runtime, LibreOffice | Hosted [prototype checks](https://github.com/danielep71/K-FRTB-ASA/actions/runs/43935089359) on the merge ref of #336, whose tree is main's; and on #328 ([run](https://github.com/danielep71/K-FRTB-ASA/actions/runs/43739258090)), whose tree is `308cbc6` | main, `308cbc6` |
| Fresh checkout | On a clean worktree of `be73867`: `tools/check_repository.py` with every quality tool passes, and `pytest` with LibreOffice and PowerShell 7 gives 378 passed and 4 skipped (Ruff and mypy, which the gate runs), with 115 subtests | `be73867` |
| Hosted checks on main | [Repository integrity](https://github.com/danielep71/K-FRTB-ASA/actions/runs/43935659517), [CodeQL](https://github.com/danielep71/K-FRTB-ASA/actions/runs/43935659554), [traffic tooling](https://github.com/danielep71/K-FRTB-ASA/actions/runs/43935659590) and [label sync](https://github.com/danielep71/K-FRTB-ASA/actions/runs/43935659586) passed | `be73867` |
| Independent numerical comparison | The Python reference (`tests/frtb_reference.py`) reproduces every output of the VBA run in LibreOffice, on the sample and on a variant with every discretion on and invalid rows; hand-calculated tests cover the formulas | main |
| Methodology | I decided the six interpretations on October 27, 2026 (#29), and they are implemented (#134, #136, #138, #140, #142, #144) | main |

**The Excel evidence holds for main.** No file under `src/`, `tests/modules/`
or the PowerShell and command scripts of `tools/` changed between `308cbc6`
and `be73867`, and the template is the same blob. The 51 files that changed
are CI, tooling, tests and documentation.

LibreOffice runs the VBA in a compatibility mode; it is not Excel. The
reference and the VBA share one reading of MAR21–MAR23, so their agreement
does not prove that reading is right.

## 🔧 Source-first build: effort and risks

The workbook is never committed. It is built from the macro-free template
and the exported modules, by `tools/Build-Workbook.ps1` in Excel or
`tools/lo_runner.py` in LibreOffice. Around it sit the sync into an existing
workbook, the read-only export with its source comparison, the evidence
validator, the recorder and the unattended check.

| Risk in the interim evaluation | Outcome |
| --- | --- |
| LibreOffice accepts VBA that Excel rejects, or the reverse | Real: the build of `3c20005` stopped on a compile error in `M_Setup` that LibreOffice accepted (#327). Fixed in #328, and `tools/vba_lint.py` now refuses module-level code after the first procedure. Excel compiled both builds on `308cbc6` |
| The template was produced by LibreOffice | Excel builds an equivalent template. The differences come from the host that saves the file: the margin column width, √2 to 15 digits, and Booleans stored as numbers ([record](2026-11-12-excel-check-1badef1.md)) |
| Locale-dependent number parsing | Checked on a comma-decimal host: numeric text is rejected unless it uses a period, and numbers display in the host's format (`NUMERIC_TEXT`, `DISPLAYED_VALUES`) |
| Workbook and source drift apart | The round trip passed in both modes with a verified binding: 13 and 9 modules, 0 differences |
| Excel settings block the build | The build exits 1 with a message naming the Trust Center setting (`TRUST_CENTER_ERROR`) |

The effort was real. The tools above, and three Excel sessions with their
fixes, exist to keep one workbook reproducible from reviewed sources. The
v0.3.0 platform decision should weigh that effort against the benefit of
staying in Excel.

## 🗂️ Committed template

The template is macro-free: no VBA project, form controls or personal
properties. `tools/build_template.py` regenerates it, and `--check` requires
it to equal a fresh build. It carries the shared style (#76), the reviewed
sheet formatting (#206: Aptos throughout, vertical centering, navy output
headers, input formats), the MAR citations and version `0.2.0-prototype`.
The build adds buttons and tab colors.

## 🔌 Worksheet access in the core modules

Calculation is still not separated from worksheet access. `CORE_Util` wraps
sheet access; `CORE_SBM`, `CORE_DRC` and `CORE_RRAO` read their input
sheets through it, and `CORE_SBM` and `CORE_DRC` write and format their
output sheets directly. Input columns are read by position. Separating
calculation from worksheet access belongs to v0.4.0 in the
[roadmap](../ROADMAP.md).

## ⚠️ Known limits

- **Excel scope.** One host: Excel 16.0, 64-bit, comma decimal separator.
  Other Excel versions, 32-bit Excel, other regional settings and printing
  are untested.
- **Interpretations** (#29): my readings of MAR21–MAR23, not independently
  reviewed; v0.3.0 reviews them for the certified scope. See the
  [methodology](../methodology.md#interpretations-and-simplifications-to-review).
- **Input contract.** Input columns are read by position, without a header
  check. Flags accept only TRUE, FALSE or blank (#144).
- **Scope.** The exclusions under coverage above; no jurisdiction profile.
- **Scale.** Only the synthetic sample and small test cases have run; the
  engine has not been timed on large inputs.

## 🧾 Issue reconciliation

Readback of milestone 2 on November 16, 2026, through the GitHub API:

- **96 issues, no pull requests** (#104). Every issue has `danielep71` as
  assignee and exactly one priority. P1 is used only on #41, a confirmed
  wrong result on valid input (#162).
- **95 closed:** 94 completed. #222 was closed as not planned on
  November 4, 2026, and its traffic fix was done in #332.
- **Open:** #32, this closeout. It closes with this record.

| Workstream | Issues |
| --- | --- |
| Prototype engine and input validation | #20, #41, #45, #48, #85, #245 |
| Template, layout and workbook formatting | #30, #76, #113, #203, #206, #327 |
| Excel tooling and evidence | #28, #72, #79, #82, #199, #201, #247, #274, #276, #279, #283, #290, #295, #297, #302, #304, #306, #316, #324 |
| Branches, milestones and governance | #31, #51, #55, #62, #66, #68, #104, #117, #118, #132, #151, #168, #187, #321 |
| CI, labels and repository controls | #124, #149, #158, #160, #162, #163, #165, #167, #177, #179, #191, #217, #218, #219, #220, #221, #223, #233, #234, #235, #246, #263, #320, #330, #332, #334 |
| Traffic history | #181, #185, #222 |
| Documentation, website and rename | #40, #59, #77, #87, #93, #95, #98, #100, #102, #110, #112, #115, #126, #174, #248, #249, #250, #261, #268, #288, #311 |
| Closeout | #32 |

Integrations into main since the interim evaluation, each a merge commit of
`release/v0.2.0` or `release/v0.3.0` on my request: #130 (October 27); #146
and #148 (October 30); #154, #156, #172, #189 and #197 (October 31); #214
(November 3); #230 (November 4); #239 (November 5); #242 (November 9); #257,
#265, #271, #278 and #285 (November 11); #292, #299 and #308 (November 12);
#313 (November 14); and #336 (November 16). Earlier integrations are listed
in the [interim evaluation](2026-10-27-prototype-evaluation.md#-issue-reconciliation).

## 🧭 Adopted and deferred scope

| Item | Decision |
| --- | --- |
| Excel/VBA for the prototype, committed template plus exported modules | Adopted (#20, #30) |
| Independent Python reference and LibreOffice reconciliation as the automated gate | Adopted |
| Evidence format, offline validator and the unattended Excel check | Adopted (#82, #274) |
| Project website on GitHub Pages | Adopted (#100) |
| Methodology interpretations | Decided and implemented (#29); reviewed for the certified scope in v0.3.0 |
| CSR securitization, CTP, securitization DRC | Excluded |
| Jurisdiction profile | Deferred to v0.3.0 |
| Long-term platform | Deferred: I decide it in v0.3.0, informed by this record |
| Large-portfolio performance | Deferred to v0.7.0 |
| Tag, GitHub release, workbook distribution | Not done. Under [Releasing](../../RELEASING.md), a tag needs my explicit request and a main commit with a dated changelog section |

## ➡️ After closeout

1. Close milestone 2. Closing it merges nothing, creates no tag and selects
   no platform.
2. A tag of v0.2.0, if I want one, needs a separate request: a task PR dates
   the changelog section and reaches main through an integration, and the
   candidate is validated under [Releasing](../../RELEASING.md).
3. Once the milestone is closed and this record is on main,
   `release/v0.1.0` and then `release/v0.2.0` become obsolete and are
   deleted under the
   [release-branch lifecycle](../GOVERNANCE.md#release-branch-lifecycle).
   Work continues on `release/v0.3.0`.
