# Repository structure and file policy

This document owns placement of durable source, generated files, and evidence.
[Scope](SCOPE.md) owns platform and functional decisions;
[installation](../INSTALLATION.md) owns the available setup commands.
The layout separates source, tests, documentation, and tools while distinguishing maintenance from the approved Excel/VBA prototype.

## Current layout

| Path | Purpose today and intended content |
| --- | --- |
| `src/` | Production source boundary |
| `src/modules/` | Prototype VBA modules: calculation engine (`CORE_*`) and workbook macros (`M_*`), issues #20 and #30 |
| `src/workbook/` | Macro-free workbook template `FRTB_Template.xlsx`, generated from `src/modules` (#30) |
| `tests/` | Verification source: independent Python reference, reconciliation and tooling tests |
| `tests/modules/` | VBA test modules imported into development builds (`TEST_*`) |
| `docs/` | Authoritative versioned scope, governance, structure, methodology, validation, and milestone documentation |
| `docs/assets/` | Documentation images and their sources: the social preview, `market-risk/` with the market risk and CVA method maps, and `screenshots/` from the recorded Excel run |
| `docs/evidence/` | Dated, sanitized evidence records bound to reviewed commits |
| `docs/Milestones/` | Detailed milestone guides |
| `docs/licenses/` | Third-party license notices |
| `tools/` | Tooling index; prototype build and template generation, VBA synchronization into an existing workbook and their `.cmd` launchers, read-only VBA export and source comparison, LibreOffice runner, VBA lint, expected-value generator, Excel run recorder and evidence validator, unattended Excel check with its acceptance-data generator and template comparison, repository, issue and external-link checks, traffic history merge and publication, and diagram generator; label tooling remains in `.github/scripts/` |
| `.github/` | Labels, forms, universal maintenance/analytics workflows, prototype checks, dependency updates and the Excel evidence policy |
| `index.html`, `assets/`, `.nojekyll` | Project website, published by GitHub Pages ([website](WEBSITE.md)) |
| Root | Entry documentation, licenses, contribution/setup guidance, and repository-control files |

Each new root has a purposeful README. Add further subdirectories with their
first real file and documented purpose, rather than creating empty engine
modules, fixtures, or placeholder test runners. Source extensions, runtime
dependencies, and build commands must follow an approved platform decision.

## Source of truth

Reviewable production source belongs in `src/`; test source and independently
specified expected results belong in `tests/`. Tools must not become a second
calculation engine. Versioned documents govern the repository; issue discussions
record decisions and evidence and should link to the resulting source.

Do not commit client portfolios, real counterparty identifiers, production
exports, credentials, private keys, or unsanitized logs. Examples and fixtures
must be synthetic and document their origin and usage rights.
An ignore rule is a tracking aid, not a data-security control.

## Generated files, workbooks, and evidence

| Material | Policy |
| --- | --- |
| Generated build/release packages | Keep local output under ignored root `build/`, `dist/`, or `release-assets/`; publish only through a separately approved release procedure |
| Test run output | Keep disposable output under ignored root `test-results/` or `test-output/` |
| Durable reviewed evidence | Add selected sanitized records under `docs/evidence/`; record commit, environment, commands, results, and limitations |
| Synthetic test inputs and expected values | Add `tests/fixtures/` and `tests/expected/` when real cases exist; record independent provenance, units, signs, tolerances, and reference versions |
| Excel workbooks, add-ins, and templates | Ignored by default; only `src/workbook/FRTB_Template.xlsx` is currently allowed by exact rooted exception |
| Local dependencies and environments | Ignored caches/installations; commit necessary manifests and locks when a real ecosystem is introduced |

The approved template exists at that path for prototype builds; this does
not approve distribution of other workbook files or a supported engine. A
template change needs a reviewable explanation of sheets, names, formulas,
data, and compatibility because Git treats Office packages as binary.
Never include confidential inputs or hidden credentials in a template.

Evidence must bind to the reviewed commit and identify what ran. A static
check is not runtime execution. Future numerical expected values must come
from a published reference or a separately derived calculation, never the
implementation under test. Missing or skipped execution is not a pass.
Do not retain every disposable log as authoritative evidence.

## Editing and Git behavior

The root controls already define:

- `.editorconfig`: CRLF for exported VBA source (`.bas`, `.cls`, `.frm`) and
  the Windows scripts, LF for everything else; 4-space indentation for VBA,
  Python and shell, 2 spaces for JavaScript, JSON, YAML, XML, HTML and CSS.
- `.gitattributes`: text normalization, VBA working-tree CRLF and diff driver,
  CRLF for the Windows scripts, LF for other text, and binary Office
  packages, form binaries (`.frx`), images, PDFs, fonts, media and archives.
- `.gitignore`: built workbooks, Office and Access locks and temporary files,
  secrets, caches, coverage and local environments, with the exact approved
  workbook-template exception.

The three files are a complete baseline (#179). Handling
depends on the file type only: they describe no tooling this repository does
not have (#87), and a path rule is added only with a layout change that needs
it. No regulatory text is stored; a byte-for-byte copy would need its own
directory, `-text` and editor rules and a registered SHA-256 in the same
change. VBA encoding is
intentionally not forced: the modules are ASCII (`tools/vba_lint.py`). Do not add working-tree encoding
or renormalize existing files without a separate justified compatibility
check. Markdown may retain intentional hard-break spaces.
The settings describe file handling; they do not select an engine platform.

<a id="conditional-excelvba-policy"></a>

## Excel/VBA source policy

Excel/VBA is approved for the prototype ([Scope](SCOPE.md)); the long-term
platform decision remains open. Record supported hosts and the build/import
procedure before claiming workbook support. The following source-first
requirements apply:

1. Commit reviewable exported `.bas`, `.cls`, and `.frm` files under
   documented source roots. Keep each required `.frx` beside its `.frm`.
   Include workbook and sheet document-module code; a binary workbook must
   never be the only record of a code change.
2. Keep the macro-free template at the approved path and construct local
   runnable workbooks from the template and reviewed exported source.
   Document component names, import order, sheet CodeNames, references,
   encoding, and production-versus-test component selection.
3. Export workbook edits back to reviewable source before committing.
   Verify a source-to-workbook-to-export round trip against the candidate
   commit, recording workbook/template digests and comparison limitations.
   Never silently overwrite workbook code or treat a source-only rebind as
   proof of a new runtime run.
4. Require actual host compilation and scenario/test execution for VBA
   changes. Evidence identifies the exact commit, Excel/OS/bitness,
   workbook digest, suites/assertions, results, failures, and skipped work.
   Independently specified numeric expectations remain a separate requirement.

The prototype has source export/build helpers and a LibreOffice runtime gate.
Desktop Excel runs on `323260b` and `1badef1` recorded the builds,
compilation, regression test and both source round trips on one host
([Excel run](evidence/2026-10-31-excel-run-323260b.md),
[Excel check](evidence/2026-11-12-excel-check-1badef1.md)); the second also recorded the
acceptance checks and the synchronization. The remaining Excel review is
#206; that gap must stay visible.

**Prototype implementation (issues #20, #30).** Exported modules live in
`src/modules` and `tests/modules`; the committed template at the approved
path is generated from them by `tools/build_template.py` (LibreOffice) or
`tools/Build-Workbook.ps1 -Template` (Excel), and `tests/test_template.py`
checks that it matches a fresh build and carries no VBA project, controls or
personal properties. Excel builds start from the template and
`M_Setup.FinishWorkbook` adds the buttons and window settings. The automated
checks run LibreOffice Calc in VBA-compatibility mode, which is not Excel host
evidence; desktop Excel runs, with both round trips and their digests, are
recorded on `323260b` and `1badef1` (#28, #30), and a change after `1badef1`
needs a new run. See [validation](validation.md).
