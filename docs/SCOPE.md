# Scope and project decisions

## Current stage

Repository setup (v0.1.0 — Repo setup) is closed. The active workstream is the
SA prototype and repository maintenance on release/v0.2.0, milestone 2,
v0.2.0 — SA prototype (#117).
Main and release/v0.2.0 contain the setup controls, Excel/VBA prototype,
committed macro-free template and independent Python reference.
The setup branch release/v0.0.1 was deleted after the closeout; its history
is on main. No functional release or regulatory certification has been
published. The prototype has run in LibreOffice Calc VBA-compatibility mode,
and one desktop Excel run is recorded on `323260b`, on one host
([Excel run](evidence/2026-10-31-excel-run-323260b.md)): both builds, both
compilations, the regression test and both round trips. The remaining manual
Excel checks are in #28, #72, #85 and #206.

The intended outcome is transparent, testable market risk calculation for
practitioners, developers, validators, researchers and educators.

## Prototype coverage and boundaries

| Workstream | Implemented prototype scope | Remaining boundary |
| --- | --- | --- |
| SBM | GIRR, CSR non-securitization, equity, commodity and FX; delta, vega, curvature; three correlation scenarios | CSR securitization and CTP excluded; sensitivities/CVRs are inputs, not priced here |
| DRC | Non-securitization JTD, maturity scaling, seniority netting and bucket aggregation | Securitization DRC out of scope: `In_DRC` cannot identify one, so it must not be entered; the DRC readings of #29 are decided and implemented (#134, #142) |
| RRAO | Gross notional with exotic/other residual-risk weights | Eligibility remains an input contract, not an instrument classifier |
| Validation | Independently derived Python reference, hand-calculated cases, LibreOffice reconciliation and one desktop Excel run on `323260b` | Excel checks beyond that run (#28, #72, #85, #206), an independent review of the #29 readings (v0.3.0) and production-scale performance are not established |

Coverage and data contracts are detailed in [methodology](methodology.md),
[workbook guide](workbook-guide.md) and [validation](validation.md). Each
increment needs observable input/output contracts, rejection behavior and
independently specified expected values; prototype coverage is not complete
regulatory applicability.

## Platform and environment

**Excel/VBA is approved for the prototype; the long-term platform decision
remains open.** The approved layout places a committed macro-free template under
src/workbook and production/workbook modules under src/modules (#30/#34).
This is no approval of production support, an Excel-certified host matrix or
public workbook distribution. Windows desktop Excel/PowerShell is the target
build path; one run of it is recorded on `323260b`, on one host. LibreOffice
execution is a distinct compatibility runtime.

| Decision or evidence | Current record and remaining work |
| --- | --- |
| Long-term engine language/runtime and supported platforms | Explicit architecture decision informed by prototype evaluation (#32) |
| Basel baseline and initial jurisdiction profile | Basel MAR21-MAR23 prototype baseline in methodology; no jurisdiction profile selected |
| Initial products/risk classes | Prototype coverage above; CSR securitization rows rejected, securitization DRC out of scope and not identifiable on `In_DRC` |
| Inputs, units, signs, identifiers and outputs | Workbook guide and methodology; the #29 interpretations are decided and implemented, and v0.3.0 reviews them for the certified scope |
| Expected-value provenance/tolerances | Independent Python reference, hand-calculated cases and validation record |
| Host execution/export and release packaging | Excel run recorded on `323260b` (#28; round trip #30 closed), remaining checks in #28, #72, #85 and #206, setup release procedure #8 and prototype closeout #32 |

Maintenance tooling uses Node built-ins and Python with a pinned CI parser;
the prototype test environment has its own packages. These are documented in
[Repository checks](CI.md) and do not select the long-term engine platform.
[Governance](GOVERNANCE.md) assigns current prototype and maintenance work
to release/v0.2.0, milestone 2 (v0.2.0); it superseded release/v0.1.0, and the setup
branch release/v0.0.1 was deleted (#62). Source and evidence placement follow the
[structure policy](REPOSITORY_STRUCTURE.md).

## Source and validation policy

Start from primary sources and distinguish their roles:

- The [BIS Basel Framework market risk standard](https://www.bis.org/committees/bcbs/basel-framework/standard/mar)
  provides the Basel reference structure. MAR20 covers the standardized
  approach structure, MAR21 the sensitivities-based method, MAR22 default
  risk, and MAR23 the residual risk add-on.
- [EUR-Lex's January 1, 2026 consolidated CRR text](https://eur-lex.europa.eu/eli/reg/2013/575/2026-01-01/eng)
  is an initial EU research reference, not a selected implementation baseline
  or a statement that all provisions apply on the same date. Review the
  underlying Official Journal acts, amendments, delegated/implementing acts,
  and applicable transition dates when defining an EU profile.

These reference entry points were checked on October 17, 2026. Their presence
does not establish the project's regulatory applicability or completeness.

Before implementing a rule, record its issuer, title, direct URL, chapter or
article, publication/version date, effective/application date, retrieval date,
profile, and interpretation assumptions. Retain amendment history and map
implemented behavior to specific source provisions. Keep Basel and
jurisdiction profiles explicit; do not silently mix parameters or assumptions
from different versions or regimes. Resolve material interpretation questions
before marking the corresponding calculation complete.

Expected results must come from published examples or a separately derived
calculation with documented independence. Never obtain the expected value by
running the implementation under test. Define intermediate checks, units,
tolerances, edge cases, rejected inputs, and provenance before asserting a
benchmark passes. Tie execution evidence to the exact reviewed source.

Distinguish static checks, actual runtime execution, independent numerical
comparison, and methodology review. A passing repository workflow does not
establish correct capital calculations. Future releases must state which of
these activities were performed and which remain incomplete.

## Data and use boundaries

Use synthetic datasets and examples only. Do not commit client portfolios,
real counterparty identifiers, confidential positions, credentials, or
production exports. Future fixtures must document their origin and usage
rights. A regulatory example must be cited and its redistribution conditions
respected.

The prototype is an evaluation tool and supports no regulatory
reporting or production capital determination. Internal-model calculations,
pricing engines, live market-data ingestion and regulatory reporting templates
are outside the approved prototype scope. The [project website](WEBSITE.md)
describes the project; it does not run the engine. Additional
capabilities require their own explicit scope and validation work.
