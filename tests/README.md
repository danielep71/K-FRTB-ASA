# 🧪 Verification source and data

This directory holds the verification source of the SA prototype, milestone
v0.2.0 (issue #20). Expected results are derived by an independent Python reference,
never copied from the VBA under test.

| Path | Content |
| --- | --- |
| `modules/TEST_Regression.bas` | VBA regression test imported into development builds; compares the last run with `Test_Expected` |
| `modules/TEST_Numeric.bas` | VBA numeric validation and conversion contract test imported into development builds |
| `frtb_reference.py` | Independent Python reference implementation of the prototype's scope |
| `regulatory_defaults.py` | Regulatory parameters typed independently of `M_Setup` |
| `test_reference.py` | Hand-calculated checks of the reference |
| `test_vba_engine.py` | Headless LibreOffice runs of the VBA reconciled with the reference |
| `test_tooling.py` | VBA lint and freshness of the embedded expected values |
| `test_template.py` | The committed template matches a fresh build and a build from it runs and reconciles |
| `test_compare_vba_export.py` | Offline tests of `tools/compare_vba_export.py` on synthetic exports of the source at `HEAD`, and static checks of `tools/Export-FRTB-VBA.ps1` |
| `test_excel_evidence.py` | Offline tests of `tools/check_excel_evidence.py` with synthetic bundles in a temporary Git repository, and the policy check |
| `test_record_excel_run.py` | Offline tests of `tools/record_excel_run.py` with synthetic bundles: `host.json` follows the logs; never Excel evidence |
| `test_sync_script.py` | Offline PowerShell 7 tests of `tools/Sync-FRTB-VBA.ps1` (module selection, name checks, CRLF conversion, module sets); skipped without `pwsh` |
| `test_numeric_contract.py` | Headless LibreOffice runtime regressions for the numeric-text and parameter-table contracts and the empty state of the output sheets (#41, #45, #48, #85) |
| `conftest.py` | Shared LibreOffice fixtures |
| `requirements.in`, `requirements.txt` | Python test dependencies: the direct pins, and the file compiled from them with every package pinned by hash ([Repository checks](../docs/CI.md#reproducible-runs-220)) |

The sample portfolio and the test cases are synthetic, written for this
project. Each hand-calculated case states its inputs and the MAR formula it
works through; tolerances are in the tests. Do not put disposable run output
here; use ignored root `test-results/`.

`tooling/` contains offline repository-integrity and traffic-collector tests.
The universal entry point `python tools/check_repository.py` runs them with
the existing label commands; [Repository checks](../docs/CI.md) defines its
maintenance-only coverage. The prototype pytest environment installs both
`.github/requirements.txt` and `tests/requirements.txt` so these tests can also
be collected without relying on incidental OS packages.

The label scripts listed in [Tooling](../tools/README.md) keep their own
self-tests. [Scope](../docs/SCOPE.md) owns validation principles,
[Repository structure](../docs/REPOSITORY_STRUCTURE.md) owns placement and
durable evidence, and [Validation](../docs/validation.md) records what the
prototype checks do and do not establish.
