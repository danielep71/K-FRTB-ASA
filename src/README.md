# 🧩 Production source

This directory is the production-source boundary. It holds the exported VBA
of the ASA prototype, milestone v0.2.0 (issue #20), and the macro-free [workbook template](workbook/README.md).
The prototype uses Excel/VBA to evaluate the approach; it is not the platform
decision, which stays open in [Scope](../docs/SCOPE.md).

[Repository structure](../docs/REPOSITORY_STRUCTURE.md) owns placement and the
source-of-truth policy; the [workbook guide](../docs/workbook-guide.md) owns
the module layout, build and sheet conventions.

| Path | Content |
| --- | --- |
| `modules/CORE_Config.bas` | Product name, sheet names, layout, palette and the risk class, component and scenario codes |
| `modules/CORE_Util.bas` | Sheet access, tagged parameter tables, conversions, table formatting and the run log |
| `modules/CORE_Params.bas` | Loading of the Control settings and `P_*` parameters; risk weight and correlation accessors |
| `modules/CORE_SBM.bas` | Sensitivities-based method (delta, vega, curvature) under the three correlation scenarios |
| `modules/CORE_DRC.bas` | Default risk charge for non-securitizations |
| `modules/CORE_RRAO.bas` | Residual risk add-on |
| `modules/M_Main.bas` | Sheet-button macros and automation entry points; Results and Out_Summary reports |
| `modules/M_Setup.bas` | Workbook builder (sheets, palette, parameters, sample portfolio, expected results) and `FinishWorkbook` |
| `modules/M_Dev.bas` | Export of the modules back to the repository |
| `workbook/FRTB_Template.xlsx` | Macro-free template generated from the modules; see [workbook/README.md](workbook/README.md) |

`CORE_` modules declare `Option Private Module`. Known deviation: `CORE_Params`,
`CORE_SBM`, `CORE_DRC` and `CORE_RRAO` read the input and parameter sheets,
`CORE_SBM` and `CORE_DRC` write the output sheets and `CORE_Util` writes the
`Log`, so core is not yet free of worksheet access.

Keep test code under `tests/` and maintenance logic under `tools/` or the
existing `.github/scripts/`. Workbook binaries must never be the only source
record.
