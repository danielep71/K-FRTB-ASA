# 🧾 Workbook template

`FRTB_Template.xlsx` is the macro-free template of the K-FRTB-ASA prototype,
at the approved path. It is the workbook that `M_Setup.BuildWorkbook`
produces from [`src/modules`](../modules), saved without a VBA project,
sheet buttons or personal document properties. Other workbook names, paths,
and macro-enabled variants remain ignored.

| Content | Detail |
| --- | --- |
| Sheets | README, Control, Results, In_Delta, In_Vega, In_Curvature, In_DRC, In_RRAO, Out_Summary, Out_RiskClass, Out_Buckets, Out_DRC, Log, P_General, P_GIRR, P_CSR, P_EQ, P_COM, P_FX, P_DRC, P_RRAO, Test_Expected |
| Formulas and names | No calculation formulas or engine names: values are written by `M_Setup`; built-in print-area/title names store printing settings |
| Data | Default Basel MAR21-MAR23 parameters, the synthetic sample portfolio and the expected regression results; no real portfolio data |
| Presentation | Hidden gridlines, role-colored tabs, frozen headers and identifiers, readable titles, bordered input cells, output headers, landscape printing and Aptos footers are saved in the XLSX |
| Not included | VBA, calculation results and macro buttons; `M_Setup.FinishWorkbook` adds the buttons and reapplies view settings at build time |

Never edit the template by hand. Change `src/modules` and regenerate it with
`python3 tools/build_template.py` (LibreOffice, used for the committed file)
or compare with `tools\Build-Workbook.ps1 -Template` (Excel).
The headless builder also runs `tools/template_presentation.py` to preserve
view/print settings and README guide borders that LibreOffice's VBA
compatibility layer omits. That step changes no cell values or formulas;
tab colors come from the existing VBA palette, and the guide borders
restore the dashed interiors and solid outlines specified by `WriteGuideRows`.
The VBA builder grows non-output print areas with the formatted table extents,
so Excel source builds use the same bounded areas without Python normalization.
`tests/test_template.py` fails when the committed file differs from a fresh
build. See the [workbook guide](../../docs/workbook-guide.md) and
[repository structure](../../docs/REPOSITORY_STRUCTURE.md).
