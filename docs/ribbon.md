# 🎀 Ribbon tab

The workbook can carry an **FRTB-SA** ribbon tab with the commands of the
sheet buttons (#353). The tab is optional: the buttons stay on the sheets, so a
workbook still works where a policy removes custom UI. Excel's object model
saves no custom UI part, so `tools\Build-Workbook.ps1 -Ribbon` adds
`src\ribbon\customUI14.xml` to the saved file with `tools\Add-RibbonTab.ps1`;
the committed template and a build without `-Ribbon` have no tab.

## The contract

Every control, its callback and the procedure the callback calls. A callback
holds no logic: it calls one procedure, and `tests/test_ribbon.py` fails when
the custom UI part, this table and `M_Main` disagree.

| Control | Id | Type | Callback | Calls | Enabled | Screen tip |
| --- | --- | --- | --- | --- | --- | --- |
| Run FRTB-SA | `btnRun` | button | `Ribbon_Run` | `RunFRTB` | while idle (`Ribbon_IsIdle`) | Run FRTB-SA |
| Clear outputs | `btnClear` | button | `Ribbon_Clear` | `ClearOutputs` | while idle (`Ribbon_IsIdle`) | Clear outputs |
| Run regression test | `btnTest` | button | `Ribbon_Test` | `Test_RunRegression` | while idle (`Ribbon_IsIdle`) | Run regression test |
| Navigation bar | `tglNav` | toggle button | `Ribbon_ToggleNav` | `SetNavigationVisible` | always; pressed while the bar is shown (`Ribbon_NavShown`) | Navigation bar |

Two more callbacks serve the tab: `Ribbon_OnLoad` keeps the ribbon object, and
`Ribbon_IsIdle` and `Ribbon_NavShown` return a state without changing anything.

## States

- **Busy:** `M_Main` marks a run as busy when it starts and idle when it ends,
  also when it fails, and asks the ribbon to read its states again, so the
  three commands are disabled for the duration of a run.
- **Navigation bar:** the toggle shows or hides every `Nav*` shape of every
  sheet. Its state is read from the marker of `Home`, so a workbook saved with
  the bar hidden opens with the toggle released. The write goes through a write
  window, because the shapes are locked on a protected sheet (#349).
- **Production build:** `-Production` leaves out `TEST_Regression`, so **Run
  regression test** has no macro there, as the sheet button has none.

## What the part adds to the file

`tools\Add-RibbonTab.ps1` opens the saved `.xlsm` as a zip package and:
- writes `customUI/customUI14.xml`;
- adds one relationship of type
  `http://schemas.microsoft.com/office/2007/relationships/ui/extensibility`
  to `_rels/.rels`, unless one exists;
- makes sure `[Content_Types].xml` has the default content type of `xml`.

The VBA project and every other part stay as they are, and a second run leaves
one part and one relationship. Close the workbook before you run it.

## Feasibility

What a test shows, here and in CI: the part is valid XML in the 2009 custom UI
namespace, the contract agrees with it and with `M_Main`, and the script adds
the part and the relationship to a synthetic package and is idempotent. What
only desktop Excel shows, and the issue stays open for: that Excel opens the
saved workbook without a repair prompt, shows the tab, runs each command, keeps
the macros, and that the synchronization, the export and the round trip pass
with the tab present. Build with `-Ribbon`, open the file and check each
control; the Excel evidence records the result.
