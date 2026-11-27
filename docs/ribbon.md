# 🎀 Ribbon tab

The workbook can carry an **FRTB-ASA** ribbon tab with the commands of the
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

## Icons

Each control has an icon of its own, in one style: a navy rounded tile
(`#1F3864`, the color of the workbook's headers) with a white glyph, which reads
on a light and on a dark ribbon. The `image` attribute of a control names the
icon, and `tests/test_ribbon.py` fails when the part, this table and the files
disagree.

| Id | Image | Glyph | Files in `src/ribbon/icons` |
| --- | --- | --- | --- |
| `btnRun` | `run` | a play triangle | `run.svg`, `run.png` |
| `btnClear` | `clear` | an eraser over a line | `clear.svg`, `clear.png` |
| `btnTest` | `test` | a checklist with three ticks | `test.svg`, `test.png` |
| `tglNav` | `nav` | a window with its column of links | `nav.svg`, `nav.png` |

The `.svg` is the source and the `.png`, 64 by 64 pixels with a transparent
background, is what the workbook carries; Excel scales it to the size of the
screen. To change an icon, edit the `.svg` and render it again at 64 px.

## States

- **Busy:** `M_Main` marks a run as busy when it starts and idle when it ends,
  also when it fails, and asks the ribbon to read its states again, so the
  three commands are disabled for the duration of a run.
- **Navigation bar:** the toggle shows or hides every `Nav*` shape of every
  sheet and column A with them, except that the two triangles in B1 take
  opposite states: `NavToggleHide` (points left) is shown with the bar and
  `NavToggleShow` (points right) while it is hidden. Both run
  `M_Main.ToggleNavigationBar`, which asks for the state the bar is not in and
  refreshes the ribbon, so its toggle follows a click on a triangle. Column A: column A is the black column the bar sits in,
  and hidden it leaves no empty strip. The sheet buttons are rounded
  rectangles that move with their cells, so they follow column A without
  being moved. Its state is read from the marker of
  `Home`, so a workbook saved with the bar hidden opens with the toggle
  released. The write goes through a write window, because the shapes and the
  column are locked on a protected sheet (#349). The bar's shapes take their
  visibility in one call for each sheet, through a range of their names
  (`NavigationShapeNames`), which keeps the toggle quick; a sheet with a missing
  name goes shape by shape. The write window also holds
  screen updating off, so the sheets do not flicker, and the toggle gives the
  `Saved` flag of the workbook back, so that it is not a change that asks to be
  saved. The acceptance check
  `NAVIGATION` toggles the bar and reads column A, the shapes and the first
  button of all 23 sheets after each call.
- **Production build:** `-Production` leaves out `TEST_Regression`, so **Run
  regression test** has no macro there, as the sheet button has none.

## What the part adds to the file

`tools\Add-RibbonTab.ps1` opens the saved `.xlsm` as a zip package and:
- writes `customUI/customUI14.xml`;
- writes each icon the part names to `customUI/images/` and the relationships
  of the part that point to them to `customUI/_rels/customUI14.xml.rels`, and
  makes sure `[Content_Types].xml` has the default content type of `png`;
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

The full profile of the unattended Excel check (`tools\Check-FRTB-Excel.ps1`)
runs most of this itself, in a stage after the builds:

| Result | What it shows |
| --- | --- |
| `BUILD_RIBBON` | `Build-Workbook.ps1 -Ribbon -RunTests` exits 0, and the file carries the part with the control ids of `src\ribbon\customUI14.xml`, its four icons, the relationship and the VBA project |
| `RIBBON_CALLBACKS` | In Excel, `Ribbon_ToggleNav` hides and shows the navigation bar and `Ribbon_Clear` changes `Results`; the callbacks ignore their control, so the check passes the workbook in its place |
| `RIBBON_KEPT_BY_EXCEL` | After Excel opened and saved the workbook, the file still carries the part, its icons, its relationship and the VBA project |
| `RIBBON_ROUND_TRIP` | The exported modules of the workbook with the tab equal the source |
| `RIBBON_SYNC` | The modules synchronize into a copy, and the part is still there |
| `RIBBON_TAB_SHOWN` | Always `NOT_RUN`: the tab on the screen, the enabled state while a run is busy and a click on each command need a person |

What stays with a person: build with `-Ribbon`, open the file, look at the
**FRTB-ASA** tab, click each command once and watch the three commands go
disabled during a run.
