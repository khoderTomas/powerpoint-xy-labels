# PowerPoint XY chart label layout

A PowerPoint add-in (`RozmistiPopisky.ppam`, pure VBA) that lays out the data labels of native XY (scatter)
charts: labels taken from cells, with leader lines, as in research decks that are regenerated several times
a month and whose labels always end up piled on top of each other.

- Isolated points get their label right next to them.
- The rest go into two aligned columns, left and right of the cluster, with no overlaps and no crossing
  leader lines. A label stays in the same half of the chart as its point.
- Optional **shorter lines** placement: starts from the columns, then pulls labels closer to their points.
- When the columns do not fit (points spread over the whole chart), a fallback places as many labels as
  possible next to their points. Labels that still do not fit stay where they were and the report lists them.
- When the placement does not come out without defects, the chart is pre-adjusted first: a slightly smaller
  label font (at most 2 pt) and wider axis ranges (at most 1.3×, only axes without tick labels). The mildest
  step that gives a placement without defects wins; data points are never changed.
- The font size of the labels in a chart is unified to the most common one.
- The result goes to a new presentation next to the original (default), or the open presentation is edited
  in place and left unsaved (a backup goes to the TEMP folder first).

The result is still a native chart: labels can be dragged with the mouse and stay linked to the data.

**Download:** [RozmistiPopisky.ppam](https://github.com/khoderTomas/powerpoint-xy-labels/releases/latest/download/RozmistiPopisky.ppam)
(latest release). Installation and use: [USER_GUIDE.md](USER_GUIDE.md).

The user interface of the add-in (ribbon button, dialog, report) is in Czech.

## What it handles

- XY scatter charts with data labels (typically "Value From Cells" labels with leader lines).
- Linear axes with fixed or automatic bounds. Reversed and logarithmic axes are skipped.
- Charts inside groups (unrotated), other shapes on the slide as obstacles, vertical reference lines
  (a series without markers whose points share one x value) and a horizontal quadrant divider (a connector).
- A slide with `#popisky-ne` in its notes is skipped.
- Not handled yet: bubble charts.

## How it works

All the work happens in VBA inside PowerPoint, no external program.

| Module | Role |
|---|---|
| `addin/Popisky.bas` | Ribbon button, processing of the presentation, slide selection, review report (HTML) |
| `addin/Graf.bas` | Reading and writing through the object model, geometry, defect counting |
| `addin/Rozlozeni.bas` | Placement into aligned columns |
| `addin/Blizko.bas` | "Shorter lines" placement and the fallback placement |
| `addin/Predupravy.bas` | Pre-adjustment: smaller font and wider axis ranges, remembered in shape tags |
| `addin/Pisma.bas` | Text measurement (GDI) |
| `addin/Nastaveni.vba` | Code of the dialog (an empty UserForm that creates its controls at run time) |
| `addin/customUI14.xml` | Ribbon button on the Home tab |

- **Geometry from the object model.** Chart coordinates = object model coordinates + `ChartArea.Left/Top`.
  A point's position = `PlotArea.Inside*` and the linear axis (`MinimumScale/MaximumScale`; an automatic bound
  returns the value PowerPoint draws with).
- **Text measurement through GDI**: the font at 1000 px per em, width from `GetTextExtentPoint32W`, line spacing
  from the font tables (`GetFontData`). Label size = text width + 6 pt, line spacing × size + 3 pt.
- **Columns**: a direct label where there is room next to the point, the rest in columns. Every left/right split
  is tried and the cheapest wins (leader line length, labels across reference lines, does not fit). Order in a
  column without crossings (2-opt), positions as close to the points as possible (least squares), going around
  boxes and the legend.
- **Shorter lines**: threshold accepting from the column placement (moving a label or swapping the rows of two
  labels), a growing penalty for overlaps and crossings, then repair and polish. Random numbers are MINSTD
  and the computation runs on a 1/1024 pt grid, so the result is deterministic.
- **Pre-adjustment** when the placement has defects or does not fit: a ladder of steps from the mildest one
  (font smaller by 0 to 2 pt in 0.5 pt steps, axis ranges 1.0 to 1.3× in 0.1 steps, the X axis or both); the
  first rung whose placement has no defects and the full gap between labels wins. A range is widened around the
  quadrant dividing lines (slide shapes that do not move with the axis), so every point stays in its quadrant.
  An axis with tick labels is never changed. The chart keeps its state before the pre-adjustment in shape tags;
  the next run first returns the axis ranges and the font size, so repeated runs on the same file (after a data
  update) do not add up.
- **Fallback** when the columns do not fit even then: a label whose column position is invalid may stay "not placed";
  the search places as many labels as possible without conflicts, labels not placed keep their original
  position and the placed ones move out of their way. With fewer than half of the labels placed, the chart is
  left unchanged.
- **Writing**: font size, a custom width for long labels, automatic size for the others, then `Left`/`Top`.
  After writing, the add-in reads the positions back and reports any difference above 0.5 pt.

## Building from source

Requirements: Windows, PowerPoint 2016 or Microsoft 365, Python 3.12+ and [uv](https://docs.astral.sh/uv/).

```powershell
uv run addin/build_addin.py dist\RozmistiPopisky.ppam   # close PowerPoint first
```

The script builds the add-in through COM: it inserts the VBA modules and the dialog code, compiles the project
(a compile error stops the build with the module and line) and adds the ribbon. Inserting code needs the
option *Trust access to the VBA project object model* (`HKCU\...\PowerPoint\Security\AccessVBOM`): the script
turns it on only for the build and restores it after PowerPoint exits, also after an error.

Source files are UTF-8. The VBA editor uses the system code page, so the build turns string literals with
non-ASCII characters into `ChrW()` calls; comments must be plain ASCII.

## Object model pitfalls

- `TextFrame2.AutoSize = 0` on a chart data label crashes PowerPoint. A custom width can be set directly
  through `Width` and `Height`; back to automatic size with `AutoSize = 1`.
- After a font change a label returns its old `Width/Height` and `Chart.Refresh` does not help. Writing `Left`
  forces the recalculation, so the position is written last. Writing `Width` also moves `Left`.
- A label PowerPoint does not draw returns `Left = NaN`; `Point.Left/Top` return nonsense.
- After changing an axis through the object model, labels keep returning their old `Left/Top`; the real
  position shows only after saving and reopening.
- A label with `c15:xForSave val="1"` is drawn in its default position.
- A path with mixed slashes (`C:/...\x.pptx`) opens, but saving fails with "file in use".
- A compile error opens a modal dialog that blocks PowerPoint; the build reads and closes it.
- VBA is case-insensitive: a local variable named like a constant silently shadows it.
- VBA evaluates a whole expression in extended precision and rounds only the result; this is why the
  shorter-lines placement computes on a 1/1024 pt grid.
- At 12 pt PowerPoint draws some labels up to 1.6 pt narrower than the GDI measurement (positions are exact).
