# User guide

The add-in lays out the labels in the XY charts of a presentation: two aligned columns left and right of the
cluster of points, labels right next to isolated points, no overlapping labels and no crossing leader lines.
It unifies the font size of the labels in a chart. By default it creates a new presentation and leaves the
original file unchanged; it can also edit the open presentation in place.

## Installation (once)

1. Download [RozmistiPopisky.ppam](https://github.com/khoderTomas/powerpoint-xy-labels/releases/latest/download/RozmistiPopisky.ppam)
   and save it to a folder you will not move it from, for example `Documents\RozmistiPopisky`.
2. Right-click the file → Properties → tick **Unblock** → OK. Otherwise PowerPoint blocks its macros.
3. In PowerPoint: File → Options → Add-ins → at the bottom *Manage: PowerPoint Add-ins* → Go → Add New →
   choose `RozmistiPopisky.ppam` → enable macros if asked → Close.
4. A group **Chart labels** with the button **Lay out labels** appears on the Home tab.

Updating from an older version: close PowerPoint, replace the file with the new one and start PowerPoint again.
Up to version 0.6 the texts of the add-in were in Czech.

## Use

Open the presentation and click Home → **Lay out labels**. A dialog opens with three choices:

- **Which slides**: all, the slides selected in the thumbnails (select them before clicking the button,
  Ctrl+click for more), or slide numbers such as `2, 5-7`.
- **Placement**: aligned columns, or shorter lines (trial). Shorter lines
  put labels as close to their points as the space allows; in a dense cluster some labels stay in a column.
  It takes longer; for a presentation with dozens of charts expect a minute or more.
- **Result**: a new presentation next to the original, or edit the open presentation.

The dialog remembers the last placement and result.

With the new presentation, `name-labels.pptx` is saved next to the original (for files from OneDrive or
SharePoint in the TEMP folder; save it where you need it) and opened. Unsaved changes are included. The
original file stays unchanged.

With the open presentation, the labels are laid out in the presentation you have open and no other file is
created. The add-in does not save it: if you do not like the result, close it without saving. Undo (Ctrl+Z)
takes the change back in parts (four charts take about 13 steps), so it may not reach back over a whole
presentation. For files from OneDrive or SharePoint with AutoSave on, the change is saved automatically.
Before the edit the add-in saves a backup to `%TEMP%\RozmistiPopisky` as `name-backup-date-time.pptx` (the
three newest backups of a presentation are kept) and names it in the final message and in the report.

In both cases, slides you did not select stay unchanged.

When a chart stays unchanged or needs a check, a review report `name-labels-check.html` opens in the
browser (`name-check.html` when the open presentation was edited).

## Good to know

- Run it right after the data in the presentation is updated, before manual edits. A new run recalculates every XY
  chart on the selected slides and overwrites manual fine-tuning of the labels.
- To keep the add-in away from a slide, write `#labels-off` into the slide notes (the marker of older versions,
  `#popisky-ne`, still works).
- When the placement does not come out without defects, the add-in first adjusts the chart slightly: a smaller
  label font (at most 2 pt) and wider axis ranges (at most 1.3×, only axes without tick labels). It starts with
  the smallest change and stops at the first one that gives a placement without defects. Ranges are widened
  around the dividing lines, so every point stays in its quadrant; data points are never changed. The report
  lists the change as *pre-adjustment*.
- The chart remembers its state before the pre-adjustment. The next run (after a data update) first returns the
  axis ranges and the font size and adjusts them again only as needed, so the changes do not add up.
- When even that does not help (points spread over the whole chart), the add-in places as many labels as
  possible next to their points. Labels that still do not fit stay where they were and the report lists
  them for manual fine-tuning. When fewer than half of the labels fit, the chart stays unchanged and the
  report marks it.
- The resulting labels are ordinary PowerPoint labels; they can still be moved with the mouse.
