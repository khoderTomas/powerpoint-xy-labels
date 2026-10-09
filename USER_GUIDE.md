# User guide

The add-in lays out the labels in the XY charts of a presentation: two aligned columns left and right of the
cluster of points, labels right next to isolated points, no overlapping labels and no crossing leader lines.
It unifies the font size of the labels in a chart. It always creates a new presentation; the original file
stays unchanged. The add-in's own texts are in Czech; the names below are given in Czech with a translation.

## Installation (once)

1. Download [RozmistiPopisky.ppam](https://github.com/khoderTomas/powerpoint-xy-labels/releases/latest/download/RozmistiPopisky.ppam)
   and save it to a folder you will not move it from, for example `Documents\RozmistiPopisky`.
2. Right-click the file → Properties → tick **Unblock** → OK. Otherwise PowerPoint blocks its macros.
3. In PowerPoint: File → Options → Add-ins → at the bottom *Manage: PowerPoint Add-ins* → Go → Add New →
   choose `RozmistiPopisky.ppam` → enable macros if asked → Close.
4. A group **Popisky grafu** (Chart labels) with the button **Rozmístit popisky** (Lay out labels) appears
   on the Home tab.

## Use

Open the presentation and click Home → **Rozmístit popisky**. A dialog opens with two choices:

- **Which slides**: all, the slides selected in the thumbnails (select them before clicking the button,
  Ctrl+click for more), or slide numbers such as `2, 5-7`.
- **Placement**: aligned columns (*Zarovnané sloupce*), or shorter lines (*Kratší čáry*, trial). Shorter lines
  put labels as close to their points as the space allows; in a dense cluster some labels stay in a column.
  It takes longer; for a presentation with dozens of charts expect a minute or more. The dialog remembers
  the last placement.

The new presentation `name-popisky.pptx` is saved next to the original (for files from OneDrive or
SharePoint in the TEMP folder; save it where you need it) and opened. Unsaved changes are included.
Slides you did not select stay unchanged.

When a chart stays unchanged or needs a check, a review report `name-popisky-kontrola.html` opens in the
browser.

## Good to know

- Run it right after the presentation is generated, before manual edits. A new run recalculates every XY
  chart on the selected slides and overwrites manual fine-tuning of the labels.
- To keep the add-in away from a slide, write `#popisky-ne` into the slide notes.
- When the columns do not fit (points spread over the whole chart), the add-in places as many labels as
  possible next to their points. Labels that still do not fit stay where they were and the report lists
  them for manual fine-tuning. When fewer than half of the labels fit, the chart stays unchanged and the
  report marks it.
- The resulting labels are ordinary PowerPoint labels; they can still be moved with the mouse.
