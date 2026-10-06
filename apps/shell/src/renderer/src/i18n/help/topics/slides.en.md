# Slides: presentations

Slides is the PowerPoint-like editor: reads and writes genuine.pptx.

## The interface

- **Ribbon**: the Home tab carries insert & format (text box/shape/picture/table/chart), font & paragraph, align & arrange (z-order, align, distribute).
- **Thumbnail rail** (left): click to switch, drag to reorder, context menu for new/duplicate/delete.
- **Canvas**: WYSIWYG editing; drag, scale handles, guides.
- **Notes**: per-slide speaker notes, kept in export flows.

## Ribbon tabs

The tab strip (macOS starts at Home; Windows adds a File tab):

- **Home**: insert & format — text box, shapes, pictures, tables, charts; font & paragraph; align & arrange (z-order/align/distribute); layout.
- **Insert**: text box, table (rows/columns selectable), pictures, page number, a jump button (click it while presenting to jump to a slide) and more.
- **Draw**: the **pen** (hand-draw on the slide, saved as ink on the page) and the **highlighter** (translucent, thicker strokes), with pen width; click the tool again to cancel.
- **Design**: themes, color scheme and background; masters and layouts.
- **Transitions**: pick a transition for the current slide (in effect in PowerPoint's presenter), with apply-to-all; None removes it.

![The Transitions tab](img/slides-transitions.png)

- **Animations**: entrance/emphasis effects for the selected shape, **motion paths** (move along a path); **preview** plays the slide's animations on the canvas; None removes them.

![The Animations tab](img/slides-animations.png)

Try it: select the title text box ▸ Animations tab ▸ pick an entrance effect ▸ **Preview** plays it on the canvas.

- **Slide Show**: present from the start or the current slide, plus show settings.
- **Review**: **new comment** on the current slide (written into the pptx, visible in PowerPoint).
- **View**: **Normal** (thumbnails + canvas), **Outline** (browse and jump by text), **Slide Sorter** (grid overview, double-click to edit), **Reading** (full-screen, page by page; Esc exits); plus **Slide Master**, **Presenter View**, **Custom Show**, **Hide Slide**, and toggles for **Ruler**, **Gridlines**, **Guides**, **Notes** and the **Thumbnail Pane**. Zoom in/out/100% and **Fit Window** live here too.
  - Those toggles are per session — they come back on their default when you reopen the app. The exceptions are the ones you would not want to lose: **guide positions** are saved with the deck, and the thumbnail pane remembers its width.
  - **Custom Show** picks which slides a kiosk run plays, and is saved with the deck.

## Right-click menus

The menu follows where you right-click:

- **On empty canvas**: new slide (layout selectable), cut/copy/paste the slide (also **paste as picture** and **paste keeping source formatting**), duplicate/delete/hide slide, **add section** (before), section rename/move up/down/remove/collapse-all/expand-all, **format background** / **change background image**, reset slide layout.
- **On a selected element**: edit text, hyperlink, size & position, **align** (left/center-h/right/top/center-v/bottom) and distribute horizontally/vertically, bring to front / send to back, group/ungroup/regroup, flip horizontal/vertical, crop picture / replace picture / save as picture, change shape, **set as default shape**, edit points; text elements also get bold/italic/underline/bullets/numbering.
- **On a selected table**: insert rows/columns (above/below/left/right), delete rows/columns, merge / merge right / merge down, split cell, **cell shading**, cell content anchoring (top/middle/bottom).

## Content

- Text boxes, shapes (fill/stroke/shadow), pictures (crop/replace), tables, charts (column/bar/line/pie..., editable data).
- Masters and layouts: unify fonts, placeholders and backgrounds; new slides inherit the chosen layout.
- Theme colors and fonts follow the theme.

## AI generation

- Home's AI Slides card: give a topic or outline and the AI builds the deck; cloud generation falls back to local generation on failure.
- Keep adjusting with the AI panel afterwards.

## Present and export

- **Export to PDF**: page-by-page rasterization in a hidden window, with progress and a timeout watchdog for big decks.
- Image export: PNG per page.
- Full-screen presenting where the version provides it.

## Saving

Full.pptx round-trip: untouched pages stay byte-identical; notes, masters and annotations persist.
