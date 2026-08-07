# XLSX compatibility contract

The product claim for V1 is:

> Open, edit, and save common XLSX workbooks with high fidelity. Preserve or explicitly warn about unsupported advanced content.

“Can open” is not sufficient. Every supported feature must pass no-op round-trip, edited round-trip, structural validation, semantic comparison, and Excel visual review.

## V1 supported target

- Scalar values, dates, booleans, errors, shared and inline strings.
- A1 formulas and cross-sheet references within the published function matrix.
- Fonts, fills, borders, alignment, number formats, row heights, and column widths.
- Merged cells, hidden rows/columns/sheets, frozen panes, sort, and filter.
- Conditional formatting and data validation in the published subset.
- Tables, defined names, comments, hyperlinks, and images.
- Column, bar, line, area, pie, and scatter charts in the published subset.
- Workbook calculation settings and cached formula values.

## Preserve but do not execute

- VBA projects in XLSM packages.
- External workbook links and data connections.
- Power Query, Power Pivot, and embedded data models.
- OLE objects, ActiveX controls, slicers, and unsupported chart extensions.
- Unknown vendor extension and custom XML parts.

Any save operation that cannot preserve these parts must fail closed. It must not silently produce a reduced workbook.

## Explicitly unsupported in V1

- Editing or executing VBA.
- Refreshing Power Query or external data.
- Editing Power Pivot models, ActiveX, OLE, and form controls.
- Pixel-identical print layout on every font and printer configuration.
- Guaranteed compatibility with every Excel private extension.

## PoC capability

The current PoC intentionally supports a narrower write set:

- Edit, clear, or add cell values and formulas in streamed external workbooks, saved back (or Save As) with a whole-file SHA-256 drift check; only the edited worksheet entries, `xl/styles.xml` (when styles changed), and `xl/workbook.xml` (forced recalculation flag) are rewritten, every other ZIP entry is verified byte-identical or the save fails closed. Cell style indices survive edits; shared-string cells are rewritten as inline strings; new rows and cells are inserted in schema order.
- Ribbon style edits (bold/italic/underline/strikethrough, font family/size, arbitrary font/fill colors with a full picker, fill clearing, per-edge borders with 13 line styles and a border color, alignment, wrap, number formats) on streamed workbooks, written copy-on-write into `styles.xml`: new deduped cellXfs/fonts/fills/borders derived from each cell's original format (untouched border edges and diagonal attributes survive verbatim), built-in number-format ids mapped, custom ids allocated from 164. Style-only edits leave cell content byte-identical. Pattern-fill editing is deliberately out of scope (no editor model for it).
- Rename a worksheet (demo workbook DSL flow).
- Preserve the uncompressed content hash of untouched ZIP entries.

- Insert or delete rows/columns and merge/unmerge cells in every load mode — streamed large workbooks included: viewport reads translate screen ↔ file coordinates through the journaled operation stream, and operations replay in order at save time, renumbering rows/cells and shifting same-sheet formula references (absolute and relative, ranges with Excel clamp/expand semantics, whole-row/column refs, shared/array formula anchors, CF/DV rule formulas), merges, dimension, col widths, hyperlinks, autoFilter, and CF/DV ranges. Qualified references to the edited sheet in other sheets' formulas, workbook defined names, and chart series (`c:f`) shift along with it. A stale calcChain is dropped with its content-type and relationship. Entries too large to rewrite (>256MB) are streamed through an XML text scan: if they cannot reference the edited sheet they are raw-copied untouched, otherwise the save fails closed. Saves also fail closed when a reference points into a deleted range or when the edited sheet carries tables or drawings.

- Chart edits (works on streamed large files too): title text (first-run styling preserved, or inserted with `autoTitleDeleted` reset), chart type conversion within the column/bar/line/area family (series and axis references kept, cross-type series options stripped), and series solid colors (fill + line). Pie/scatter/combo charts accept title and color edits only. Unsupported shapes fail closed.

- Sheet duplication (fully-loaded workbooks, or sheets created this session): the tab context menu's Copy journals a `duplicate-sheet` op; the save clones the source worksheet part verbatim (tabSelected stripped), replays the source's unsaved edits captured at copy time, and appends the copy's own edits. Hyperlink relationships clone with the part; printerSettings relationships are dropped (with their `pageSetup` r:id). Sources carrying drawings, tables, comments, or pivots, and sources with sheet-scoped defined names, fail closed. Duplicate-of-duplicate chains resolve to the original file part.

- Sheet management (works on streamed large files too): rename — the workbook.xml entry plus every qualified reference (other sheets' `<f>`/CF/DV formulas, defined-name bodies, chart `c:f` series, internal hyperlink anchors) rewrite together, with Excel quoting rules for the new name; add — a blank worksheet part with its content-type override and relationship, inserted at its on-screen tab position, immediately editable (values, styles, and row/column ops journal into the new part); remove — the part, its rels, its content-type override, and its relationship are dropped, defined names scoped to the sheet die with it, surviving `localSheetId` scopes and the active tab renumber. Removal fails closed while any surviving formula, chart, or defined name references the sheet, when the sheet carries drawings/tables/comments, or when no visible sheet would remain. Adding or removing sheets invalidates calcChain, which is dropped for Excel to rebuild. Duplicating, hiding, and reordering sheets remain blocked.

- Whole-row moves (drag a selected row's header; fully loaded sheets only): journaled as a move operation modeled as two adjacent row blocks swapping places — a bijection, so nothing is deleted and no reference ever turns into #REF!. The save relocates the `<row>` elements themselves (heights, hidden flags, outline levels, and cell styles travel with their rows, sheetData is re-sorted ascending) and remaps every reference through the swap: same-sheet and cross-sheet formulas, defined names, chart series, merges, CF/DV ranges, hyperlinks, drawing anchors (judged as from/to pairs), and table refs. Ranges fully inside a moved block move with it; ranges outside or spanning the blocks stay put; anything partially overlapping fails the save closed (shared/array formula anchors, table header/totals rows, and torn drawing anchors included). Undo arrives as the inverse move and cancels the journal pair. Column moves remain blocked.
- Row heights, column widths, autofit resets, and hidden rows/columns (every load mode, including sheets with tables or charts — sizing shifts nothing): journaled as ordered axis-attribute operations and replayed into row `ht`/`customHeight`/`hidden` attributes and split-preserving `<cols>` rewrites. Clear Formats / Clear All reset cells to the default xf (`s="0"`); a format applied after a clear derives from the default style.

- Outline groups (Data → Group/Ungroup/Hide Detail/Show Detail, every load mode): row/column `outlineLevel` and `collapsed` are parsed by the sidecar (rows stream with range reads, columns ride the sheet metadata) and tracked per sheet; Group/Ungroup shifts each contiguous equal-level run in the selection by ±1 (clamped 0-7) and journals declarative `set-rows/cols-outline` ops, replayed into row attributes and split-preserving `<cols>` rewrites with `sheetFormatPr` `outlineLevelRow`/`outlineLevelCol` recomputed. Hide/Show Detail hides or shows the selected detail span through the normal hidden-rows pipeline and toggles `collapsed` on the following summary line (Excel's summaryBelow/Right default). Level edits journal directly (Univer has no outline model), so they are not undoable and no +/- gutter renders — collapsed groups appear as hidden rows/columns.

- Conditional-formatting rule create/edit/delete through Univer's rule panel, saved declaratively: every `<conditionalFormatting>` section of an edited sheet is rewritten from the live rule model (priorities renumbered), highlight styles intern as new dxfs. Supported: cellIs, text operators (with Excel-compatible formula bodies), blanks/errors, duplicate/unique, top10, above/below average, expression, color scales, data bars, and the OOXML icon sets. Fail closed: x14 extended rules, date-occurring rules, Univer-only icon sets, mixed/custom icon orders. Editing CF on a sheet re-serializes its untouched rules through the model — dxf attributes the reader does not map (borders, number formats) are dropped on that sheet.

- Data-validation rule create/edit/delete through Univer's validation panel, saved declaratively: the `<dataValidations>` section of an edited sheet is rewritten from the live rule model. The reader parses the full attribute set (allowBlank, showDropDown, prompts, error style/title/text) and installs rules verbatim into Univer — install fidelity is save fidelity — including cross-sheet list references and `time` rules (which render without a validator but round-trip byte-exact). Mappings are bijective: OOXML `none`↔Univer `any`, list literals `"a,b"`↔`a,b`, reference/custom formulas gain/lose a leading `=`, panel-edited date/time strings become Excel serials, and OOXML's inverted `showDropDown` flag is normalized. Fail closed: x14 extended validations and Univer-only multi-select list rules. Checkbox rules (Insert → Checkbox) degrade on save to a two-value OOXML list — the default 1/0 pair saves as `"1,0"` and is restored to a checkbox rule on reopen; Excel shows a plain 1/0 dropdown validation. Editing is gated until the sheet's own file rules are installed (indexing complete), so a declarative rewrite can never drop rules the model has not seen.

- Defined names (Name Manager): add, edit, and delete workbook- and sheet-scoped names, saved by rewriting `<definedNames>` declaratively. `_xlnm.*` built-ins, hidden names, and names the engine could not model are preserved byte-verbatim; name rules are validated (fail closed); defined-name edits cannot be saved in the same pass as row/column or sheet structure changes.

- Worksheet protection toggle (Review tab): protecting writes Excel's default `<sheetProtection sheet="1" objects="1" scenarios="1"/>` (no password); unprotecting removes the element and fails closed when the sheet is password-protected. The file's protection state is parsed by the index thread and echoed on the button after indexing. The editor itself does not enforce protection — combined with the Format Cells locked/hidden flags, enforcement happens in Excel.

- Hyperlink add/edit/remove at any cell: external targets write `<hyperlink r:id>` plus a TargetMode="External" relationship (the rels part is created on demand and orphaned rels are collected); internal `Sheet!A1` anchors write `location`. One link per cell; tooltip/display attributes on untouched links are preserved.

- Format Cells dialog (⌘1) covers number format, alignment (horizontal including justify/distributed, vertical, wrap, orientation presets, indent), font, border, fill, and cell protection flags. Indent renders as left cell padding (`pd.l`, `INDENT_STEP_PX` px per step), is undoable, journals through the normal set-range-values channel, saves into the xf's `<alignment indent>`, and reads back from opened files. Locked/hidden protection still has no Univer model: it journals directly and saves into `<protection>` (defaults omitted), but is not rendered on screen, not undoable, and not enforced by the editor — it takes effect when Excel protects the sheet.

- Text rotation (OOXML `textRotation` 0-180 and 255 stacked) and double underline (`<u val="double"/>`) through the same copy-on-write style pipeline. Text to Columns splits a single selected column by comma/semicolon/space/tab (fully-loaded workbooks; the split's cell writes and any inserted columns ride the existing journal). File-side text rotation is write-only for now — existing rotations render unrotated.

- Table creation (AI `add_table` on imported workbooks): renders as a native Univer table (banded styles, filter dropdowns) and saves as a brand-new `xl/tables/tableN.xml` part with the worksheet `<tableParts>` hookup, relationship, and content-type override. Column names are captured from the on-screen header row at apply time, sanitized to Excel's rules (blank → ColumnN, duplicates suffixed) and written back into the cells so the part and the sheet agree. Fail-closed at save: table/defined-name collisions, overlap with existing tables / the sheet auto-filter / merged cells, blank or duplicate column names, and a same-save row/column shift on the table's sheet. Editing or deleting tables already in the file, structured references in formulas, and rendering of pre-existing table styles remain unsupported.

- Pivot creation (AI `add_pivot` on imported workbooks): the apply aggregates the source range in the renderer (sum/count/average/max/min; one row field, optional column field) and bakes the grid — grand totals included — into the target cells as ordinary journaled writes. The save adds native pivot parts: a pivotCacheDefinition with `refreshOnLoad="1"` and zero records, an empty pivotCacheRecords, and a pivotTableDefinition whose `<location>` covers the baked grid, wired through workbook `<pivotCaches>`, workbook/worksheet/pivot rels, and content-type overrides. Excel (and LibreOffice — round-trip verified) rebuilds the cache from the source on open, turning the baked grid into a live pivot. Fail-closed at save: duplicate pivot names, field indexes outside the source, and same-save row/column shifts or sheet management on the involved sheets. Reopening the saved file protects the pivot output area from edits (existing behavior).

- Sorting and filtering (fully-loaded workbooks only): sorting reorders the model and journals the affected range as cell edits; filter interaction (create/edit/clear) snapshots the live filter into `<autoFilter>`/`<filterColumn>` XML — value lists, blank flags, and custom conditions (the six OOXML operators plus wildcard values) — with filtered-out rows written as `hidden="1"` declaratively inside the filter's row span. Color filters have no XLSX mapping and abort the save; filters owned by Excel tables are blocked from editing (their state lives in the table part).

Range moves are not saveable and are actively blocked on imported workbooks. Images and shapes already in the file can be moved, resized, and deleted: the drag/✕ editor journals a surgical edit against the visual's (drawingPath, anchor index) locator, and the save rewrites or removes just that `<xdr:*CellAnchor>` element (orphaned image rels/media stay — harmless, preserved bytes). Deleting a chart cascades: the anchor, its drawing relationship, the chart part with its own rels, and the content-type override all go (chart-owned colors/style parts stay as harmless unreferenced entries); the delete fails closed while another anchor still references the chart, and non-chart graphic frames stay view-only. Absolute anchors fail closed on move. Chart series data ranges are editable (inline ✎ editor and AI `edit_chart`): the new single-row/column range is read from the sheet (streamed files fetch through the sidecar with journal edits overlaid), baked into the series cache, and saved as a rewritten `<c:f>` + `numCache`/`strCache` with the original `formatCode` preserved; scatter charts (xVal/yVal) stay read-only. Saving streams through the Rust sidecar: untouched entries are raw-copied compressed bytes (verified by CRC manifests), so total workbook size is unbounded — only an individual entry being patched must be ≤256MB uncompressed; edits to larger entries fail closed.

External XLSX files additionally support read-only streaming of worksheet dimensions, scalar values, shared and inline strings, booleans, formulas, cached formula values, and source style indices. Fonts (including underline and strikethrough), fills, borders, alignment (including wrap text), and the full ECMA-376 built-in number-format table are mapped into the worksheet renderer. RGB, indexed-palette, and theme+tint colors are resolved in the sidecar (theme1.xml clrScheme with the Excel index swap and HSL luminance tint). Merged cells, frozen panes, hidden rows and columns, custom row heights, hidden sheets, sheet tab colors, per-sheet gridline visibility, and default column widths are applied to the rendered worksheet. Rich-text runs in shared strings render as per-run styled cell text. Hyperlink cells (external and internal) render with link styling. Conditional formatting — cellIs operators, text operators, blanks, duplicate/unique, top10, expression, color scales, and data bars — is evaluated live by the Univer conditional-formatting engine from parsed rules and dxf styles. Worksheet tabs and large sparse ranges are displayed without creating a complete JavaScript workbook snapshot.

Two-cell drawing anchors, embedded PNG/JPEG/GIF/BMP/SVG images, and cached column, horizontal bar, line, area, scatter, pie, doughnut, and bar-line combination charts (with multi-series grouping, a series legend, y-axis ticks with gridlines, and value labels on small series) are displayed on Univer's floating object layer. Basic unsupported shapes receive a visible placeholder instead of silently disappearing.

This milestone does not claim native Excel chart fidelity. For streamed external workbooks specifically, chart themes, secondary axes, trendlines, 3D effects, complex DrawingML shapes, icon-set conditional formatting, tables/autofilter, comments, hyperlink click-through, inline-string rich runs, structural row/column edits, formula calculation, and editing or saving remain unsupported and are planned for later milestones (the native open/edit path above is unaffected).

## Release evidence

A V1 feature is supported only when:

1. Golden files cover creation in supported Excel versions.
2. No-op save preserves semantic structure.
3. A targeted edit preserves unrelated OOXML parts.
4. Open XML validation reports no new errors.
5. Excel opens the output without a repair warning.
6. Formula and visual comparisons pass documented thresholds.
7. Unsupported content is preserved or the save is rejected.
