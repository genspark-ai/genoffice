# Desktop interface

The desktop shell follows familiar spreadsheet conventions so existing Excel users can transfer their navigation habits without learning a new workspace model. It uses its own product identity and does not reproduce Microsoft assets or trademarks.

## Layout

The shell layout, from top to bottom:

1. A light toolbar row with the macOS traffic lights inset (`titleBarStyle: hiddenInset`), an AutoSave pill (off — local files), Save/Undo/Redo icons, the centered workbook name with save state, and Open XLSX. The row is a window-drag region.
2. Ribbon tabs for Home, Insert, Page Layout, Formulas, Data, Review, View, and AI. There is no File tab — file commands live in the macOS application menu (File → Open ⌘O / Save ⌘S / Save As ⇧⌘S), forwarded to the renderer over IPC.
3. A compact single-band ribbon with hairline group separators and no group captions (group names remain as `aria-label`s). Home exposes clipboard, font, alignment, number, styles, cells, and editing groups.
4. Univer's name box + formula bar (preset `header: true, toolbar: false` renders only these), worksheet canvas, sheet tabs, statistics, and zoom controls.
5. A right-side Workbook Copilot panel that can collapse to a narrow rail.

The ribbon establishes the final command hierarchy. Commands whose workbook capability is not implemented are presented as reserved surfaces rather than active operations. In development builds the application menu title shows "Electron" — fixing that requires packaged builds (P0).

## AI interaction

The AI panel is secondary to the worksheet and can be collapsed to maximize working space. The local demo supports:

- natural-language command planning;
- change preview before mutation;
- explicit apply or discard;
- revisioned undo.

External XLSX files support direct cell value/formula editing and ribbon style edits, saved back with preservation checks. AI editing remains demo-only: the panel's AI inputs are disabled for imported files and display the save/streaming state.

## Large workbooks

Opening an XLSX creates only worksheet skeletons in Univer. The visible range and a bounded buffer are loaded from the Rust sidecar. Switching sheets and scrolling request new sparse ranges, while the previous renderer range is cleared.

The UI therefore keeps the same worksheet, formula-bar, and tab navigation model for small and large files without loading every cell into the renderer.

## Workbook visuals

Streamed cells display their source font, fill, alignment, and number format where those values can be resolved directly from `styles.xml`.

Images and charts remain anchored to worksheet ranges while the user scrolls. Embedded images load on demand. Column, horizontal bar, line, pie, and bar-line combination charts render from the cached series stored in the XLSX package. This is a read-only compatibility view, not an editable native chart designer.

## Current limitations

The main interface-level gaps:

- row/column inserts/deletions and merge/unmerge save on fully-loaded workbooks (≤50k cells) only; sheets with tables or anchored charts fail closed at save time (cross-sheet references are rewritten automatically);
- sheet rename/add/delete save on any external workbook; range moves remain blocked;
- complete Excel chart themes, effects, 3D variants, and chart/image editing panels;
- pattern-fill editing and diagonal-border controls (edge borders, full color pickers, and fill clearing shipped);
- page layout and printing configuration.

Unsupported drawing shapes are surfaced as placeholders. The corresponding editing controls remain reserved for later compatibility milestones.
