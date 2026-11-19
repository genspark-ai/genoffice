# GenOffice Sheets

An AI-native spreadsheet app for macOS and Windows.

## What is implemented

- Electron privilege separation with a sandboxed renderer and validated IPC.
- Univer OSS spreadsheet renderer.
- Desktop shell with inset traffic lights, a toolbar with AutoSave pill and Save/Undo/Redo, a ribbon without a File tab (file commands live in the macOS application menu), name box + formula bar, sheet tabs, and a collapsible AI panel.
- Rust sidecar for streaming XLSX metadata and bounded worksheet ranges.
- Viewport-driven loading with a bounded renderer cell window.
- Read-only XLSX cell styles, embedded images, and anchored chart visualization.
- Cache-based column, bar, line, pie, and bar-line combination chart rendering.
- Versioned, schema-validated workbook command DSL.
- Deterministic local planner for repeatable AI workflow tests.
- Side-effect-free preview, atomic transaction, revision conflict detection, and undo.
- XLSX package inventory and surgical worksheet mutation.
- Cell value, formula, and ribbon style editing on streamed external workbooks with preservation-checked save / Save As (⌘S / ⇧⌘S); style changes are written copy-on-write into styles.xml.
- Row/column insertion/deletion and merge/unmerge on fully-loaded workbooks, replayed into the file with reference shifting (including other sheets' formulas, defined names, and chart series) and fail-closed guards for tables, anchored charts, and deleted references.
- Worksheet management on external workbooks: rename (qualified references in formulas, defined names, chart series, and hyperlink anchors rewrite file-side), add blank sheets (part + content-type + relationship created, editable immediately), and delete (fail-closed while formulas, charts, or defined names still reference the sheet, or when it carries drawings/tables).
- Interactive sorting and filtering on fully-loaded workbooks, saved back as autoFilter XML (value lists, blank flags, custom conditions) plus declarative row visibility; color filters and table-owned filters fail closed.
- Per-edge border editing, full font/fill/border color pickers, and fill clearing, written copy-on-write into the stylesheet's borders/fills tables.
- Row heights, column widths, autofit resets, and hidden rows/columns saved as ordered axis-attribute operations; Fill Down/Right and Clear Contents/Formats/All (format clears reset cells to the default style in the file).
- Chart editing (title, column/bar/line/area type conversion, series colors) with surgical chart-part patching and live SVG preview.
- Preservation checks for untouched OOXML entries.
- Generated compatibility fixtures.

## Run

```bash
brew install rust
npm install
npm run dev
```

The local planner accepts:

- `set A1 to 42`
- `formula B4 = SUM(B2:B3)`
- `rename sheet to Budget`

## Interface

The desktop layout follows established spreadsheet conventions to reduce migration cost for Excel users. Ribbon areas that depend on unsupported workbook features are reserved but not presented as working commands. See [docs/interface.md](docs/interface.md) for the layout, AI interaction model, and current visual limitations.

## Safety boundary

The renderer cannot access Node.js or the filesystem. AI output is treated as untrusted input and must pass the DSL schema, revision checks, dry-run preview, and explicit approval. Arbitrary JavaScript and Python execution are intentionally excluded.

The UI opens external XLSX packages in streaming mode. The Rust sidecar indexes worksheet XML into temporary row chunks while Univer holds only the current viewport and its buffer. Styles are attached to streamed cells. Drawing relationships, chart caches, anchors, and image metadata are read from their small OOXML parts without loading large worksheets into memory. Embedded image bytes are fetched through a separate size-limited IPC request. Cell edits are journaled in the renderer and saved through the main process, which verifies the on-disk file hash, rewrites only the edited worksheet entries, and fails closed if any other package entry would change. The demo workbook retains the local AI preview, transaction, and undo workflow.

Run the large-workbook gate with:

```bash
npm run benchmark:large -- "/path/to/workbook.xlsx"
```

## Scope

See [docs/compatibility.md](docs/compatibility.md) for the frozen V1 feature matrix and [docs/architecture.md](docs/architecture.md) for process boundaries.
