# AGENTS.md

Suite-wide theming, shell-build, and i18n rules live in `CLAUDE.md`. Remaining
Hangul work is in `apps/hwp/todo.md`. Do not treat that file as “won’t do”
unless an item is under **하지 않음**.

This file is the Hangul / rhwp contract. Breaking it blanks the editor or
silently corrupts documents. Implementation lives in `apps/hwp`.

## Embed

Hangul tabs host self-vendored rhwp-studio (`@rhwp/editor` 0.8.6). GenOffice
owns chrome. Studio owns typing, tables, formatting.

- `createStudio(host, { studioUrl: new URL('rhwp/?chrome=embed', location.href).href, renderer: 'canvas2d' })`
- After create, the host **must** `loadFile` or `commands.execute('file:new-doc')`. Embed boot does not create a document.
- Open: `loadFile(bytes, name, { skipUnsavedGuard: true, suppressDialogs: true })`.
- Save: export HWP/HWPX/HML → IPC atomic write → `notifySaved`.
- Dirty: poll `getDocumentState().dirty`. That is the SDK path.

## Vendor

Patch source: `apps/hwp/scripts/studio-snapshot.mjs`. Snapshot output
`apps/hwp/vendor/rhwp-studio/` is **gitignored**. After vendor or preload
changes, restart `npm run dev`. Blank pane: `node apps/hwp/scripts/vendor-studio.mjs --ensure` then restart.

Do not `npm install -w @genoffice/hwp` alone. Hangul `src/main` and preload
compile into the **shell** build.

## Why we patch the studio

`getSelectionContext` computes SHA fences and drops them. The public SDK cannot
add those fields back. The patch injects sibling Document **class methods**; the
host calls `studio._request`, then public `applyTextCommand` / `focusTarget`.

If the minified bundle shape changes, needles must throw. Update the patch.
Do not call Hangul WASM from the host as a bypass.

**Class method lists must not have commas.** Object literals and `switch` cases
do. A comma between class methods makes the vendor bundle fail to parse — blank
Hangul page, Enter does nothing. `stripClassMethodCommas()` repairs that. The
snapshot test locks it.

Body edits go through `applyTextCommand` (no `\n`, 4000 code-point cap). New
paragraphs are inserted first, then each is filled. Re-list immediately before
every apply — adjacent hashes change. List RPCs that are not arrays throw; never
swallow as `[]`. `PutFieldText` is only a fallback after a successful field list.

Studio cell/table/page calls also go through those Document methods, not host
WASM: `applyCharFormatInCell` / `applyParaFormatInCell` (`apply_format` with
`table`+`row`), `insertTableRow` / `insertTableColumn` / `deleteTableRow` /
`deleteTableColumn` / `mergeTableCells` / `splitTableCellInto` (`edit_table`),
`setCellProperties` / `setTableProperties` (`style_table`), `getPageDef` /
`setPageDef` / `getColumnDef` / `setColumnDef` (`set_page`). Cell char ranges
are UTF-16 (`string.length`), matching WASM offsets. `table: 0` is a valid
index — do not treat it as omitted. After a JS snapshot patch, run
`node apps/hwp/scripts/vendor-studio.mjs --ensure` and close/reopen the Hangul
tab.

## Do not

- Lift the 4000-character body cap.
- Use `SetTextFile` (or any whole-document write) as a general editor.
- Re-apply page-turn / caret-below-page studio patches (`eb3c4f8` reverted them).
- Call header/footnote WASM from the host. Undo and layout break.
- Treat `null` / `''` / whitespace required indexes as `0`.
- Restore `aiEmptyBody` to “편집 불가”. Every locale’s empty-state line must
  describe the current edit tools, including new paragraphs, tables, row/column
  edits, cell fill, and page setup. New AI strings go in
  `apps/hwp/src/renderer/i18n/ai/zh.ts` and every sibling shard.

## Print / PDF (later)

rhwp-studio already has `file:print` / `file:print-to-pdf`. They need a sibling
`print.html` (not in our snapshot) and a shell File menu hook. Not a new PDF
engine. Not current work.
