# AGENTS.md

Suite-wide theming, shell-build, and i18n rules live in `CLAUDE.md`. Follow those
for every app. Remaining Hangul work is listed in `apps/hwp/todo.md` — do not
treat that file as “won’t do” unless an item is under **하지 않음**.

This file is the Hangul / rhwp contract. Breaking it blanks the editor or
silently corrupts documents.

## Embed

Hangul tabs host self-vendored **rhwp-studio 0.8.6** (`@rhwp/editor`). GenOffice
owns chrome (tabs, open/save, recents, AI dock). Studio owns typing, tables,
formatting.

```ts
createStudio(host, {
  studioUrl: new URL('rhwp/?chrome=embed', location.href).href,
  renderer: 'canvas2d',
})
```

- After `createStudio`, the host **must** `loadFile` or `commands.execute('file:new-doc')`. Embed boot does not create a document.
- Open: `loadFile(bytes, name, { skipUnsavedGuard: true, suppressDialogs: true })`.
- Save: `exportHwp` / `exportHwpx` / `exportHml` → IPC atomic write → `notifySaved`.
- Dirty: poll `getDocumentState().dirty` every 750ms. That is the SDK path; do not invent a replacement event.
- `.hwp-studio` is `position: relative`; the iframe is `position: absolute; inset: 0`.

## Vendor snapshot

- Source of truth for the patch: `apps/hwp/scripts/studio-snapshot.mjs`.
- Download / ensure: `apps/hwp/scripts/vendor-studio.mjs` (`npm run vendor:studio -w @genoffice/hwp`).
- Output `apps/hwp/vendor/rhwp-studio/` is **gitignored**. After vendor or preload changes, restart `npm run dev`. If the Hangul pane is blank, run `node apps/hwp/scripts/vendor-studio.mjs --ensure` and restart.
- Do **not** `npm install -w @genoffice/hwp` alone. Workspace packages in an app’s `dependencies` must also be on the shell `externalizeDepsPlugin` `exclude` list.
- App `src/main` and preload compile into the **shell** build. A Hangul-only rebuild will not pick them up.

## Studio patch (prepare / list / edit)

`getSelectionContext` already computes SHA fences in minified `Gk(wasm, target)`
and then drops them (`exactKeys`). The host cannot add fields there.

`exposePrepareTextCommand` injects **class methods** on the Document object,
before `async applyTextCommand`:

`prepareTextCommand`, `listBodyParagraphs`, `listFields`, `setField`,
`listTables`, `replaceCell`

The host calls private `studio._request(...)`, then public `applyTextCommand` +
`focusTarget`. Mark: `/*genoffice-prepare-text-v3*/`. Table scan cap:
`TABLE_CONTROLS_PER_PARA = 8`.

Needles throw if the upstream bundle shape changes. Update the patch; do not
bypass with ad-hoc WASM from the host.

### Class bodies must not have commas

Object literals and `switch` cases need commas. **Class method lists must not.**

```js
// WRONG — blank Hangul page, Enter does nothing
prepareTextCommand(){...},listBodyParagraphs(){...}

// RIGHT
prepareTextCommand(){...}listBodyParagraphs(){...}
```

A comma here made `index-*.js` fail to parse. `stripClassMethodCommas()` repairs
already-patched vendor. Completeness check: `listTables(){this.syncGeneration()`
exists **and** there is no `,listBodyParagraphs(){this.syncGeneration()}`. The
vendored bundle must acorn-parse as a module. Snapshot test asserts
`not.toMatch(/,listBodyParagraphs\(\)\{this\.syncGeneration\(\)/)`.

## Host edit rules (`studio-text.ts` / `hangul-skill.ts`)

- Body replace is UTF-16 (`String.slice` / `spliceParagraphText`). Out of range
  **throws**. Do not add an `Array.from` code-point fallback.
- Re-list paragraphs immediately before each `replace_paragraph`. Adjacent SHAs
  change after every apply.
- `listDocumentFields` / `listDocumentTables` / `listBodyParagraphs`: non-array
  → throw `PARAGRAPH_PREPARE_UNAVAILABLE`. Never swallow as `[]`.
- `set_field`: WASM `setField` first. `hwpctrl.PutFieldText` only after a
  **successful** list. A missing list RPC is not “no fields”.
- Caps: body 4000 code points (`applyTextCommand` / `exactKeys` contract).
  Fields and cells 8000. No C0, no `\n` in replacements.
- `requireToolIndex`: `null` / `''` / whitespace is an error, not `0`.
  `optionalToolIndex` (paragraph `index`) treats `null` / `''` as caret.
- Cell replace edits the **first** paragraph in the cell only.

## Do not

- Lift the 4000-character body cap.
- Use `SetTextFile` (or any whole-document write) as a general editor.
- Re-apply page-turn / caret-below-page studio patches (`eb3c4f8` reverted them).
- Call Hangul WASM header/footnote APIs from the host. Undo and layout break.

## Print / PDF

rhwp-studio already has `file:print` and `file:print-to-pdf`. It paints pages as
SVG into sibling `print.html` and opens the browser print dialog. Embed chrome
hides the studio File menu. The vendor snapshot does **not** ship `print.html`.
Add an empty same-origin `print.html` shell, wire the File menu with
`commands.execute`, and include that file in the snapshot. This is host work,
not a new PDF engine.

## AI / i18n

Tools: `get_document_text`, `get_selection`, `get_paragraphs`,
`replace_paragraph`, `replace_selection`, `get_fields`, `set_field`,
`get_tables`, `replace_cell`, plus composed `web_search`.

`aiEmptyBody` in all 19 locales describes how to edit (selection, paragraph,
field, cell). Do not restore “편집 불가”. New keys go in `apps/hwp/src/renderer/i18n/ai/zh.ts` and every sibling shard.
