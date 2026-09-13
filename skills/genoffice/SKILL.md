---
name: genoffice
description: Create, convert, read and edit Office documents locally with GenOffice's command line. Build a new presentation (pptx) from a brief through a checked outline, per-page spec and render pipeline, a new spreadsheet (xlsx) from CSV or JSON data with formulas, a new Word document (docx) from Markdown or HTML, or a PDF; convert between pdf, docx, xlsx, pptx, md, html and csv; read the structure and text of an existing file (including the user's own docx / xlsx / pptx as source material) and apply structured edits to it. Use whenever the user asks for a slide deck, presentation, spreadsheet, workbook, report, Word document or any real Office file, a format conversion, a rewrite of part of an existing document, or wants the result opened in the GenOffice editor. Documents are processed locally; only search, image and media send the query or the referenced file to the provider configured in GenOffice.
metadata:
  version: 2.7.1
  cli: '>=0.6.0'
---

# genoffice — GenOffice from the terminal

`genoffice` is installed with GenOffice. Run `genoffice --version` first. If the command is not found, GenOffice writes the launcher directory to `~/.genoffice/launcher` (Windows: `%USERPROFILE%\.genoffice\launcher`) on every start: read that one line and run `"<dir>/genoffice"` from bash or `"<dir>\genoffice.cmd"` from PowerShell / cmd, quoted, in place of `genoffice` below. Default locations when that file is missing: macOS `/Applications/GenOffice.app/Contents/Resources/cli`, Windows `%LOCALAPPDATA%\Programs\GenOffice\resources\cli`, Linux `/opt/GenOffice/resources/cli`. Every command runs headless on the app's own engines; only `search`, `image` and `media` send data off the machine, to the provider configured in GenOffice.

Always add `--json` when a program reads the output: one JSON object on stdout, `{ "status": "ok", "command", "summary", "output_path"?, "detail"? }` or `{ "status": "error", "code", "message", "detail"? }`. Exit codes: `0` ok, `1` usage or rejected ops, `2` file not found / output exists, `3` conversion failed, `4` app not available.

## Commands

| Command                                                                                                                                                            | What it does                                                                                                                                                                                                                                                                                  |
| ------------------------------------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `genoffice info <file>`                                                                                                                                            | Metadata and structure summary (docx blocks and headings, pptx slides, xlsx sheets, pdf pages)                                                                                                                                                                                                |
| `genoffice convert <file> --to <fmt> [--out <path>] [--force]`                                                                                                     | `pdf→docx/pptx/xlsx`, `csv→xlsx`, `xls/xlsb/ods→xlsx`, `md→docx/html`, `docx→html/md`, `html→docx`, `xlsx→csv` (`--sheet`), and `docx/xlsx/pptx/md/html→pdf`                                                                                                                                  |
| `genoffice create --type pptx --ops <file> --out <path>`                                                                                                           | New deck from a list of slide ops                                                                                                                                                                                                                                                             |
| `genoffice create --type pptx --spec <dir\|file> [--outline <file>] --out <path>`                                                                                  | New deck from page spec files (one per slide, name order) or one deck spec file: px-positioned text, shapes and images (the designed-deck path, see below)                                                                                                                                    |
| `genoffice create --type xlsx --from <data.csv \| table.json> --out <path>`                                                                                        | New workbook from CSV or a 2-D array / `{sheets:[{name,rows}]}`; `"=..."` cells are formulas                                                                                                                                                                                                  |
| `genoffice create --type docx --from <text.md \| fragment.html> --out <path>`                                                                                      | New Word document from Markdown or a restricted-HTML fragment                                                                                                                                                                                                                                 |
| `genoffice create --type pdf --from <document> --out <path>`                                                                                                       | PDF printed by the GenOffice renderer from any md/html/docx/xlsx/pptx file                                                                                                                                                                                                                    |
| `genoffice slides read <pptx> [--slide n] [--full]` / `genoffice slides apply <pptx> --ops <file> [--dry-run]`                                                     | Inspect ids, geometry and text previews (`--full`: whole text, every table row, speaker notes), then edit with ops                                                                                                                                                                            |
| `genoffice slides audit <pptx> [--slide n]` / `genoffice slides render <pptx> --out <dir> [--scale 2]`                                                             | Geometry audit (out of bounds, text overflow, overlap) with op-targetable ids; one PNG per slide to look at                                                                                                                                                                                   |
| `genoffice render <file> --out <dir> [--page n] [--scale 2]`                                                                                                       | One PNG per page of a docx, xlsx, pptx, md, html or pdf file, laid out by the GenOffice renderer: look at it before reporting a document or workbook as done                                                                                                                                  |
| `genoffice slides check <outline.json\|page.json>` / `genoffice slides replace <pptx> --slide n --spec <page.json>`                                                | Validate a deck outline (exit 1 on errors) or build-and-audit one page file and check it against its outline entry and style.md; rebuild one slide from its page file, other slides untouched                                                                                                 |
| `genoffice sheet read <xlsx> [--sheet name] [--range A1:D20] [--formats]` / `genoffice sheet apply <xlsx> --ops <file> \| --cells <file>`                          | Read a range plus `features` (panes, filter, charts, tables, rule counts for the sheet; merges and links in the range; `--formats` adds cell formats, widths, heights); edit with the workbook DSL or by cell address                                                                         |
| `genoffice docs read <docx> [--range a-b] [--html] [--full] [--comments] [--revisions] [--header-footer]` / `genoffice docs apply <docx> --ops <file> [--dry-run]` | Read blocks by index (`--full`: whole block text instead of a 200-character preview; plus comment threads, tracked changes, header/footer text), then edit with ops, restricted HTML, `insert_image`, `insert_chart` / `edit_chart`, `set_header_footer`, `reply_comment` / `resolve_comment` |
| `genoffice guide slides [group\|design\|spec] \| docs \| sheets`                                                                                                   | The op reference for each domain (read this before writing ops); `slides design` / `slides spec` for building a deck                                                                                                                                                                          |
| `genoffice search <query> [--images] [--max n]`                                                                                                                    | Web / image search via the provider configured in GenOffice; `detail.results[]` `{title,url,snippet}`                                                                                                                                                                                         |
| `genoffice image <prompt> [--out f.png] [--aspect 16:9] [--ref a.png,b.png]`                                                                                       | Generate an image with the configured provider and save it; `detail.mime`, `output_path`                                                                                                                                                                                                      |
| `genoffice media <file\|url> [--ask <question>]`                                                                                                                   | Describe / question an image, video or audio file; answer in `summary` and `detail.text`                                                                                                                                                                                                      |
| `genoffice capabilities`                                                                                                                                           | Which cloud features are configured in GenOffice (search, image search, image generation, media analysis) and whether the app is installed; no network call                                                                                                                                   |
| `genoffice open <file>`                                                                                                                                            | Open the file in the GenOffice app; only when the user asks to see it there                                                                                                                                                                                                                   |

`--ops -` and `--cells -` read stdin. `--out` defaults to the input file's folder with the new extension (`convert`) or edits in place (`apply`). Unknown options are rejected (exit 1).

## Behaviour to know

- **Word (`docs`)**: block indexes inside one batch are live — after `insert_content` adds two blocks, later indexes shift by two. A target that matches no block fails the whole batch (exit 1). `findReplace` does not reach table cells (use `replace_blocks` on the table). Tags outside the restricted set are stripped silently. `convert docx→md` writes GFM: `& < >` become entities, formulas become `$…$`, images are dropped and counted in `detail.skipped`. `create --from x.md` embeds local images referenced by relative path. `insert_image` takes a local path (relative to the current directory, then to the ops file), a data: URL or an http(s) URL (png/jpg/gif) and appends unless `afterBlockIndex` is given; `edit_chart` targets blocks `docs read` lists with kind `chart`; `set_header_footer` writes plain text with `{PAGE}` / `{NUMPAGES}` tokens, `view: "first"` switches the different-first-page setting on; comment ids come from `docs read --comments`.
- **Excel (`sheet`)**: ops run in order, each seeing the previous result; structural ops (rows/cols/sheets) still need their own batch, and defined-name ops cannot share one with sheet or row/column ops. Beyond cells, `sheet apply --ops` writes charts (`add_chart`, `edit_chart` on `xl/charts/chartN.xml`), images and shapes, Excel tables, freeze panes, hidden rows/columns, filters with checked values, conditional formats, data validation, hyperlinks, notes, page setup, sheet protection, defined names and tab order/visibility (`genoffice guide sheets` has every field). `--cells` `style` takes the same object as `format_range.format` (`{bold, fillColor, numberFormat, …}`). `sheet read` returns raw values (0.25, serial dates) and formulas separately, `features` (sheet-wide panes, filter, charts, tables and rule counts; merges and links of the range), and with `--formats` the styled cells, column widths (px) and row heights (pt); `convert --to csv` returns the displayed text (percent, dates), UTF-8 with BOM, CRLF line ends, active sheet unless `--sheet`.
- **PowerPoint (`slides`)**: picture ops (`addPicture`, `replacePicture`, `addMedia`, `setImageFill`, `setBackground`) take a local path or data URL in `bytes`; the extension is inferred. `setText` runs inherit the element's first run style unless a field is set explicitly (`bold: false`). Ops in one `apply` batch run in order, so a later op may target a slide an earlier op added; atomic still writes nothing unless every op applied. `--isolation per_op` returns `status: "ok"` with the failed ops in `detail.failures` — check it. Created elements get new `e_*` ids: run `slides read` again before targeting them. `convert pptx→pdf` is rendered like the app's export: one bitmap per slide, no text layer.
- **Conversions in the app** (`→pdf`, `docx→html`, `html→docx`, `create --type pdf`, `render`): a hidden GenOffice process per call, 1–6 s, up to 180 s timeout, safe to run concurrently. `html→docx` is html2docx for real web pages; a restricted-HTML fragment with `<formula>` goes through `create --type docx --from fragment.html` instead.
- **Cloud (`search`, `image`, `media`)**: `--images` returns `detail.images[] {title,imageUrl,sourceUrl,source,width,height}`; providers may return fewer than `--max`. `image` saves with the provider's real format (`output_path` tells you; a `.png` request may come back `.jpg`, noted in `detail.note`); `--aspect` / `--size` are validated locally.

## Workflow

0. When unsure of the command surface, `genoffice help` (or `genoffice <command> --help`) prints the current one; this skill may be older than the installed GenOffice.
1. `genoffice info` or `genoffice <domain> read --json` to learn the structure: block indexes (docx), `s_<n>` / `e_*` ids and EMU geometry (pptx), sheet names and cell values (xlsx).
2. `genoffice guide <domain>` for the op fields, then write the ops JSON yourself.
3. `genoffice <domain> apply --dry-run` when the batch is large; a rejected op comes back with the executor's guided error and usage line — fix that op and resend the whole batch. Atomic: on rejection the file is untouched.
4. Report the output path. Do not run `genoffice open` unless the user asks to view the file in GenOffice: an open tab makes later `apply` calls refuse to write (see below).

## Reading the user's files

Source material the user hands you (a report to turn into slides, a workbook to summarize) is read by you, not by genoffice: md, txt, csv, pdf and images you open directly. Office files are zipped XML, so go through the CLI first:

- docx: `genoffice convert file.docx --to md --out file.md` for the whole text (images are dropped), or `genoffice docs read file.docx --full --json` / `--html` for blocks with indexes.
- xlsx: `genoffice convert file.xlsx --to csv --sheet <name> --out file.csv` for a whole sheet as displayed text, or `genoffice sheet read file.xlsx --range A1:H200 --json` for raw values (without `--range` it stops at 500 rows and says `truncated: true`).
- pptx: `genoffice slides read file.pptx --full --json` for every slide's text, tables and speaker notes; `genoffice slides render file.pptx --out shots` when you need to see the pages.

Take figures and wording from that material; never `cat` or `unzip` the Office file itself.

## Designing a deck

Whenever the user wants a presentation, however short, do not hand-place ops: follow the staged pipeline the GenOffice app itself uses. You are the model, genoffice is the engine, and each stage is a file the CLI checks before the next one starts. Ops (`create --ops`, `slides apply`) are for editing a deck that already exists.

1. `genoffice guide slides design` once per session, then `genoffice guide slides spec`. Run `genoffice capabilities --json` once: it says whether search, image search and image generation are configured in GenOffice; use them for facts and photos only when they are, and plan photo-free pages otherwise.
2. Make a folder for the deck. Write the style sheet to `deck/style.md` (colors, fonts, layout variants, one sentence of style). Then the outline to `deck/outline.json` (core hook + one entry per page: title, type, layout, brief with real figures, image_queries) and run `genoffice slides check deck/outline.json --json`; fix every error before going on. Show the outline to the user only if they asked to review it.
3. One page per step: re-read `deck/style.md` and the page's outline entry, resolve its photos, write `deck/pages/NN.json` (one page object echoing the entry's `title`, `type`, `layout`), run `genoffice slides check deck/pages/NN.json --json`; it finds `outline.json` and `style.md` one folder up and checks the page against entry N and the palette. Fix until `detail.audit`, `detail.outline.findings` and `detail.style.offPalette` are empty (a disagreement with the outline exits 1), then the next page. Never write several pages in one file or one step.
4. `genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/deck.pptx --json` (refuses to run while an outline page has no file or a page disagrees with its entry). Then `genoffice slides render deck/deck.pptx --out deck/shots --json` and look at the PNGs; `genoffice slides audit deck/deck.pptx --json` for the geometry findings. Fix a page in its file and `genoffice slides replace deck/deck.pptx --slide n --spec deck/pages/NN.json --json`. At most two fix rounds.
5. Report the path. Photos come from `genoffice search --images` results, `genoffice image`, or local files the user gave you; never draw a grey box in place of a photo. Figures come from the user's material or `genoffice search`; say so when a number is illustrative.

## Editing an existing deck

A rewrite of one slide, a new title, a swapped picture: use ops, not the staged pipeline, and keep the deck's own design.

1. `genoffice slides read deck.pptx --slide n --full --json` for the element ids and the current text; `genoffice slides render deck.pptx --slide n --out shots --json` when the words alone do not tell you what the page looks like.
2. `setText` on the existing element (runs you leave unstyled inherit the original font, size and color), `addElement` / `addPicture` for new content, `deleteElement` for what goes. Write with `--out` when the user wants the original kept.
3. `genoffice slides audit deck.pptx --slide n --json` after the edit: a longer sentence in the same box overflows; then render once more and look.
4. Ids of edited elements change: run `slides read` again before a second round of ops. `slides replace --slide n --spec` rebuilds the page from a spec and is for decks the staged pipeline built, not for restyling a deck the user brought.

## Examples

Deck from ops (an existing deck's edits, or a fixture; a presentation for a person goes through the staged path above):

```json
[
  {
    "op": "addElement",
    "target": { "slide": 0 },
    "kind": "textbox",
    "offset": { "x": 914400, "y": 685800, "cx": 7315200, "cy": 914400 },
    "paragraphs": [{ "runs": [{ "text": "Q3 Review", "bold": true, "fontSize": 36 }] }]
  },
  { "op": "addBlankSlide", "target": { "slide": 0 } },
  {
    "op": "addTable",
    "target": { "slide": 1 },
    "rows": 3,
    "cols": 3,
    "offset": { "x": 914400, "y": 1828800, "cx": 7315200, "cy": 2286000 }
  }
]
```

```
genoffice create --type pptx --ops deck.json --out q3.pptx --json
```

Workbook with formulas, then a formatted header:

```
echo '[["item","qty","price"],["Apple",2,1.5],["Total","=SUM(B2:B2)",""]]' > table.json
genoffice create --type xlsx --from table.json --out sales.xlsx
echo '[{"op":"format_range","range":"A1:C1","format":{"bold":true,"fillColor":"#FFFF00"}}]' | genoffice sheet apply sales.xlsx --ops - --json
```

Word from Markdown, then insert a paragraph after block 0:

```
genoffice create --type docx --from report.md --out report.docx
echo '[{"op":"insert_content","afterBlockIndex":0,"html":"<p>Executive summary.</p>"}]' | genoffice docs apply report.docx --ops -
```

## Path rules

- Pass absolute paths or paths relative to the current directory; genoffice never writes outside the `--out` path or the file being edited.
- `create` and `convert` refuse to overwrite an existing output without `--force`. `apply` edits in place unless `--out` is given; `--out` onto another existing file also needs `--force`.
- genoffice refuses to rewrite a file the GenOffice window currently has open (exit 2, message "GenOffice has this file open"). Ask the user to close that tab, write to another path with `--out`, or pass `--force` only if they accept that the editor may overwrite the change.
- If `GENOFFICE_ALLOWED_ROOTS` is set in the environment, genoffice only reads and writes inside those directories (exit 2 otherwise); stay within them. Every executed command is logged to `~/.genoffice/cli-audit.jsonl` (`GENOFFICE_AUDIT_LOG` moves or disables it). `GENOFFICE_USER_DATA` points genoffice at a non-default GenOffice profile (open-documents registry, AI settings).
- Image ops (`addPicture`, `setImageFill`, `addMedia`) accept a local file path in their `bytes` field; genoffice reads it.

## Limits

- Conversions to PDF run inside the GenOffice app (hidden, a few seconds each); they need GenOffice installed, or `GENOFFICE_APP_BIN` pointing at it; the same holds for `docx→html` and `html→docx`.
- Excel ops that need the live editor (pivots, sparklines, edits to session-created tables and shapes, `convert_to_values`) are refused with the reason; adding a conditional format or data validation to a sheet that already has rules is refused too (`clear_conditional_formats` first, or use the app).
- Formulas are evaluated with the workbook engine after writing, so `sheet read` and cached-value readers see numbers.
