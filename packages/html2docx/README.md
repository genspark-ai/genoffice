# @genoffice/html2docx

HTML → DOCX conversion for the HTML app's "Export as Word": the page is rendered
in a real browser, reduced to a document intent tree (headings, paragraphs,
lists, tables, cards, KPI rows, form fields, page backgrounds…) and written out
as native OOXML with the `docx` library. Only visuals that have no Word
counterpart (charts, icons, decorated boxes) are screenshotted.

See [ARCHITECTURE.md](./ARCHITECTURE.md) for the pipeline, IR and OOXML
mapping.

## Usage

```ts
import { convertHtmlToDocx } from '@genoffice/html2docx'

const { docx, stats } = await convertHtmlToDocx({ url }, driver, {
  onProgress: ({ stage, pct }) => …,   // load | extract | screenshot | generate
  signal,                               // AbortSignal
})
```

`driver` is a `BrowserDriver` (`src/driver.ts`): the host owns the browser and
supplies a page-like adapter. `src/drivers/playwright.ts` implements it over
playwright-core and a locally installed Chrome for tests and the CLI; the app
supplies an Electron `webContents` implementation. The converter never closes
the driver.

CLI (single file, needs Chrome; `CHROME_PATH` overrides the lookup):

```bash
npx tsx packages/html2docx/tools/cli.ts input.html output.docx --verbose
```

## Layout

- `src/browser/` — in-page code, plain JS kept verbatim: `core/*.js` are slices
  of one `extractIR()` function body (assembled in `src/extract.ts` from `?raw`
  imports), the other files install `__html2docx*` helper globals.
- `src/convert.ts` — orchestration: viewport/authored-width detection, lazy
  image warm-up, extractor injection, element screenshots, generation.
- `src/generate/` — IR → OOXML (`docx` library).
- `tests/features.test.ts` — 66 end-to-end cases, each rendering in Chrome.

The generation layer was ported from untyped CommonJS with `strict: false`;
types are tightened file by file without logic changes.
