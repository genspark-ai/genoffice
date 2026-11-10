# HTML → DOCX Conversion Logic and Architecture

This document describes the conversion logic implemented in `packages/html2docx`. It is not a strict attribute-by-attribute translator from the HTML spec to the OOXML spec, but a hybrid converter based on "real browser rendering + document intent recognition + native Word reconstruction".

## 1. Goals and basic principles

The goals of the converter are:

1. Run on Linux or macOS without depending on Microsoft Word.
2. Keep body text, headings, lists, tables, and forms editable as much as possible.
3. Keep layout and visual effects as close as possible to the HTML rendered in Chromium.
4. Use local screenshots for effects that Word is not good at expressing, rather than rasterizing the whole page.
5. Keep the converter independent and not modify existing business code.

The most important design principle currently is "structure first, visual fallback second":

- Content that can be reliably mapped to Word is generated as native OOXML.
- Gradients, complex SVGs, badges, and pure decoration that cannot be reliably mapped are rendered as local images.
- Page backgrounds and absolute-positioned decorations are generated as floating images that do not enter the body text flow.
- No PDF intermediate format is used, to avoid losing semantics, text structure, and editability.

## 2. Overall pipeline

```text
HTML file or URL
    │
    ▼
Headless Chromium loads the page
    │
    ├─ Wait for network resources and web fonts
    ├─ Detect the HTML's design width
    └─ Re-render at the design width
    │
    ▼
Read the DOM, computed style, and real geometric dimensions
    │
    ▼
Recognize document structure, generate an intermediate representation (IR)
    │
    ├─ Editable structures: paragraphs, headings, lists, tables, cards, forms
    └─ Image fallback: gradients, SVG, badges, page decoration, backgrounds
    │
    ▼
Collect DOM elements that need screenshots according to the IR
    │
    ▼
Use the docx library to generate native OOXML
    │
    ▼
DOCX
    │
    ▼
Microsoft Word exports PDF/PNG, compared against HTML screenshots for regression
```

Code responsibilities:

- `tools/cli.ts`: command-line entry point (Playwright driver over a local Chrome).
- `src/convert.ts`: orchestrates browser loading, IR extraction, screenshot collection, and the generation flow.
- `src/extract.ts`: a lightweight assembly entry point for the browser extraction functions.
- `src/browser/core/`: shared-scope extraction modules for style, text, lists, layout, classifiers, etc.
- `src/browser/`: independent browser modules for tables, pages, form media, and visual fallback.
- `src/generate.ts`: a lightweight orchestration entry point for the DOCX generation flow.
- `src/generate/`: modules for page setup, rendering context, Word styles, streaming nodes, and table rendering.

Because the browser extraction layer needs to be serialized by `page.evaluate` into a single self-contained function, `src/browser/core/` uses function-body fragments split by responsibility, which `extract.ts` assembles on the Node side in a fixed order. Peripheral, independent capabilities are injected via `page.addScriptTag`, rather than stuffing every implementation back into a single file.

## 3. Browser rendering stage

### 3.1 Chrome selection

`convert.ts` looks for a browser in the following order:

1. The `CHROME_PATH` environment variable
2. Google Chrome on macOS
3. Common Google Chrome / Chromium paths on Linux

The package talks to the browser through the `BrowserDriver` interface (`src/driver.ts`); the CLI and tests use the Playwright implementation over a locally installed Chrome, the app uses Electron's own webContents. Nothing downloads Chromium. When deploying to Linux, Chromium must be installed in the image, or `CHROME_PATH` must be set.

### 3.2 Initial A4 viewport

The initial viewport is:

```text
794 × 1123 CSS px
```

This is the approximate size of A4 at 96 DPI. `deviceScaleFactor` is set to 2 to improve the clarity of local screenshots.

### 3.3 Design-width detection

Many AI-generated documents use:

```css
body {
  max-width: 880px;
}
```

If extraction is done directly in a 794px-wide viewport, extra line wrapping occurs, and table row heights are also enlarged. The current logic reads the computed `max-width` of `body`:

- If the design width is greater than the initial A4 viewport;
- and does not exceed 1600px;
- then the page is reloaded using a design viewport no smaller than 1024px.

This way, the line wrapping and line height extracted reflect the author's designed size, rather than the result of being squeezed by an overly narrow viewport beforehand.

If a flow document is roughly 0.96-1.22 page-heights tall in the browser (converted to an A4 ratio based on the design canvas width), it is treated as having single-page design intent, and `pageFitSqueeze` applies a light, whole-document compression to geometry and fonts, to avoid Word's slightly looser layout pushing the last block (an invoice stamp, a footer note) onto a second page. Table-heavy documents do not get the benefit of "Word layout being denser," so above 1.05 pages, no compression discount is applied. When generating OOXML, native font size, spacing, padding, line height, and borders are then scaled by `target page width / design viewport width`; local screenshots are scaled proportionally to the target content-area width.

### 3.4 Waiting for resources

While the page loads:

- it prefers to wait for `networkidle0`;
- after a timeout it proceeds with the conversion anyway, to avoid external resources blocking the whole task;
- it waits for `document.fonts.ready`;
- then waits for a brief settling period to let layout and fonts finish reflowing.

Inaccessible images are identified as broken during extraction and discarded.

## 4. Extraction layer: DOM and computed style

The `extractIR` function assembled by `extract.ts` is injected into the browser page and executed there, so it can directly access:

- DOM nodes;
- `getComputedStyle`;
- `getBoundingClientRect`;
- `::before` / `::after` pseudo-elements;
- image loading state;
- the real coordinates of text nodes.

The extracted content falls into four categories.

### 4.1 Text runs

Text is not handled crudely by DOM tag alone; instead it is converted into styled runs. Currently supported:

- font family and size;
- bold, italic, underline, strikethrough;
- font color;
- inline background color;
- letter spacing;
- superscript, subscript;
- `text-transform: uppercase`;
- hyperlinks;
- line breaks (including author line breaks split by line inside `white-space: pre-wrap/pre-line` text);
- small inline images;
- column positions and tab stops produced by wide spaces.

Consecutive text nodes and inline elements are merged into a single paragraph, so that not every `<span>` produces a separate paragraph.

### 4.2 Paragraph styling

Paragraph-level styling includes:

- left, center, right, and justified alignment;
- spacing above/below;
- line-height ratio;
- left indent;
- background color;
- left and bottom borders;
- the exact height of single-line color bars;
- CJK and Latin text markers.

The converter distinguishes between Word font leading for English and CJK:

- Latin line spacing uses a smaller compensation;
- CJK line spacing uses a larger compensation;
- single-line background heading bars use an exact line height, to prevent Word from stretching the color bar taller.

### 4.3 Geometry information

Structural nodes store the browser's measured dimensions, for example:

- table column widths;
- table row heights;
- flex/grid child widths;
- card padding;
- block margins;
- image width/height;
- absolute-positioned element coordinates;
- page content-area width/height;
- body's four-sided margins.

This data is used to suppress Word's AutoFit and default paragraph leading.

### 4.4 Page setup

After extraction, a `docsettings` node is generated, which stores body's actual four-sided margins:

```js
{
  type: 'docsettings',
  marginsPx: {
    top,
    bottom,
    left,
    right,
  },
}
```

The DOCX no longer uniformly uses a fixed 2cm top/bottom margin; instead it tries to use the HTML's actual page padding. This resolves the issue of "the content size is correct, but the last form box is pushed to a second page by the fixed page margin."

## 5. Intent-recognition layer: DOM → IR

The IR is an intermediate representation between browser structure and Word structure. It proactively discards unimportant HTML implementation details and retains only the document intent needed to generate Word output.

Current main node types:

| IR type       | Meaning                                     | Word output                                     |
| ------------- | ------------------------------------------- | ----------------------------------------------- |
| `docsettings` | Page margins                                | Section page margin                             |
| `heading`     | Heading                                     | Heading Paragraph                               |
| `para`        | Ordinary paragraph                          | Paragraph + TextRun                             |
| `list`        | Ordered/unordered list                      | Native Word numbering                           |
| `table`       | HTML/CSS Grid table                         | Fixed-layout Word Table                         |
| `card`        | Card with background, border, and padding   | Single-cell Table                               |
| `kpirow`      | flex/grid horizontal layout                 | Single-row, multi-column table                  |
| `formfield`   | Input box, textarea, select box             | Underlined paragraph or single-cell table       |
| `code`        | Multi-line monospace text                   | Monospace-shaded paragraph                      |
| `image`       | Ordinary image or local screenshot          | Inline ImageRun                                 |
| `floatimg`    | Page decoration, signature, footer          | Floating ImageRun                               |
| `pagebg`      | Page background                             | Floating image/background color behind the page |
| `pagebreak`   | Forced page break                           | PageBreak                                       |
| `hr`          | Horizontal rule or pseudo-element color bar | Paragraph bottom border                         |
| `spacer`      | Large vertical blank space                  | Empty paragraph with exact height               |

### 5.1 Classification order

The classification order matters a great deal. `processElement` currently handles things roughly in the following priority:

1. Skip invisible elements and irrelevant tags such as `script/style`.
2. Recognize large signatures at the bottom of the page.
3. Recognize absolute-positioned or fixed decoration.
4. Extract page breaks and large top-of-page blank space.
5. Handle native semantic tags:
   - headings;
   - HTML tables;
   - lists;
   - code;
   - HR;
   - images;
   - forms.
6. Handle SVG, canvas, checkbox/radio.
7. Recognize badges, pills, decorative leaves, and gradient containers.
8. Recognize CSS Grid tables.
9. Recognize page-level multi-column layout.
10. Recognize multi-row, structurally uniform grid/flex layouts.
11. Recognize single-row flex/grid layouts.
12. Recognize cards.
13. Fall back to an ordinary paragraph and recurse into child nodes.

High-fidelity rules must come before generic block-level recursion; otherwise a complex component would be split into ordinary paragraphs prematurely.

## 6. Main structure-recognition logic

### 6.1 HTML tables

A native `<table>` has the following extracted:

- measured width of each column;
- measured height of each row;
- colspan;
- cell padding;
- borders on all four sides;
- background color;
- horizontal and vertical alignment;
- header rows;
- paragraphs or nested structures inside cells.

When generating:

- a fixed table layout is used;
- table grid column widths are written;
- cell widths are written at the same time;
- adjacent grid widths are merged according to colspan;
- row heights use the browser's measured values;
- `HeightRule.EXACT` is used when strict fidelity is required.

Setting both the table grid and cell widths at the same time is meant to stop Word's AutoFit from redistributing column widths based on text length.

### 6.2 CSS Grid tables

If a container contains multiple rows of a structurally identical Grid:

- the number of visible child items is the same in each row;
- the number of columns is within a reasonable range;
- the widths of each row are close to each other;
- there are enough rows or a clear outer border;

then it is converted into a `table`, rather than a set of independent text blocks.

This is used for questionnaires, class schedules, survey forms, and other templates that don't use `<table>`.

### 6.3 Flex/Grid horizontal rows

Side-by-side KPIs, personal-info bars, tag rows, page headers, etc. are converted into `kpirow`:

- column widths are computed from the real bounds of the child elements and the midpoints of the gaps between them;
- the left/center/right position of each child element within its corresponding column is detected;
- the parent container's row height and padding are stored;
- on the Word side, a fixed-layout, single-row, multi-column table is generated.

The default cap on child-element height is 500px (taller than that is treated as a page-level two-column layout); however, tall card rows (pricing plans) with ≥3 columns of similar width are relaxed to 950px — a Word page can still fit one such row, preserving the side-by-side structure.

If there is an absolute-positioned color block on the left side of the container:

- the color block becomes the first column on its own;
- the body content becomes the second column;
- the color block's width uses the browser's measured value;
- the body's left indent, which was originally relative to the whole container, is corrected.

This is how the header color block on the Service Evaluation page is handled.

### 6.4 Cards

Blocks with a background, border, left accent line, or noticeable padding are recognized as a `card`.

Cards are output as single-cell Word tables, preserving:

- background color;
- full border;
- left or top accent border;
- padding on all four sides;
- top/bottom margin;
- editable inner content.

A table rather than paragraph shading is used because Word's paragraph shading cannot reliably express four-sided padding and a complete card border.

**Flattening**: a light-colored top-level card (a full-page section) that is taller than half a page and allowed to break across pages cannot be expressed as a single-row table — Word nested tables cannot break across pages — so it is flattened into a sequential paragraph flow. During flattening, the card's shading is propagated to all child blocks: paragraph shading for para/heading, spacer paragraphs, list items, kpirow cells and gaps, and table cells; the spacing before/after each child block itself is marked with `ambientShading` for the surrounding section's background color, to avoid exposing the page background and creating stripes.

### 6.5 Page-level two columns

Page-level two-column layouts such as resumes are output as:

- a single row, two columns;
- a borderless Word table;
- each column placed in full inside its corresponding cell;
- table rows are allowed to break across pages;
- column background and inner padding are preserved.

The page is not sliced into a large number of small rows by Y coordinate, because that approach is prone to left/right column misalignment once Word paginates.

Paragraph and card spacing within columns is moderately compressed, to reduce extra pages caused by differences between the HTML and Word pagination engines.

### 6.6 Forms

Currently, real Word Content Controls are not generated; instead, visually editable forms are generated:

- input boxes with only a bottom border → a paragraph with a bottom border;
- boxed inputs, textareas, selects → a single-cell table;
- width/height, border, background, padding, and existing text are preserved;
- checkbox/radio → Unicode box or circle characters.

This approach is highly compatible but lacks the validation and field properties of real Word form controls.

### 6.7 Lists

Two sources are supported:

1. native `<ul>` / `<ol>`;
2. pseudo-lists exported by Word, e.g. Symbol-font characters such as "·", "•", "▪".

For pseudo-lists, the fake bullet is stripped from the text and converted into native Word numbering, avoiding duplicate dots.

### 6.8 Horizontal rules and pseudo-elements

`<hr>` and qualifying `::before` / `::after` bars have the following extracted:

- width proportion;
- height;
- color;
- left-aligned or centered.

On the Word side, these are rebuilt using an indented paragraph bottom border.

## 7. Image fallback strategy

### 7.1 Why not screenshot everything

A full-page screenshot achieves the highest static visual consistency, but leads to:

- text that cannot be edited;
- inability to search and copy;
- large file size;
- inability to paginate normally;
- unsuitability as a DOCX.

For this reason, only the local areas that Word struggles to express are currently screenshotted.

### 7.2 Content that gets screenshotted

Mainly includes:

- ordinary `<img>`;
- SVG and canvas;
- gradient-background containers;
- rounded pills and complex badges;
- small decorative shapes;
- icon-font elements with no text, whose glyph lives in `::before/::after`;
- complex badge grids;
- page background images;
- absolute-positioned decoration;
- oversized bottom signatures.

Elements to be screenshotted are given a unique `data-h2d-id`. `convert.ts` walks the IR to collect these IDs, then screenshots the corresponding DOM elements via Puppeteer.

### 7.3 Page background

The page background is copied to an offscreen backdrop:

- background color, image, size, position, and repeat mode are preserved;
- body/html backgrounds are supported;
- simple `::before` background overlays are supported.

While screenshotting, the other body child elements are temporarily hidden and only the background is kept, to avoid the body content being captured twice into the background image.

When generating the DOCX, the background image is placed behind the page as a behind-document floating image.

### 7.4 Absolute-positioned decoration

Absolute-positioned or fixed elements do not enter the body flow. Elements that satisfy the size and page-range limits are:

- screenshotted locally;
- stored with coordinates relative to the body content area;
- stored with the content area's width/height;
- scaled proportionally to the content area in Word;
- positioned as a floating image in the page-margin coordinate system.

Absolute-positioned elements that don't fall within the safe range are discarded, to avoid incorrect decoration covering the body content.

**In-ancestor decoration (`inParent`)**: if a decoration falls entirely within the nearest positioned ancestor box (a leaf icon to the right of a banner, a watermark numeral in the top-left corner of a section), the classifier additionally records its coordinates and the ancestor's width/height relative to that ancestor. Note that the positioned ancestor must be found by manually walking the parent chain — SVG elements have no `offsetParent`. On the generation side, when that ancestor is rendered as a card, the decoration is anchored to the card cell (`relativeFrom: column/paragraph`) and drawn on top of the shading (an underlying image would otherwise be covered by the cell shading); when the card is flattened, it is anchored to the zero-height anchor paragraph at the start of the flattened output, and the top blank space is only compressed down to a height that can still accommodate the decoration. This kind of decoration is not subject to the "first page only" restriction; when there is no card context to fall back on and it falls beyond the page-anchored path, decoration beyond the first page is discarded by the generation side (consistent with prior behavior).

## 8. DOCX generation logic

### 8.1 Units

Under the 794px A4 viewport, the base conversion relationships are:

```text
1 CSS px ≈ 15 twips
1 CSS px ≈ 1.5 half-points
1 CSS px ≈ 9525 EMU
```

Used respectively for:

- paragraphs, margins, table width/height;
- font size;
- floating-image coordinates.

When the HTML uses a wider design viewport, the above text and box-model sizes are also multiplied by the page scale ratio. Fonts retain a small amount of width compensation, because Word's Arial is usually narrower than Inter/Roboto as rendered in the browser; without this, the original line wrapping would be lost.

### 8.2 Font mapping

Word and browsers have different available fonts, so the generator includes a font-family mapping:

- monospace → Consolas;
- common sans-serif / Inter → Arial;
- common serif → Times New Roman;
- DM Serif / Playfair / Cormorant → Georgia;
- Archivo / Bebas / Oswald → Arial Narrow;
- script/cursive → Segoe Script;
- CJK sans-serif → Microsoft YaHei;
- CJK serif → SimSun.

If a Latin-font run actually contains CJK characters, it is switched to the corresponding CJK font, to avoid arbitrary Word fallback.

### 8.3 Paragraphs

Paragraph generation handles:

- TextRun styling;
- hyperlinks;
- tab stops;
- alignment;
- indentation;
- spacing above/below;
- line spacing;
- background and border;
- native bullets.

Headings enable `keepNext` and `keepLines`, to reduce cases where a heading lands at the bottom of a page and the body text gets pushed to the next page.

### 8.4 Tables

Word tables are the core of the current layout system, used not only for HTML tables but also for:

- flex horizontal rows;
- cards;
- page two-column layouts;
- form boxes;
- KPIs;
- color-block titles.

This makes it possible to take advantage of Word's:

- fixed column widths;
- cell shading;
- cell margins;
- row height;
- vertical alignment;
- row pagination control.

### 8.5 Spacing

Ordinary paragraphs use Word paragraph spacing.

A table itself cannot reliably carry block-level margin above/below, so a spacer paragraph with an exact height is inserted before and after the table. Small spacing is not forced to write a visible character; larger spacing uses a zero-width character to prevent Word from collapsing it.

The current spacing priority is:

1. HTML computed margin;
2. HTML padding;
3. the browser's measured line height;
4. Word's minimum compatible default value.

Uniformly compressing all table padding to force a single page is no longer done, because it makes tables, cards, and body content overly cramped as a whole.

### 8.6 Pagination

The current pagination strategy:

- CSS `break-before: page` → an explicit page break;
- headings try to stay with the following paragraph;
- ordinary data-table rows are not split by default;
- page-level two-column tables are allowed to break across pages;
- KPI horizontal rows are kept whole;
- large cards are preferentially moved to the next page as a whole;
- a signature at the bottom of the page can be converted into a floating image fixed at the bottom.

Word and Chromium use different line-breaking algorithms, so pagination can only be approximated, not made mathematically identical.

## 9. Core lessons from current quality fixes

### 9.1 Color blocks becoming narrower

Cause:

- setting only the Table Grid width is not enough;
- Word's AutoFit keeps squeezing blank columns narrower.

Fix:

- use a fixed table layout;
- set both the grid column width and cell width at the same time;
- sum the actual spanned width for colspan.

### 9.2 Tables becoming overall too cramped

Cause:

- cell padding was once uniformly compressed in order to keep everything on a single page;
- table rows did not use the browser's measured height.

Fix:

- restore the computed padding;
- preserve the browser's row height;
- use an exact row height when strict fidelity is required;
- prioritize fixing the page margins instead of compressing the body content.

### 9.3 Incorrect line spacing

Cause:

- Word's built-in font leading;
- Latin and CJK have different leading;
- uniformly using the CJK compensation makes English text too cramped.

Fix:

- detect CJK per paragraph;
- use different line-height compensation for Latin and CJK;
- use an exact line height for single-line color blocks.

### 9.4 Content pushed to a second page

Cause:

- the DOCX fixedly used roughly a 2cm top/bottom page margin;
- the HTML actually might have only 32px of page padding.

Fix:

- extract the actual four-sided margins from body;
- write them into the DOCX section margin;
- preserve the content size instead of solving pagination by compressing tables.

### 9.5 A narrow viewport produces extra line wrapping

Cause:

- the template is designed at 880px;
- extracting directly in a 794px A4 viewport causes premature line wrapping.

Fix:

- detect body's `max-width`;
- re-render at the author's designed width;
- then map the structure onto an A4 Word page.

## 10. Verification process

Final quality must be judged against Microsoft Word:

1. Chrome captures a screenshot of the source HTML.
2. Run the converter to generate the DOCX.
3. Open the DOCX with Microsoft Word.
4. Word exports a PDF.
5. Convert the PDF to PNG.
6. Generate an HTML/Word side-by-side comparison image per page.
7. Check:
   - number of document pages;
   - page margins;
   - font, font size, and line spacing;
   - color block size;
   - table column width and row height;
   - card padding and inter-block spacing;
   - image size and position;
   - whether the content is still editable.

LibreOffice can quickly check whether a file is corrupted, but it cannot serve as the visual regression baseline, because its pagination, font, table, and background rules differ from Word's.

Regression baselines (source HTML, converted `.docx`, the PDF Word exports from it, and a side-by-side comparison image per template) are kept outside the repository.

## 11. Known limitations

### 11.1 Word and the browser are not the same layout engine

The following cannot be guaranteed to match exactly:

- automatic line wrapping;
- font fallback;
- font leading;
- pagination;
- the free-form layout of flex/grid;
- margin collapsing.

### 11.2 CSS support is not a complete spec implementation

This is currently not a general-purpose CSS-to-OOXML engine; it focuses on the high-frequency structures found in document templates. Complex selectors themselves are handled by the browser, but the mapping from computed style to Word properties remains a limited set.

### 11.3 Screenshotted areas are not editable

Fallback areas such as gradients, SVG, badges, page decoration, and backgrounds become images.

Native Word tables do not support the CSS `border-radius`. Currently, wide rounded title bars use a local screenshot while keeping the body text editable; the light-colored outer border of a rounded card is preferentially omitted, rather than outputting an incorrect right-angle border. If the rounded outer border must be fully preserved, a decorative background image or a screenshot of the whole card is needed as a further step.

### 11.4 Real Word forms are not yet implemented

Currently, forms are editable text and table appearances, not Content Controls, Legacy Form Fields, or ActiveX controls.

### 11.5 Multi-page backgrounds

Page background images are currently handled with the page-floating-image approach; a complete section/background model for complex, different backgrounds across multiple pages has not yet been formed.

### 11.6 Headers and footers

Currently, absolute-positioned headers/footers are mainly handled as page floating decoration, and have not been converted into native Word Header/Footer Parts.

### 11.7 Nesting depth

When cards and tables are nested too deeply, they are downgraded to ordinary paragraphs, to avoid generating deeply nested tables that Word doesn't support or that are extremely unstable.

## 12. Future direction

Suggested order for continued improvement:

1. Establish automated Word regression testing over a fixed set of templates.
2. Output the IR as JSON, to make it easier to pinpoint whether an issue is in extraction or generation.
3. Add element-level diagnostic logging: DOM selector → IR node → OOXML node.
4. Support native Word Header/Footer.
5. Support real Content Control forms.
6. Add multi-section, landscape pages, and custom paper sizes.
7. Establish font-availability detection and a font-substitution configuration.
8. Establish separate visual-diff thresholds for images, cards, and tables.
9. Consolidate template special cases into configurable rules, to reduce hardcoding.
10. Add timeouts, a browser pool, and resource-size and concurrency limits when running on a server.

## 13. Usage

Command line:

```bash
npx tsx packages/html2docx/tools/cli.ts input.html output.docx --verbose
```

Linux:

```bash
CHROME_PATH=/usr/bin/chromium npx tsx packages/html2docx/tools/cli.ts input.html output.docx
```

As a module: see [README.md](./README.md) — `convertHtmlToDocx({ url }, driver, options)` returns the DOCX bytes together with the IR and stats; the IR can be saved for debugging, rule analysis, or as a test snapshot.

## 14. Empirically observed behavior of the Word (Mac) pagination engine (measured in experiments, 2026-07)

The following conclusions were all obtained by manually constructing variant docx files and rendering/comparing them with the same build of Word; they are the basis for several defensive design choices in the converter — read this section before touching related code.

1. **A forced ~26px inset at the top of the page**: even with `pgMar top=0`, the first line still starts roughly 26px (~405 twips) below the top of the page, and this is unrelated to image height, paragraph-mark font size, or docGrid. Full-page screenshot slicing must therefore reserve a margin for this: the slice budget is 96% of the page height (pages.js), and the rendering fallback is content height − 45px (renderer.js pageScale). Otherwise the bottom ~25px of a slice gets silently cut off (this once caused the last line of a resume's CONTACT card to be lost).
2. **`keepNext` semantics only extend to the next paragraph**: when a heading has `keepNext` plus an ordinary spacer paragraph, Word only guarantees that "heading + spacer" stay on the same page — the body text still gets paginated away → an orphaned heading. Every spacer (and at most one subheading paragraph) between the heading and the body must carry `keepNext`.
3. **A keep-chain fallback drags along the table before it**: when the large block at the tail of a keepNext chain (a `cantSplit` card row) doesn't fit, Word doesn't just move the chain to the next page — it also drags along the table before the chain, even across paragraphs with kn=0 in between (neither empty paragraphs nor ZWSP paragraphs stop this). Therefore the trailing spacer must never carry `keepNext`, otherwise the entire document links into one chain, and Word paginates at an absurd point (an entire page with nothing but a heading).
4. **Body text ending with a table produces a ghost blank page**: when the last block in the body is a `tbl`, Word automatically appends a paragraph mark with the default line height; if the table happens to exactly fill the last page, that mark overflows into an entire extra page with nothing but the background. The generation layer must terminate with a 1px EXACT empty paragraph (page-settings.js renderSection); the same applies to table cells (the docx library likewise appends a default paragraph to a cell that ends with a table).
5. **Screenshots of transparent wrapper elements get embedded at the full box height**: a stretched photo container in a column layout (box height 1591px, with only the top ~300px actually visible) embedded as an inline image occupies space at the full box height, immediately pushing out an extra page. `markForScreenshot` crops elements that have "no own drawing, box height ≥600, and an invisible tail ≥150px" down to the bottom of the visible content.

## 15. Findings from the fourth round of fixes (2026-07-20, 30 samples)

1. **`w:lang w:val="zh-CN"` makes Word break Latin words character-by-character using CJK rules** (e.g. "Report" → "Rep/ort"). Controlled experiments show that the `w:wordWrap` element runs in the opposite direction and doesn't rescue this (val=0 = allow character-level breaking; under CJK lang, val=1 is also ineffective); the only effective fix is to keep docDefaults' `w:lang w:val` fixed at en-US, marking only `eastAsia` as CJK (page-settings.js). Note that the docx library silently drops the wordWrap option for ordinary paragraphs.
2. **Nested fixed-layout tables must have an explicit `tcW`**: Word for Mac collapses a nested fixed-layout table without `tcW` down to the grid's placeholder width (a ~100dxa sliver). `renderCard` now always emits the trio FIXED + columnWidths + cell width; an autofit nested card can get blown out by an unbreakable inner run (an nbsp used to fill a blank line) and pushed out past the right edge of its container, getting clipped.
3. **The centering offset from a wide-viewport reload is not a page margin**: after reloading at 1024px, body's auto-margin centers it; `docsettings` uses `canvasWidth = min(innerWidth, max(pageSizePx, bodyWidth))` as the measurement baseline and strips out the `centerOffset` — otherwise the whole document ends up narrowed by ~20% plus a false ~3cm margin.
4. **`oneLineWidthPx` must re-measure using the font after Word's mapping**: `convert.ts` injects `__h2dWordFontLineWidth` (with FONT_RULES serialized into the page); Cormorant → Georgia is 15-40% wider, and measuring directly with the browser's narrower font causes heading letters to drop to the next line (e.g. "EDUCATIO/N").
5. **`kpirow` widening order**: gap track → spare room in generator.avail (widen the table) → compress sibling columns. Widening the table directly when there's already spare room in the browser's layout is faithful to the source.
6. **Inline merging for deep flatten (`inlineCellPara`)**: a single para / a small checkbox formfield (box mode ≤20px) / a nested kpirow are recursively merged into one line of runs; the checkbox run must carry `text: ''` (`tabStopsFor` assumes `text` exists). This restored horizontal layouts such as "COMBUSTIBLE: ☐ GASOLINA ☐ DIÉSEL ☐ GAS".
7. **List bullet determination looks at the `li`, not the `ul`**: frameworks such as pico.css set `list-style: none` on `ul` but set the marker individually on `li`; `extractList` decides plain based on the `li`'s computed `listStyleType` (tables-lists.js).
8. **Circular badges** (`circularBadge`: border-radius ≥ height/2, width and height similar, no whitespace-only text): `display:flex` can also be screenshotted inline (`inlineBgHex` rejects non-inline, needing `bgHex` as a fallback), and the displayed size is capped at the parent font size × 1.35 — otherwise a 42px badge embedded at its original height raises the line box, pushing an ~800px `cantSplit` card (with Word's ×1.2 factor) over the page limit, producing an isolated hint-box page plus a fully blank page. After capping, 12 pages return to normal. Rectangular text chips still go through the text-run path so they remain searchable.
9. **`isBoxRow` column-count limit only counts child elements wider than 40px**: a "step ▶ step" flow diagram with 7 child elements (4 steps + 3 arrows) is no longer stacked vertically; geometric symbols such as ▶◀◆ (U+25A0-25FF) now go through SYMBOL_RE and render as text glyphs (Arial Unicode MS + FE0E), instead of turning into blue emoji; `align-self: center` child elements are mapped to cell-level `vAlign`.
10. **The hallucination rate of vision-LLM scoring is measured to be extremely high**: in this round, more than half of the "below-100" findings were disproved programmatically (line-merging in a `pdftotext` count, a `document.xml` count, an actual render). Review discipline: for any "missing content" finding, first run a `pdftotext | tr -d '\n' | grep -o` count and compare it against the HTML source before deciding whether to fix it.
