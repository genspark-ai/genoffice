import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { afterAll, beforeAll, test } from 'vitest'
import type { Browser } from 'playwright-core'

import { convertHtmlToDocx, EXTRACTOR_SOURCE, mapFont } from '../src'
import { launchChrome, PlaywrightDriver } from '../src/drivers/playwright'

const A4 = { width: 794, height: 1123, deviceScaleFactor: 2 }
let browser: Browser

beforeAll(async () => {
  browser = await launchChrome()
})
afterAll(async () => {
  await browser?.close()
})

test('assembles a self-contained browser extractor function', () => {
  assert.match(EXTRACTOR_SOURCE, /^function extractIR\(\) \{/)
  assert.match(EXTRACTOR_SOURCE, /function markForScreenshot/)
  assert.match(EXTRACTOR_SOURCE, /function extractTable/)
  assert.match(EXTRACTOR_SOURCE, /function processElement/)
  assert.match(EXTRACTOR_SOURCE, /__html2docxPages\.build/)
})

async function convertHtml(html: string, name: string) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'html2docx-'))
  const input = path.join(dir, `${name}.html`)
  await fs.writeFile(input, html)
  const driver = await PlaywrightDriver.create(browser, A4)
  let result
  try {
    result = await convertHtmlToDocx({ url: pathToFileURL(input).href }, driver)
  } finally {
    await driver.close()
  }
  const zip = await JSZip.loadAsync(result.docx)
  const xml = async (file: string) => zip.file(file)?.async('string')
  return { zip, xml, ir: result.ir, screenshotText: result.screenshotText }
}

test('preserves rowspan, numbering, controls, headers, media, and page size', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html>
    <html><head><style>
      @page { size: letter landscape; margin: 0.5in; }
      body { max-width: 900px; margin: 0 auto; font-family: Arial; }
      .row { display:flex; gap:24px; }
      .row > div { width:180px; border:1px solid #999; padding:8px; }
      .shadow { width:240px; padding:16px; box-shadow:0 4px 12px rgba(0,0,0,.3); }
    </style></head><body>
      <header data-docx-header><p>Repeated header</p></header>
      <table><tbody>
        <tr><th rowspan="2">Merged</th><th>A</th></tr>
        <tr><td>B</td></tr>
      </tbody></table>
      <ol style="list-style-type:lower-alpha"><li>Alpha</li><li>Beta</li></ol>
      <input name="customer" value="Alex">
      <input type="checkbox" name="approved" checked>
      <select name="status"><option>Draft</option><option selected>Final</option></select>
      <div class="row"><div>Left</div><div>Right</div></div>
      <div class="shadow">Shadow content</div>
      <iframe title="Frame preview" srcdoc="<p>Embedded frame</p>" style="width:240px;height:100px"></iframe>
    </body></html>`,
    'features',
  )
  const documentXml = await xml('word/document.xml')
  const numberingXml = await xml('word/numbering.xml')
  const headerXml = await xml('word/header1.xml')
  assert.match(documentXml, /w:vMerge w:val="restart"/)
  assert.match(documentXml, /w:vMerge w:val="continue"/)
  assert.match(documentXml, /w:sdt/)
  assert.match(documentXml, /w14:checkbox/)
  assert.match(documentXml, /w:dropDownList/)
  assert.match(documentXml, /w:pgSz w:w="15840" w:h="12240"/)
  assert.match(numberingXml, /w:numFmt w:val="lowerLetter"/)
  assert.match(headerXml, /Repeated header/)
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'shadow and iframe should use screenshot media',
  )
})

test('repeats page backgrounds from the header without adding a body page', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; background:#0b0b0b url("data:image/gif;base64,R0lGODlhAQABAIAAAAUEBAAAACwAAAAAAQABAAACAkQBADs=") repeat; color:#f4f1e8; }
      .second { break-before:page; }
    </style></head><body>
      <p>Dark page one</p>
      <p class="second">Dark page two</p>
    </body></html>`,
    'repeated-background',
  )
  const documentXml = await xml('word/document.xml')
  const headerXml = await xml('word/header1.xml')
  assert.match(headerXml, /wp:anchor/)
  assert.match(headerXml, /behindDoc="1"/)
  assert.match(headerXml, /w:line="1" w:lineRule="exact"/)
  assert.match(documentXml, /w:pgMar[^>]*w:header="0"/)
  assert.doesNotMatch(documentXml, /behindDoc="1"/)
  assert.match(documentXml, /Dark page one/)
  assert.match(documentXml, /Dark page two/)
})

test('uses a plain background color for flat solid page backdrops', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; background:#0b0b0b; color:#f4f1e8; }
      .second { break-before:page; }
    </style></head><body>
      <p>Dark page one</p>
      <p class="second">Dark page two</p>
    </body></html>`,
    'flat-color-background',
  )
  const documentXml = await xml('word/document.xml')
  // no backdrop screenshot: the resampled header image's tone visibly
  // differs from same-color shading fills on content blocks
  assert.match(documentXml, /<w:background w:color="0B0B0B"\/>/)
  assert.match(documentXml, /Dark page one/)
  assert.match(documentXml, /Dark page two/)
})

test('uses a zero-height anchor paragraph for page decorations', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; background:#f4f4f4; }
      main, h1, p { margin:0; }
      .curve {
        position:absolute;
        top:0;
        right:0;
        width:120px;
        height:80px;
        background:linear-gradient(135deg,#4a90c2,#8fb8dd);
        border-radius:0 0 0 100%;
      }
    </style></head><body>
      <div class="curve"></div>
      <main><h1>Resume title</h1><p>First content line</p></main>
    </body></html>`,
    'floating-decoration-anchor',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /wp:anchor/)
  assert.match(documentXml, /w:line="1" w:lineRule="exact"/)
  // flat-color page backgrounds use w:background, not a header image
  assert.match(documentXml, /<w:background w:color="F4F4F4"\/>/)
  // a top:0 decoration must not be mistaken for a footer bar on short bodies
  assert.doesNotMatch(documentXml, /<wp:align>bottom<\/wp:align>/)
})

test('rasterizes bounded relative compositions with positioned text overlays', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .hero { position:relative; width:700px; height:300px; background:#111; }
      .hero-copy { position:absolute; inset:auto 30px 20px; color:white; }
    </style></head><body>
      <header class="hero"><div class="hero-copy"><h1>Overlay title</h1><p>Overlay subtitle</p></div></header>
      <p>Editable body</p>
    </body></html>`,
    'positioned-overlay',
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(documentXml, /Overlay title/)
  assert.match(documentXml, /Editable body/)
  assert.ok(
    Object.keys(zip.files).some((name) => name.startsWith('word/media/')),
    'positioned composition should be embedded as screenshot media',
  )
})

test('renders explicit multipage containers as separate page images', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      [data-docx-page] { width:600px; height:800px; overflow:hidden; position:relative; }
      .second { background:#ddeeff; }
      .floating { position:absolute; left:200px; top:300px; transform:rotate(8deg); }
    </style></head><body>
      <section data-docx-page><h1>Page one</h1></section>
      <section data-docx-page class="second"><h1 class="floating">Page two</h1></section>
    </body></html>`,
    'multipage',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /w:br w:type="page"/)
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'each explicit page should produce a page image',
  )
})

test('keeps semantic content between inferred page compositions', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .opener {
        width:700px;
        height:800px;
        break-before:page;
        break-after:page;
        position:relative;
      }
      .opener h1 { position:absolute; left:120px; top:300px; }
      .full-page { height:1119px; }
      .chapter { break-before:page; }
    </style></head><body>
      <section class="opener full-page"><h1>Chapter one opener</h1></section>
      <section class="chapter">
        <p>Chapter one body sentinel</p>
        <ul><li>Chapter one list sentinel</li></ul>
      </section>
      <section class="opener"><h1>Chapter two opener</h1></section>
      <section class="chapter">
        <table><tr><td>Chapter two table sentinel</td></tr></table>
      </section>
    </body></html>`,
    'mixed-page-compositions',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Chapter one body sentinel/)
  assert.match(documentXml, /Chapter one list sentinel/)
  assert.match(documentXml, /Chapter two table sentinel/)
  assert.match(documentXml, /w:pgMar[^>]*w:header="0"[^>]*w:footer="0"/)
  assert.equal(
    (documentXml.match(/w:br w:type="page"/g) || []).length,
    2,
    'a full-page image should advance naturally without creating a blank page',
  )
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'chapter openers should remain page screenshots',
  )
})

test('keeps absolute-positioned content beyond the first page', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { position:relative; width:700px; height:1800px; margin:0 auto; }
      .first { position:absolute; top:80px; left:60px; }
      .second { position:absolute; top:1320px; left:180px; transform:rotate(5deg); }
    </style></head><body>
      <h1 class="first">First page</h1>
      <h1 class="second">Second page absolute content</h1>
    </body></html>`,
    'absolute-pages',
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(
    documentXml,
    /w:br w:type="page"/,
    'full-page image slices should advance naturally without inserting blank pages',
  )
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'both page clips should be embedded',
  )
})

test('aligns a padded card border with surrounding paragraphs', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { max-width:880px; margin:0 auto; padding:32px 80px; box-sizing:border-box; }
      .note { border:1px solid #777; padding:14px 16px; background:#fafafa; }
    </style></head><body>
      <p>Before</p>
      <section class="note"><p>Inside</p></section>
      <p>After</p>
    </body></html>`,
    'card-alignment',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /<w:tblInd w:type="dxa" w:w="60"\/>/)
})

test('does not invent padding around zero-padding card headers', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .card { border:3px solid #c0392b; padding:0; width:700px; }
      .header {
        display:flex;
        align-items:center;
        gap:14px;
        padding:12px 18px;
        background:#c0392b;
        color:white;
      }
      .dice { white-space:nowrap; }
    </style></head><body>
      <div class="card"><div class="header"><span class="dice">🎲 目：6</span><span>追加要望カード｜データ取得ミッション</span></div></div>
    </body></html>`,
    'flush-card-header',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(
    documentXml,
    /<w:tcMar><w:top w:type="dxa" w:w="0"\/><w:left w:type="dxa" w:w="0"\/><w:bottom w:type="dxa" w:w="0"\/><w:right w:type="dxa" w:w="0"\/><\/w:tcMar>/,
  )
})

test('keeps flex icon headings on one line with their authored gap', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .heading { display:flex; align-items:center; font-size:14px; margin-bottom:25px; }
      .icon { width:20px; height:20px; margin-right:12px; }
    </style></head><body>
      <div class="heading">
        <svg class="icon" viewBox="0 0 24 24"><path d="M2 10l10-5 10 5-10 5z"></path></svg>
        GRADUATION
      </div>
    </body></html>`,
    'inline-flex-icon-heading',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /<w:p>[\s\S]*?<w:drawing>[\s\S]*?GRADUATION[\s\S]*?<\/w:p>/)
  assert.match(documentXml, /\u00a0/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('keeps bordered inline-block cards shrink-wrapped', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .title { display:inline-block; border:2px solid #000; padding:8px 10px; font-weight:bold; }
    </style></head><body><div class="title">数学的帰納法（不等式証明）</div></body></html><!-- public-hygiene: fixture -->`,
    'shrink-wrapped-card',
  )
  const documentXml = await xml('word/document.xml')
  const width = Number(documentXml.match(/<w:tblW\b[^>]*w:w="(\d+)"[^>]*\/>/)?.[1])
  assert.ok(width > 1000 && width < 6000, `inline card should not span the page, got ${width}`)
})

test('keeps vertical heading borders out of CSS margins', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      h2 {
        border-left:5px solid #148f77;
        padding-left:12px;
        margin-top:40px;
        margin-bottom:12px;
      }
    </style></head><body>
      <p>Before</p>
      <h2>Bordered heading</h2>
      <p>After</p>
    </body></html>`,
    'bordered-heading-spacing',
  )
  const documentXml = await xml('word/document.xml')
  const headingXml = documentXml
    .split('</w:p>')
    .find((paragraph) => paragraph.includes('Bordered heading'))
  assert.ok(headingXml, 'bordered heading paragraph should be present')
  assert.match(headingXml, /w:pBdr/)
  assert.match(headingXml, /w:before="0"/)
  assert.match(headingXml, /w:after="0"/)
})

test('overrides built-in Word spacing for zero-margin headings', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; padding:40px 60px; }
      h1 { margin:0 0 8px; padding:0; }
    </style></head><body><h1>Top aligned report title</h1></body></html>`,
    'zero-heading-margin',
  )
  const documentXml = await xml('word/document.xml')
  const headingXml = documentXml.match(
    /<w:p>[\s\S]*?<w:pStyle w:val="Heading1"\/>[\s\S]*?<\/w:p>/,
  )?.[0]
  assert.ok(headingXml)
  assert.match(headingXml, /w:spacing[^>]*w:before="0"/)
})

test('rasterizes gradient cards nested inside layout rows', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .summary { display:grid; grid-template-columns:1fr 1fr; width:700px; }
      .fee {
        background:linear-gradient(135deg,#ffd700,#ffc700);
        border-radius:12px;
        box-shadow:0 4px 8px rgba(255,170,0,.25);
        padding:14px 20px;
      }
    </style></head><body>
      <section class="summary">
        <div>Event details</div>
        <div class="fee"><strong>¥3,000</strong><div>per person</div></div>
      </section>
    </body></html>`,
    'gradient-row-card',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Event details/)
  assert.match(documentXml, /w:line="\d+" w:lineRule="atLeast"/)
  assert.ok(
    Object.keys(zip.files).some((name) => name.startsWith('word/media/')),
    'gradient fee card should use screenshot media',
  )
})

test('preserves clipped gradient heading text as visual media', async () => {
  const { ir, screenshotText, zip } = await convertHtml(
    `<!doctype html><html><head><style>
      body { background:#0a0f1e; }
      h1 {
        font-size:42px;
        background:linear-gradient(90deg,#e6ecff 30%,#a78bfa 70%);
        -webkit-background-clip:text;
        background-clip:text;
        color:transparent;
      }
    </style></head><body><h1>Gradient report title</h1></body></html>`,
    'gradient-heading-text',
  )
  assert.match(JSON.stringify(ir), /"type":"(?:image|floatimg)"/)
  assert.match(screenshotText, /Gradient report title/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('splits a tall visual section into flowable screenshot slices', async () => {
  const { ir, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      .visual {
        display:block;
        width:700px;
        height:1300px;
        background:linear-gradient(#0c1228,#0a0f1e);
        background-color:#0c1228;
      }
    </style></head><body>
      <img class="visual" src="/missing-map.png" alt="Consumer map" />
      <p>Following content</p>
    </body></html>`,
    'tall-visual-slices',
  )
  const slices = ir.filter((node) => node.type === 'image' && node.clip)
  assert.equal(slices.length, 2, JSON.stringify(ir))
  assert.equal(slices[0].height, 700)
  assert.equal(slices[1].height, 600)
  assert.match(screenshotText, /Consumer map/)
})

test('keeps a KPI row horizontal inside a card nested in a layout row', async () => {
  const { ir, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .hero { display:grid; grid-template-columns:1.4fr .9fr; gap:18px; width:900px; }
      .panel { background:#fff; border:1px solid #ddd; border-radius:12px; padding:22px; }
      .stats { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; }
      .stat { background:#f8fafc; border:1px solid #ddd; padding:16px; }
      .value { font-size:28px; font-weight:800; }
    </style></head><body>
      <div class="hero">
        <div class="panel">
          <h2>Highlights</h2>
          <div class="stats">
            <div class="stat"><div>Views</div><div class="value">48,433,718</div></div>
            <div class="stat"><div>Hours</div><div class="value">972,931.9915</div></div>
            <div class="stat"><div>Revenue</div><div class="value">3,988,282.297</div></div>
          </div>
        </div>
        <div class="panel"><h2>Summary</h2><p>Supporting details</p></div>
      </div>
    </body></html>`,
    'nested-kpi-row',
  )
  const outerRow = ir.find(
    (node) => node.type === 'kpirow' && JSON.stringify(node).includes('Highlights'),
  )
  assert.ok(outerRow)
  assert.ok(JSON.stringify(outerRow).includes('"cells":[{"children"'))
  const documentXml = await xml('word/document.xml')
  assert.ok(
    (documentXml.match(/<w:tbl>/g) || []).length >= 1,
    'outer layout row should remain a table',
  )
  // the lone panel card hoists its border onto the outer row's cell so
  // sibling panels share the row height (flex-stretch equal heights)
  assert.ok(
    /<w:tcBorders>(?:(?!<\/w:tcBorders>)[^])*?w:color="DDDDDD"/.test(documentXml),
    'panel border should sit on the outer row cell',
  )
})

test('preserves vertical padding on plain section wrappers', async () => {
  const { ir } = await convertHtml(
    `<!doctype html><html><head><style>
      section { padding:22px 0 8px; }
      h2 { margin:0; }
    </style></head><body><section><h2>Section heading</h2><p>Section body</p></section></body></html>`,
    'section-wrapper-padding',
  )
  const headingIndex = ir.findIndex(
    (node) => node.type === 'heading' && JSON.stringify(node).includes('Section heading'),
  )
  assert.ok(headingIndex > 0)
  assert.equal(ir[headingIndex - 1].type, 'spacer')
  assert.equal(ir[headingIndex - 1].px, 22)
  assert.equal(ir.at(-1).type, 'spacer')
  assert.equal(ir.at(-1).px, 8)
})

test('keeps empty checkbox boxes inside flex rows', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      ul { list-style:none; padding:0; }
      li { display:flex; align-items:center; gap:12px; }
      .box { width:22px; height:22px; border:1.5px solid #cbd5e1; border-radius:6px; }
    </style></head><body>
      <ul><li><span class="box"></span><span>Unchecked criterion</span></li></ul>
    </body></html>`,
    'flex-row-checkbox',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /w:color="CBD5E1"/)
  assert.ok((documentXml.match(/<w:tbl>/g) || []).length >= 2)
  assert.match(documentXml, /Unchecked criterion/)
})

test('rasterizes compact rounded cards with exact edge weight', async () => {
  const { zip } = await convertHtml(
    `<!doctype html><html><head><style>
      .card {
        width:600px;
        border:1px solid #e2e8f0;
        border-radius:12px;
        overflow:hidden;
        box-shadow:0 2px 8px -4px rgba(15,23,42,.08);
      }
      .header { background:#059669; color:white; padding:12px 18px; }
      .body { padding:12px 18px; }
    </style></head><body>
      <div class="card"><div class="header">SANGAT KOMPETEN</div><div class="body">Compact assessment card.</div></div>
    </body></html>`,
    'compact-rounded-card',
  )
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('preserves CSS display-table rows as one native table', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .competencies { display:table; width:700px; }
      .row { display:table-row; }
      .code, .description { display:table-cell; padding:12px; }
      .code { width:120px; color:white; background:#1a5276; }
      .description { background:#dceef8; }
    </style></head><body>
      <div class="competencies">
        <div class="row"><div class="code">KD 3.2</div><div class="description">Analyze quantities</div></div>
        <div class="row"><div class="code">KD 4.2</div><div class="description">Present measurements</div></div>
      </div>
    </body></html>`,
    'css-display-table',
  )
  const documentXml = await xml('word/document.xml')
  assert.equal((documentXml.match(/<w:tbl>/g) || []).length, 1)
  assert.equal((documentXml.match(/<w:tr>/g) || []).length, 2)
  assert.equal((documentXml.match(/<w:tc>/g) || []).length, 4)
  assert.match(documentXml, /w:fill="1A5276"/)
  assert.match(documentXml, /w:fill="DCEEF8"/)
})

test('removes default Word paragraph spacing inside compact table cells', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      table { border-collapse:collapse; width:600px; font-size:12px; }
      td { padding:7px 12px; border:1px solid #d5d8dc; }
    </style></head><body>
      <table><tbody>
        <tr><td>Panjang</td><td>Keindahan</td></tr>
        <tr><td>Massa</td><td>Kejujuran</td></tr>
      </tbody></table>
    </body></html>`,
    'compact-table-cell-spacing',
  )
  const documentXml = await xml('word/document.xml')
  assert.ok((documentXml.match(/<w:spacing w:after="0" w:before="0"\/>/g) || []).length >= 4)
  assert.doesNotMatch(documentXml, /<w:trHeight/)
})

test('preserves horizontal padding on highlighted inline chips', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .dim {
        display:inline-block;
        background:#1a5276;
        color:white;
        font-weight:700;
        padding:1px 7px;
      }
    </style></head><body><p>Dimension <span class="dim">[L]</span></p></body></html>`,
    'highlighted-inline-padding',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /\[L\]/)
  assert.ok((documentXml.match(/w:fill="1A5276"/g) || []).length >= 3)
  assert.ok((documentXml.match(/>\u00a0+</g) || []).length >= 2)
})

test('subtracts page-column padding from nested layout width', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .resume { display:grid; grid-template-columns:278px 516px; width:794px; }
      aside { min-height:800px; background:#1e3a5f; padding:30px; }
      main { min-height:800px; padding:40px 50px; }
      .skill { display:flex; justify-content:space-between; }
      .track { width:187px; height:8px; background:#ff6633; }
    </style></head><body>
      <div class="resume">
        <aside>Sidebar</aside>
        <main><div class="skill"><span>Planning</span><div class="track"></div></div></main>
      </div>
    </body></html>`,
    'padded-page-columns',
  )
  const documentXml = await xml('word/document.xml')
  const tableWidths = [...documentXml.matchAll(/<w:tblW w:type="dxa" w:w="(\d+)"\/>/g)].map(
    (match) => Number(match[1]),
  )
  // main column minus its 50px paddings at the margin-floor canvas scale:
  // (516 - 100) * 15 * (794 / 842) ≈ 5885 dxa; unpadded would be ~7300.
  assert.ok(
    tableWidths.some((width) => width >= 5700 && width <= 6100),
    `nested main-column table should use padded width, got ${tableWidths.join(', ')}`,
  )
})

test('keeps a full-width resume header above tall grid columns', async () => {
  const { ir } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .resume {
        display:grid;
        grid-template-columns:3fr 2fr;
        gap:24px;
        width:760px;
        position:relative;
      }
      header { grid-column:1 / -1; }
      .left, .right { min-height:850px; }
      .decoration { position:absolute; right:0; bottom:0; width:80px; height:80px; background:#eef; }
    </style></head><body>
      <div class="resume">
        <div class="decoration"></div>
        <header><h1>Sarah Mitchell</h1><p>Senior Brand Designer</p></header>
        <main class="left"><h2>Work Experience</h2><p>Left column</p></main>
        <aside class="right"><h2>About</h2><p>Right column</p></aside>
      </div>
    </body></html>`,
    'headed-page-columns',
  )
  const pageColumns = ir.find((node) => node.type === 'table' && node.pageColumns)
  assert.ok(pageColumns, 'the two tall columns should remain a page-column table')
  assert.deepEqual(pageColumns.colWidths.length, 2)
  assert.ok(
    ir.some(
      (node) => node.type === 'heading' && node.runs?.some((run) => run.text.includes('Sarah')),
    ),
    'the spanning header should stay before the column table',
  )
})

test('paints the cell edge behind a top-aligned visual page-column header', async () => {
  const { ir, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      * { box-sizing:border-box; }
      body { margin:0; background:#f4f7fb; }
      .layout { display:grid; grid-template-columns:220px 574px; width:794px; }
      aside { min-height:800px; background:#0d2137; }
      main { min-height:800px; padding:0 48px; }
      .hero { height:140px; margin:0 -48px; background:linear-gradient(135deg,#0d2137,#1e4d80); color:white; }
    </style></head><body>
      <div class="layout">
        <aside>Navigation</aside>
        <main><div class="hero"><h1>SDK Reference</h1></div><p>Body text</p></main>
      </div>
    </body></html>`,
    'page-column-visual-header',
  )
  const pageColumns = ir.find((node) => node.type === 'table' && node.pageColumns)
  assert.ok(pageColumns)
  assert.equal(pageColumns.rows[0].cells[0].gapAfterPx, 0)
  assert.equal(pageColumns.rows[0].cells[1].children[0].type, 'image')
  assert.equal(pageColumns.rows[0].cells[1].children[0].align, 'left')
  assert.equal(pageColumns.rows[0].cells[1].children[0].bleedLeftPx, 48)
  assert.equal(pageColumns.rows[0].cells[1].children[0].topEdgeFill, '0D2137')
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /<w:top w:val="single" w:color="0D2137" w:sz="48" w:space="0"\/>/)
})

test('preserves tall multi-row table grids as side-by-side visual rows', async () => {
  const { ir, zip, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .ranking-grid {
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:16px;
        width:1200px;
      }
      .table-wrap { height:760px; border:1px solid #ddd; overflow:hidden; }
      table { width:100%; border-collapse:collapse; }
      td { height:60px; border-bottom:1px solid #ddd; }
    </style></head><body>
      <div class="ranking-grid">
        <div class="table-wrap"><table><tr><td>Views table</td></tr></table></div>
        <div class="table-wrap"><table><tr><td>Watch time table</td></tr></table></div>
        <div class="table-wrap"><table><tr><td>Subscribers table</td></tr></table></div>
        <div class="table-wrap"><table><tr><td>Revenue table</td></tr></table></div>
      </div>
    </body></html>`,
    'large-table-grid',
  )
  const grid = ir.find((node) => node.type === 'table' && node.largeTableGrid)
  assert.ok(grid, 'the tall table grid should retain an explicit layout wrapper')
  const contentRows = grid.rows.filter((row) => !row.gapRow)
  assert.equal(contentRows.length, 2)
  assert.ok(
    contentRows.every(
      (row) =>
        row.cells.filter((cell) => cell.children?.length).length === 2 &&
        row.cells
          .filter((cell) => cell.children?.length)
          .every((cell) => cell.children[0].type === 'image'),
    ),
  )
  assert.equal(
    ir.some((node) => node.pageColumns),
    false,
  )
  assert.match(screenshotText, /Views table/)
  assert.match(screenshotText, /Revenue table/)
  assert.ok(Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 4)
})

test('rasterizes a page column containing positioned decorations', async () => {
  const { ir } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .resume { display:flex; gap:24px; width:760px; }
      .left, .right { min-height:850px; }
      .left { width:260px; }
      .right { position:relative; width:476px; padding-top:180px; }
      .photo { position:absolute; top:0; right:40px; width:140px; height:140px; background:#b88; }
      .dot { position:absolute; left:4px; top:300px; width:10px; height:10px; background:#843; }
    </style></head><body>
      <div class="resume">
        <aside class="left"><h1>Naval Kishor</h1><p>Education and skills</p></aside>
        <main class="right">
          <div class="photo"></div><div class="dot"></div>
          <h2>Professional Summary</h2><p>Summary text</p>
          <h2>Work Experience</h2><p>Experience text</p>
        </main>
      </div>
    </body></html>`,
    'positioned-page-column',
  )
  const pageColumns = ir.find((node) => node.type === 'table' && node.pageColumns)
  assert.ok(pageColumns)
  assert.equal(pageColumns.rows[0].cells[1].children.length, 1)
  assert.equal(pageColumns.rows[0].cells[1].children[0].type, 'image')
  assert.equal(JSON.stringify(pageColumns).includes('floatimg'), false)
})

test('preserves pseudo counter order and literal pseudo bullets', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      ol, ul { list-style:none; }
      ol { counter-reset:step; }
      ol li { counter-increment:step; }
      ol li::before { content:"STEP " counter(step); text-transform:uppercase; }
      ul li::before { content:"■"; color:#6b2c2c; }
    </style></head><body>
      <ol><li><div>Review request</div><p>Verify details</p></li></ol>
      <ul><li>Square marker item</li></ul>
    </body></html>`,
    'pseudo-list-markers',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /STEP 1/)
  assert.doesNotMatch(documentXml, /1STEP/)
  assert.match(documentXml, /■/)
})

test('preserves short literal markers on list items', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      ol { list-style:none; padding:0; }
      li { position:relative; padding-left:36px; }
      li::before { position:absolute; left:8px; color:#1e40af; }
      li:nth-child(1)::before { content:"① "; }
      li:nth-child(2)::before { content:"② "; }
    </style></head><body>
      <ol><li>First choice</li><li>Second choice</li></ol>
    </body></html>`,
    'literal-list-markers',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /①/)
  assert.match(documentXml, /②/)
  assert.match(documentXml, /w:ascii="Arial Unicode MS"/)
})

test('preserves pseudo symbols in multicolumn lists and labels', async () => {
  const { zip, xml, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      .ribbon { display:inline-block; background:#435638; color:white; padding:4px 12px; }
      .ribbon::before { content:"✓ "; color:#d8c48d; }
      ul { columns:2; width:600px; list-style:none; padding:0; }
      li { position:relative; padding-left:24px; }
      li::before { content:"☑"; position:absolute; left:0; color:#4a5e3a; }
    </style></head><body>
      <div class="ribbon">All Included</div>
      <ul><li>One</li><li>Two</li><li>Three</li><li>Four</li></ul>
    </body></html>`,
    'multicolumn-pseudo-markers',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /✓/)
  assert.match(screenshotText, /One/)
  assert.match(screenshotText, /Four/)
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 4,
    'the four exact pseudo-marker rows should be screenshots',
  )
})

test('screenshots compact nowrap labels inside table cells', async () => {
  const { zip, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      table { width:600px; border-collapse:collapse; }
      td { width:120px; border:1px solid #ddd; }
      .badge { display:inline-block; white-space:nowrap; color:white; background:#f39c12;
        padding:3px 10px; border-radius:2px; font-size:12px; }
    </style></head><body><table><tr><td>
      <span class="badge">AC・HSPは中</span>／<span class="badge">Fawn特化は空白</span>
    </td></tr></table></body></html>`,
    'nowrap-table-badges',
  )
  assert.match(screenshotText, /AC・HSPは中/)
  assert.match(screenshotText, /Fawn特化は空白/)
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 1,
    'the compact labels and separator should remain one unbroken cell image',
  )
})

test('rasterizes compact parallel reference blocks', async () => {
  const { ir, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      .references { display:flex; gap:50px; width:500px; }
      .reference { flex:1; }
    </style></head><body><div class="references">
      <div class="reference"><div>Sarah Chen</div><div>GlobalTech / Director</div><div>555-1111</div></div>
      <div class="reference"><div>Michael Thompson</div><div>Pinnacle / VP Operations</div><div>555-2222</div></div>
    </div></body></html>`,
    'parallel-reference-blocks',
  )
  assert.ok(ir.some((node) => node.type === 'image'))
  assert.match(screenshotText, /Michael Thompson/)
})

test('preserves icon-font contacts and compact floated skill bars', async () => {
  const { ir, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      .contact { display:flex; gap:12px; }
      .contact i { display:block; width:20px; height:20px; }
      .contact i::before { content:"☎"; }
      .skills { width:220px; }
      .skill { width:220px; height:50px; }
      .skill .percent { float:right; margin-top:-20px; }
      .track { height:8px; background:#8fb8dd; }
      .progress { width:88%; height:100%; background:#fff; }
      .experience { display:flex; width:420px; gap:12px; }
      .title { flex:1; }
    </style></head><body>
      <div class="contact"><i></i><span>555-123-4567</span></div>
      <div class="skills">
        <div class="skill"><div>Process Optimization <span class="percent">95%</span></div>
          <div class="track"><div class="progress"></div></div></div>
        <div class="skill"><div>Supply Chain Management <span class="percent">88%</span></div>
          <div class="track"><div class="progress"></div></div></div>
      </div>
      <div class="experience"><div>✓</div><div class="title">Senior Coordinator</div><div>2022-2024</div></div>
    </body></html>`,
    'compact-resume-controls',
  )
  assert.match(screenshotText, /Supply Chain Management/)
  assert.ok(ir.filter((node) => node.type === 'image').length >= 1)
  const experience = ir.find(
    (node) =>
      node.type === 'kpirow' &&
      node.cells?.some((cell) => JSON.stringify(cell).includes('Senior Coordinator')),
  )
  assert.ok(experience)
  assert.ok(
    experience.itemWidths[2] >= experience.colWidths[2] + 15,
    'the short date range needs intrinsic-width slack',
  )
})

test('renders closed details as a static summary without hidden content', async () => {
  const { ir, xml } = await convertHtml(
    `<!doctype html><html><body>
      <details style="border:1px solid #86efac;background:#f0fdf4;padding:8px 14px">
        <summary>Show answer</summary>
        <p>Hidden answer text</p>
      </details>
    </body></html>`,
    'closed-details',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Show answer/)
  assert.doesNotMatch(documentXml, /Hidden answer text/)
  assert.equal(
    ir.some((node) => node.type === 'card'),
    false,
  )
  assert.ok(ir.some((node) => node.type === 'para' && node.style?.shading === 'F0FDF4'))
})

test('splits emoji icons into font-unset runs for Word fallback', async () => {
  const { xml } = await convertHtml(
    '<!doctype html><html><body><h2>🏆 Performance 🟢</h2></body></html>',
    'emoji-icons',
  )
  const documentXml = await xml('word/document.xml')
  // Each emoji lands in its own run whose rPr carries no rFonts, so Word's
  // built-in color-emoji fallback applies. Forcing Segoe UI Emoji rendered
  // tofu (and flags as letters) on Mac Word.
  assert.match(documentXml, /<w:rPr>(?:(?!<\/w:rPr>|w:rFonts)[^])*?<\/w:rPr><w:t[^>]*>🏆<\/w:t>/)
  assert.match(documentXml, /<w:rPr>(?:(?!<\/w:rPr>|w:rFonts)[^])*?<\/w:rPr><w:t[^>]*>🟢<\/w:t>/)
  // The adjacent body word still gets a real font.
  assert.match(documentXml, /<w:rFonts[^>]*\/>[^]*?<w:t[^>]*> Performance <\/w:t>/)
})

test('maps rounded display fonts before generic cursive fallbacks', () => {
  assert.equal(mapFont('"Fredoka One", cursive'), 'Arial Rounded MT Bold')
  assert.equal(mapFont('"Pacifico", cursive'), 'Segoe Script')
  assert.equal(mapFont('"Open Sans", sans-serif'), 'Arial')
})

test('preserves colored literal bullets as editable runs', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><body>
      <div style="color:#e74c3c">• Colored bullet item</div>
    </body></html>`,
    'colored-literal-bullet',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /<w:color w:val="E74C3C"\/>/)
  assert.match(documentXml, /• Colored bullet item/)
  assert.doesNotMatch(documentXml, /<w:numPr>/)
})

test('keeps margins around thin decorative rules', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .divider { width:700px; height:3px; margin:30px 0; background:#e74c3c; }
    </style></head><body><div class="divider"></div><p>Question</p></body></html>`,
    'thin-rule-spacing',
  )
  const documentXml = await xml('word/document.xml')
  // 30px margins at the margin-floor canvas scale: the default 8px body
  // margin is floored to 24px per side, widening the reported canvas to
  // 826px, so 30 * 15 * (794 / 826) ≈ 433 twips.
  assert.match(documentXml, /<w:spacing w:after="433" w:before="433"/)
})

test('keeps broken image alt text as an editable fallback', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><body>
      <div><img src="file:///definitely-missing-image.png" alt="Rocket"><strong>FAUGET CO.</strong></div>
    </body></html>`,
    'broken-image-alt',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Rocket/)
  assert.match(documentXml, /FAUGET CO\./)
})

test('rasterizes painted broken-image placeholders at their authored size', async () => {
  const { zip, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      img { display:block; width:320px; height:240px; background:#e8e5df; border-radius:6px; }
    </style></head><body>
      <img src="file:///definitely-missing-stage-image.png" alt="Stage 1 Front — Earthworks">
    </body></html>`,
    'broken-image-placeholder',
  )
  assert.match(screenshotText, /Stage 1 Front/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('rasterizes compact wrapped flow diagrams without reflowing their steps', async () => {
  const steps = Array.from(
    { length: 7 },
    (_, index) =>
      `${index ? '<div class="arrow">→</div>' : ''}<div class="step">STEP ${index + 1}</div>`,
  ).join('')
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .outer { padding:10px; }
      .flow { display:flex; flex-wrap:wrap; width:420px; gap:8px; }
      .step { width:92px; padding:8px; border:1px solid #999; }
      .arrow { width:18px; }
    </style></head><body><main><section class="outer"><div class="flow">${steps}</div></section></main></body></html>`,
    'wrapped-flow',
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(documentXml, /STEP 1[\s\S]*?STEP 7/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('rasterizes a single full-page visual poster', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      @page { size:A4; margin:0; }
      body { margin:0; display:flex; justify-content:center; }
      .poster {
        width:794px; height:1123px; position:relative; overflow:hidden;
        background:linear-gradient(#fffef8,#f7f8ef);
      }
    </style></head><body><div class="poster"><h1>Visual poster</h1><p>Composed page</p></div></body></html>`,
    'single-visual-poster',
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(documentXml, /Visual poster/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('rasterizes a tall painted resume grid as natural page slices', async () => {
  const { zip, xml, ir, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; display:flex; justify-content:center; }
      .resume {
        width:700px; height:1500px; position:relative; overflow:hidden;
        background:#f5f1ed;
      }
      .header { height:220px; text-align:center; }
      .columns { display:grid; grid-template-columns:1fr 1.8fr; gap:24px; }
      .card { min-height:900px; background:white; padding:20px; }
    </style></head><body><main class="resume">
      <header class="header"><h1>Resume name</h1></header>
      <div class="columns"><section class="card">Profile</section><section class="card">Experience</section></div>
    </main></body></html>`,
    'flowing-visual-resume',
  )
  const documentXml = await xml('word/document.xml')
  assert.equal(ir.filter((node) => node.type === 'image').length, 2)
  assert.equal(ir.filter((node) => node.type === 'pagebreak').length, 0)
  assert.doesNotMatch(documentXml, /Resume name/)
  assert.match(screenshotText, /Resume name/)
  assert.match(screenshotText, /Experience/)
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'each resume viewport slice should be embedded as an image',
  )
})

test('moves a page-straddling contact card whole to the next slice', async () => {
  // Regression: the naive full-page-height slice boundary cut through the
  // contact card and Word's ~26px top inset then clipped the card's last
  // row ("Greater Houston…") at the page edge. Slices must cut in the gap
  // before the card and stay under the page budget.
  const { ir, xml, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; display:flex; justify-content:center; }
      .resume {
        width:700px; position:relative; overflow:hidden;
        background:#f5f1ed; padding-bottom:60px;
      }
      .head { height:200px; display:flex; gap:16px; }
      .work { height:760px; background:white; margin:0 24px; }
      .contact { height:320px; background:white; margin:40px 24px 0; padding:20px; }
      .contact p { margin:12px 0; }
    </style></head><body><main class="resume">
      <header class="head"><h1>Contact name</h1><span>General Manager</span></header>
      <section class="work">Experience body</section>
      <section class="contact">
        <p>Date of Birth: 01/01/1990</p>
        <p>(555) 010-0199</p>
        <p>contact@example.com</p>
        <p>Greater Houston, TX — open to relocate</p>
      </section>
    </main></body></html>`,
    'contact-card-slice',
  )
  const documentXml = await xml('word/document.xml')
  const slices = ir.filter((node) => node.type === 'image' && node.pageComposition)
  assert.equal(slices.length, 2)
  const budget = Math.round(1123 * 0.96)
  for (const slice of slices) {
    assert.ok(slice.height <= budget, `slice height ${slice.height} exceeds budget ${budget}`)
  }
  // The cut must land in the gap between the work and contact cards
  // (work bottom 960, contact top 1000), not inside the contact card.
  assert.ok(
    slices[0].height > 940 && slices[0].height < 1000,
    `first slice should end in the inter-card gap, got ${slices[0].height}`,
  )
  assert.match(screenshotText, /Greater Houston, TX — open to relocate/)
  assert.doesNotMatch(documentXml, /Greater Houston/)
})

test('binds a heading and its lede through spacers to the section body', async () => {
  // Regression: heading keepNext only reached the following spacer, so Word
  // orphaned section headings (and heading+subtitle pairs) at page bottoms
  // while the body table/image moved to the next page.
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { font-family: Arial; max-width: 700px; margin: 0 auto; }
      section { margin-top: 48px; }
      h2 { margin: 0 0 8px; }
      .subtitle { color: #667; margin: 0 0 24px; }
      td { border: 1px solid #ccc; padding: 6px 10px; }
    </style></head><body>
      <p>Intro paragraph before any section.</p>
      <section>
        <h2>Yearly Summary</h2>
        <p class="subtitle">Compares uploads and revenue by year</p>
        <table><tr><td>2024</td><td>129</td></tr><tr><td>2025</td><td>287</td></tr></table>
      </section>
    </body></html>`,
    'heading-lede-chain',
  )
  const documentXml = await xml('word/document.xml')
  const paragraphs = documentXml.split('</w:p>')
  const headingIndex = paragraphs.findIndex((p) => p.includes('Yearly Summary'))
  const ledeIndex = paragraphs.findIndex((p) => p.includes('Compares uploads'))
  assert.ok(headingIndex >= 0 && ledeIndex > headingIndex)
  assert.match(paragraphs[headingIndex], /<w:keepNext\/>/)
  assert.match(paragraphs[ledeIndex], /<w:keepNext\/>/)
  // Every paragraph between heading and lede (post-heading spacers) binds on.
  for (let i = headingIndex + 1; i < ledeIndex; i++) {
    assert.match(paragraphs[i], /<w:keepNext\/>/)
  }
  // The intro paragraph before the section must NOT join any keep chain.
  const introIndex = paragraphs.findIndex((p) => p.includes('Intro paragraph'))
  assert.doesNotMatch(paragraphs[introIndex], /<w:keepNext\/>/)
})

test('removes paragraph margins inside padded table cells', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><body><table><tr><td style="padding:8px">
      <p style="margin:12px 0">Compact row</p>
    </td></tr></table></body></html>`,
    'nested-cell-paragraph-spacing',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /<w:spacing w:after="0" w:before="0"/)
})

test('renders an active SPA page as full-page screenshots', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      nav { position:fixed; inset:0 0 auto; height:64px; background:#071426; }
      .page-view { display:none; }
      .page-view.active-page { display:block; min-height:2400px; }
      .hero { height:1200px; background:linear-gradient(135deg,#071426,#1f4e9a); }
      .container { width:90%; margin:auto; }
      .row { display:flex; }
      img { width:400px; height:300px; }
    </style></head><body>
      <nav>Navigation</nav>
      <section class="page-view active-page">
        <div class="hero"><div class="container row"><h1>Visual website</h1>
          <img src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='400' height='300'%3E%3Crect width='400' height='300' fill='%2379d0da'/%3E%3C/svg%3E">
        </div></div>
      </section>
    </body></html>`,
    'spa-page',
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(
    documentXml,
    /w:br w:type="page"/,
    'full-page image slices should advance naturally without inserting blank pages',
  )
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'each website viewport slice should be embedded as an image',
  )
  assert.doesNotMatch(documentXml, /Visual website/)
})

test('keeps empty bordered answer boxes with authored height', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .box { border:1px solid #d0d5dd; background:#fcfcfd; padding:12px; }
      .field-label { display:block; font-weight:700; font-size:12px; }
      .answer-box {
        min-height:58px;
        border:1px solid #d0d5dd;
        background:#fff;
        margin-top:8px;
      }
    </style></head><body>
      <div class="box">
        <span class="field-label">WHAT SHOULD THE AI HELP CUSTOMERS DO?</span>
        <div class="answer-box"></div>
      </div>
    </body></html>`,
    'answer-box',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /WHAT SHOULD THE AI HELP/)
  // Answer box becomes a fixed-height form cell, not a collapsed spacer.
  assert.match(documentXml, /w:sdt/)
  assert.ok(
    /w:trHeight[^>]*w:val="[1-9][0-9]{2,}"/.test(documentXml),
    'answer box row should keep a substantial height',
  )
})

test('keeps short flex meta labels on one line', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { max-width:880px; margin:0 auto; padding:32px 80px; }
      .meta-row { display:flex; gap:28px; font-size:13px; }
      .meta-row span strong { font-weight:600; }
    </style></head><body>
      <h1>Notice title</h1>
      <div class="meta-row">
        <span><strong>수신</strong> SM 전원</span>
        <span><strong>발신</strong> 박규하 RM</span>
        <span><strong>일자</strong> 2025년 7월 11일</span>
      </div>
    </body></html>`,
    'flex-meta-row',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /수신/)
  assert.match(documentXml, /SM 전원/)
  assert.match(documentXml, /발신/)
  assert.match(documentXml, /박규하 RM/)
  // Each meta cell should stay a single paragraph (no mid-label wrap split
  // into a second run/paragraph inside the cell).
  assert.doesNotMatch(documentXml, /수신[^<]{0,20}SM<\/w:t><\/w:r><\/w:p><w:p>/)
})

test('keeps an explicit white body over a gray browser canvas', async () => {
  const { ir, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      html, body { background:#f0f0f0; }
      body { width:794px; min-height:1123px; margin:20px auto; padding:80px; background:#fff; }
    </style></head><body><h1>Cover letter</h1><p>White paper content</p></body></html>`,
    'white-paper-gray-canvas',
  )
  assert.equal(
    ir.some((node) => node.type === 'pagebg'),
    false,
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(documentXml, /<w:background w:color="F0F0F0"\/>/)
})

test('keeps white cards over gradient page backgrounds', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 24px;
      }
      .content-box {
        background: white;
        border-radius: 8px;
        padding: 20px;
        margin: 20px 0;
        box-shadow: 0 2px 10px rgba(0,0,0,.1);
        border: 1px solid rgba(0,0,0,.08);
      }
    </style></head><body>
      <div class="content-box">On this day the first party agrees to provide comprehensive
      maintenance services covering documentation, training, monitoring visits,
      and on-site inspection within forty eight hours of remote failure.</div>
      <div class="content-box">The second party shall maintain strict confidentiality regarding all
      proprietary business information and financial data of the first party under
      this maintenance cooperation agreement letter for two full years.</div>
    </body></html>`,
    'white-card-on-gradient',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /first party agrees/)
  assert.match(documentXml, /second party shall maintain/)
  const whiteFills = (documentXml.match(/w:fill="FFFFFF"/g) || []).length
  assert.ok(whiteFills >= 2, `expected white card shading, got ${whiteFills}`)
})

test('composites semi-transparent backgrounds onto white', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .note {
        background:rgba(196,30,58,.08);
        color:#c41e3a;
        padding:8px 12px;
      }
    </style></head><body>
      <p class="note">Tinted note</p>
    </body></html>`,
    'rgba-tint',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Tinted note/)
  // Solid accent red must not be used as the fill; the tint composites to ~FAEDEF.
  assert.doesNotMatch(documentXml, /w:fill="C41E3A"/)
  assert.match(documentXml, /w:fill="FAEDEF"/)
})

test('composites semi-transparent table stripes onto a dark panel', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { background:#0a0f1e; color:#aab4d4; }
      table { width:600px; background:#10172e; border-collapse:collapse; }
      th { background:rgba(16,23,46,.85); color:#cfd8f5; padding:8px; }
      td { color:#aab4d4; padding:8px; }
      tbody tr:nth-child(odd) { background:rgba(15,21,40,.4); }
    </style></head><body>
      <table>
        <thead><tr><th>Tier</th><th>Audience</th></tr></thead>
        <tbody>
          <tr><td>Entry</td><td>Students</td></tr>
          <tr><td>Premium</td><td>Creators</td></tr>
        </tbody>
      </table>
    </body></html>`,
    'dark-table-stripes',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Entry/)
  assert.doesNotMatch(documentXml, /w:fill="9FA1A9"/)
  assert.match(documentXml, /w:fill="10162C"|w:fill="10172E"/)
})

test('keeps long email contact cells on one line in flex footers', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .footer-contact {
        display:flex;
        justify-content:space-between;
        align-items:center;
        border-top:2px solid #d4af37;
        padding-top:12px;
        font-size:15px;
        color:#8b6f47;
      }
    </style></head><body>
      <p>Resume body</p>
      <div class="footer-contact">
        <span>info@candidatename.email.com</span>
        <span>(555) 456-7890</span>
        <span>1234 Professional Drive, Business City, ST</span>
      </div>
    </body></html>`,
    'email-footer',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /info@candidatename\.email\.com/)
  // Email must live in a reasonably wide cell (not ~227px that wraps "m").
  const widths = [...documentXml.matchAll(/<w:tcW[^>]*\bw:w="(\d+)"/g)].map((m) => Number(m[1]))
  assert.ok(widths.length >= 3, `expected footer table cells, got ${widths.length}`)
  // ~280px email cell ≈ 4200 twips; gaps may shrink but content stays wide.
  assert.ok(
    widths.some((w) => w >= 3800),
    `expected a wide contact cell for the email, got ${widths.join(',')}`,
  )
})

test('screenshots translucent number badges on colored headers', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .head {
        background:#2563eb;
        color:#fff;
        padding:18px 20px;
        text-align:center;
        width:220px;
        display:flex;
        flex-direction:column;
        align-items:center;
        gap:8px;
      }
      .num {
        display:inline-flex;
        width:30px;height:30px;
        align-items:center;justify-content:center;
        border-radius:999px;
        background:rgba(255,255,255,.2);
        color:#fff;
        font-weight:700;
      }
      .tag-pill {
        display:inline-block;
        background:#f5f5f5;
        border:1px solid #ccc;
        padding:2px 8px;
        border-radius:12px;
        margin:2px 4px;
      }
    </style></head><body>
      <div class="head">
        <div class="num">1</div>
        <div>理由がわかる</div>
      </div>
      <div>
        <span class="tag-pill">#休養</span><!-- public-hygiene: fixture -->
        <span class="tag-pill">#疲労回復</span><!-- public-hygiene: fixture -->
        <span class="tag-pill">#コンディショニング</span>
      </div>
    </body></html>`,
    'translucent-badge',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /理由/) // public-hygiene: fixture
  // Must not become a white color-bar with white "1" (invisible badge).
  assert.doesNotMatch(documentXml, /w:fill="FFFFFF"[\s\S]{0,200}>1</)
  const media = Object.keys(zip.files).filter((name) => name.startsWith('word/media/'))
  assert.ok(media.length >= 1, 'circle badge should be captured as an image')
  // Number chip is one image; tag cloud may be a second combined screenshot.
  assert.ok(media.length <= 3, `too many rasterized chips, got ${media.length} media`)
  // Tag cloud is captured as one image (not editable text runs).
  assert.ok(
    media.length >= 2 || documentXml.includes('#休養'), // public-hygiene: fixture
    'tag pills should appear as text or as a combined screenshot',
  )
})

test('keeps empty underline fill-in fields', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .field-line { border-bottom:1px solid #000; height:20px; width:280px; }
      .dotted-line {
        border-bottom:1px dotted black;
        display:inline-block;
        min-width:60px;
        margin:0 3px;
      }
    </style></head><body>
      <p>Full Name:</p>
      <div class="field-line"></div>
      <p>On this day <span class="dotted-line"></span> date</p>
    </body></html>`,
    'fill-in-lines',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Full Name/)
  assert.match(documentXml, /w:bottom[^>]*w:color="000000"/)
  assert.match(documentXml, /w:u w:val="dotted"/)
  assert.match(documentXml, /On this day/)
})

test('anchors absolute banner decorations inside their card', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; padding:32px 80px; }
      .banner { background:#8ba05c; color:#fff; padding:40px 50px; margin:-32px -80px 40px; position:relative; }
      .banner h1 { font-size:3rem; margin:0; text-align:center; }
      .leaf { position:absolute; top:50%; right:50px; transform:translateY(-50%); width:60px; height:60px; }
    </style></head><body>
      <div class="banner">
        <h1>Impact Report</h1>
        <svg class="leaf" viewBox="0 0 24 24" fill="#ffffff"><path d="M4 4h16v16H4z"></path></svg>
      </div>
      <p>Body paragraph after the banner.</p>
    </body></html>`,
    'banner-decoration-anchor',
  )
  const documentXml = await xml('word/document.xml')
  const anchor = documentXml.indexOf('<wp:anchor')
  const table = documentXml.indexOf('<w:tbl>')
  assert.ok(anchor > -1, 'decoration should emit a floating anchor')
  assert.ok(table > -1 && anchor > table, 'anchor should live inside the card table')
  assert.match(documentXml, /wp:positionH relativeFrom="column"/)
  assert.match(documentXml, /wp:positionV relativeFrom="paragraph"/)
  assert.match(documentXml, /behindDoc="0"/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('screenshots icon-font glyphs that live in pseudo elements', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .contact { display:flex; gap:10px; align-items:center; }
      .icon { width:20px; height:20px; display:block; }
      .icon::before { content:"\\260E"; font-size:16px; }
    </style></head><body>
      <div class="contact"><i class="icon"></i><span>+1 555-789-0123</span></div>
    </body></html>`,
    'icon-font-glyph',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /555-789-0123/)
  const drawing = documentXml.indexOf('<w:drawing>')
  const phone = documentXml.indexOf('555-789-0123')
  assert.ok(drawing > -1 && drawing < phone, 'icon screenshot should precede the phone number')
})

test('keeps authored newlines in pre-wrap text blocks', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><body>
      <div style="white-space:pre-wrap;">Hi [First Name],

I am reaching out about the launch.
Second line of the same paragraph.

Closing line.</div>
    </body></html>`,
    'pre-wrap-newlines',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Hi \[First Name\],/)
  const breaks = (documentXml.match(/<w:br\/>/g) || []).length
  assert.ok(breaks >= 4, `authored newlines should become breaks, got ${breaks}`)
})

test('keeps a row of 3+ tall equal pricing cards side by side', async () => {
  const card = (name, price) =>
    `<div class="pkg"><h3>${name}</h3><p>R$ ${price}</p>
     <ul><li>Feature one</li><li>Feature two</li><li>Feature three</li></ul></div>`
  const { ir } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .grid { display:grid; grid-template-columns:repeat(5, 1fr); gap:0; }
      .pkg { border-right:1px solid #ddd; padding:40px 20px; height:560px; box-sizing:border-box; }
    </style></head><body>
      <div class="grid">
        ${card('Kit Essencial', 490)}${card('Kit Comunicacao', 890)}${card('Kit Lancamento', 1490)}
        ${card('Social Business', 2490)}${card('Social Premium', 3990)}
      </div>
    </body></html>`,
    'parallel-pricing-cards',
  )
  const row = ir.find((node) => node.type === 'kpirow' && (node.cells || []).length === 5)
  assert.ok(row, 'five tall equal columns should classify as one kpirow')
})

test('carries flattened section shading through spacers and nested blocks', async () => {
  const paras = Array.from({ length: 10 }, (_, i) => `<p>Filler paragraph ${i + 1}.</p>`).join('')
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; background:#efe9dd; }
      section { background:#ffffff; padding:40px; }
      .cols { display:flex; gap:24px; }
    </style></head><body>
      <section>
        <h2>White section</h2>
        ${paras}
        <div class="cols"><div>Column A</div><div>Column B</div></div>
        ${paras}
      </section>
    </body></html>`,
    'flatten-shading',
  )
  const documentXml = await xml('word/document.xml')
  const fills = (documentXml.match(/w:fill="FFFFFF"/g) || []).length
  assert.ok(fills >= 10, `flattened white section should shade its blocks, got ${fills}`)
  // the zero-height spacer paragraphs must be shaded too, not just text paras
  assert.match(documentXml, /<w:shd[^>]*w:fill="FFFFFF"[^>]*\/><w:spacing[^>]*w:lineRule="exact"/)
})

test('keeps section watermark numerals past the first page', async () => {
  const section = (num, tint) => `
    <section style="background:${tint}; position:relative; padding:120px 40px; height:760px; box-sizing:border-box;">
      <div class="num">${num}</div>
      <h2>Section ${num}</h2>
      <p>Section body copy.</p>
      <div style="display:flex; gap:20px;"><div>Cell A</div><div>Cell B</div></div>
    </section>`
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; background:#f4f0e7; }
      .num { position:absolute; top:24px; left:24px; font-size:110px; line-height:1; opacity:.08; }
    </style></head><body>
      ${section('01', '#ffffff')}${section('02', '#fdfbf7')}${section('03', '#ffffff')}
    </body></html>`,
    'section-watermarks',
  )
  const documentXml = await xml('word/document.xml')
  const anchors = (documentXml.match(/wp:positionH relativeFrom="column"/g) || []).length
  assert.ok(
    anchors >= 2,
    `later-page watermarks should stay anchored to their sections, got ${anchors}`,
  )
})

test('maps document meta, html lang, and footer page numbers to native Word parts', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html lang="ja"><head>
      <title>四半期業績レポート</title>
      <meta name="author" content="山田太郎"><!-- public-hygiene: fixture -->
      <meta name="description" content="第2四半期の業績分析">
      <style>footer { position: fixed; bottom: 0; left: 0; right: 0; text-align: center; }</style>
    </head><body>
      <h1>業績サマリー</h1>
      <p>売上高は前年同期比 12% 増となりました。ひらがなとカタカナのテキストです。</p>
      <footer>Page 1 of 3 — 社外秘</footer><!-- public-hygiene: fixture -->
    </body></html>`,
    'doc-meta-lang-pagenum',
  )
  const core = await xml('docProps/core.xml')
  assert.match(core, /<dc:title>四半期業績レポート<\/dc:title>/)
  assert.match(core, /<dc:creator>山田太郎<\/dc:creator>/) // public-hygiene: fixture
  const styles = await xml('word/styles.xml')
  assert.match(styles, /w:eastAsia="ja-JP"/)
  const footer = await xml('word/footer1.xml')
  assert.match(footer, />PAGE</)
  assert.match(footer, />NUMPAGES</)
  assert.match(footer, /社外秘/) // public-hygiene: fixture
})
