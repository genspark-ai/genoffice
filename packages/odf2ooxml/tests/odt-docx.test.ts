import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { convertOdtToDocx, OdfConvertError } from '../src/index'

const NAMESPACES =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
  ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
  ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' +
  ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
  ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' +
  ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
  ' xmlns:xlink="http://www.w3.org/1999/xlink"'

/** 1Ã—1 red PNG */
const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

function contentXml(body: string, automaticStyles = ''): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.2">` +
    `<office:automatic-styles>${automaticStyles}</office:automatic-styles>` +
    `<office:body><office:text>${body}</office:text></office:body></office:document-content>`
  )
}

function stylesXml(named: string, automatic = ''): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.2">` +
    `<office:styles>${named}</office:styles>` +
    `<office:automatic-styles>${automatic}</office:automatic-styles>` +
    `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="PM1"/></office:master-styles>` +
    '</office:document-styles>'
  )
}

const PAGE_A4 =
  '<style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"' +
  ' fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/></style:page-layout>'

async function buildOdt(content: string, styles?: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text')
  zip.file('content.xml', content)
  if (styles) zip.file('styles.xml', styles)
  return zip.generateAsync({ type: 'uint8array' })
}

async function readDocxEntry(docx: Uint8Array, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(docx)
  const entry = zip.file(path)
  return entry ? entry.async('string') : null
}

async function docxHasMedia(docx: Uint8Array, suffix: string): Promise<boolean> {
  const zip = await JSZip.loadAsync(docx)
  return Object.keys(zip.files).some(
    (path) => path.startsWith('word/media/') && path.endsWith(suffix),
  )
}

describe('convertOdtToDocx', () => {
  it('maps headings, character formatting and hyperlinks', async () => {
    const body =
      '<text:h text:style-name="Heading_20_1" text:outline-level="1">Quarterly Report</text:h>' +
      '<text:p>plain <text:span text:style-name="T1">bold</text:span> tail</text:p>' +
      '<text:p>see <text:a xlink:type="simple" xlink:href="https://example.com/doc">the site</text:a></text:p>'
    const named =
      '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-style="italic"/></style:style>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body), stylesXml(named)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('w:pStyle w:val="Heading1"')
    expect(document).toContain('Quarterly Report')
    expect(document).toContain('<w:b/>')
    expect(document).toContain('<w:i/>')
    expect(document).toContain('bold')
    expect(document).toContain('<w:hyperlink')
    const rels = (await readDocxEntry(result.docx, 'word/_rels/document.xml.rels')) ?? ''
    expect(rels).toContain('https://example.com/doc')
  })

  it('carries character spacing and width from a text style into the run', async () => {
    const body =
      '<text:p>plain <text:span text:style-name="Twide">wide</text:span> ' +
      '<text:span text:style-name="Ttight">tight</text:span> tail</text:p>'
    const named =
      '<style:style style:name="Twide" style:family="text"><style:text-properties fo:letter-spacing="0.05cm" style:text-scale="150%"/></style:style>' +
      '<style:style style:name="Ttight" style:family="text"><style:text-properties fo:letter-spacing="-0.02cm"/></style:style>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body), stylesXml(named)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    // ODF spells them fo:letter-spacing and style:text-scale; docx spells the
    // same two as w:spacing (twips, signed) and w:w (percent)
    expect(document).toContain('<w:spacing w:val="28"/>') // 0.05cm = 18000 EMU
    expect(document).toContain('<w:spacing w:val="-11"/>') // -0.02cm = -7200 EMU
    expect(document).toContain('<w:w w:val="150"/>')
    expect(document).toContain('wide')
    expect(document).toContain('tight')
  })

  it('resolves the parent style chain and character properties from styles.xml', async () => {
    const body = '<text:p text:style-name="P1">styled</text:p>'
    const named =
      '<style:style style:name="Base" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>' +
      '<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Base"><style:text-properties fo:color="#C9211E"/></style:style>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body), stylesXml(named)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('<w:jc w:val="center"/>')
    expect(document).toContain('<w:color w:val="C9211E"/>')
  })

  it('keeps tabs, line breaks and repeated spaces', async () => {
    const body = '<text:p>a<text:tab/>b<text:line-break/>c<text:s text:c="3"/>d</text:p>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('<w:tab/>')
    expect(document).toContain('<w:br/>')
    expect(document).toContain('c   d')
  })

  it('treats a negative space count as a single space instead of throwing', async () => {
    const body = '<text:p>a<text:s text:c="-5"/>b</text:p>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    const text = [...document.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
      .map((match) => match[1])
      .join('')
    expect(text).toBe('a b')
  })

  it('keeps an ordinary line-height percentage but ignores one beyond 1000%', async () => {
    const automatic =
      '<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:line-height="200%"/></style:style>' +
      '<style:style style:name="P2" style:family="paragraph"><style:paragraph-properties fo:line-height="999999999999999999999%"/></style:style>'
    const body =
      '<text:p text:style-name="P1">normal</text:p><text:p text:style-name="P2">huge</text:p>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body, automatic)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    const lines = [...document.matchAll(/w:line="([^"]*)"/g)].map((match) => match[1])
    expect(lines).toContain('480')
    // OOXML w:line is an integer, so no value may come out in exponent notation
    expect(lines.every((value) => /^\d+$/.test(value))).toBe(true)
  })

  it('fills a paragraph that asks for a background', async () => {
    const body =
      '<text:p text:style-name="Pfill">highlighted</text:p>' +
      '<text:p text:style-name="Pclear">plain</text:p>'
    const automatic =
      '<style:style style:name="Pfill" style:family="paragraph"><style:paragraph-properties fo:background-color="#fff2a8"/></style:style>' +
      // "transparent" is ODF's default, so it asks for no fill of its own
      '<style:style style:name="Pclear" style:family="paragraph"><style:paragraph-properties fo:background-color="transparent"/></style:style>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body, automatic)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFF2A8"/>')
    // one shaded paragraph, not two
    expect(document.match(/<w:shd /g) ?? []).toHaveLength(1)
  })

  it('keeps a paragraph with the next one and its lines together', async () => {
    const body =
      '<text:p text:style-name="Pkeep">intro</text:p>' +
      '<text:p text:style-name="Pcancel">body</text:p>'
    const automatic =
      '<style:style style:name="Pkeep" style:family="paragraph"><style:paragraph-properties fo:keep-with-next="always" fo:keep-together="always"/></style:style>' +
      // "auto" is ODF's default, but it still has to cancel an inherited "always"
      '<style:style style:name="Pcancel" style:family="paragraph" style:parent-style-name="Pkeep"><style:paragraph-properties fo:keep-with-next="auto" fo:keep-together="auto"/></style:style>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body, automatic)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('<w:keepNext/>')
    expect(document).toContain('<w:keepLines/>')
    // only the paragraph that asked for it carries either element
    expect(document.match(/<w:keepNext\/>/g) ?? []).toHaveLength(1)
    expect(document.match(/<w:keepLines\/>/g) ?? []).toHaveLength(1)
  })

  it('imports lists with per-level bullet and number formats', async () => {
    const body =
      '<text:list text:style-name="L1"><text:list-item><text:p>first</text:p>' +
      '<text:list text:style-name="L1"><text:list-item><text:p>sub</text:p></text:list-item></text:list>' +
      '</text:list-item><text:list-item><text:p>second</text:p></text:list-item></text:list>'
    const listStyle =
      '<text:list-style style:name="L1">' +
      '<text:list-level-style-bullet text:level="1" style:bullet-char="â€¢">' +
      '<style:list-level-properties text:space-before="0.5cm" text:min-label-width="0.4cm"/></text:list-level-style-bullet>' +
      '<text:list-level-style-number text:level="2" style:num-format="1" style:num-suffix=".">' +
      '<style:list-level-properties text:space-before="1cm" text:min-label-width="0.4cm"/></text:list-level-style-number>' +
      '</text:list-style>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body, listStyle)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('<w:numPr>')
    const numbering = (await readDocxEntry(result.docx, 'word/numbering.xml')) ?? ''
    expect(numbering).toContain('w:numFmt w:val="bullet"')
    expect(numbering).toContain('w:numFmt w:val="decimal"')
    expect(numbering).toContain('w:lvlText w:val="%2."')
  })

  it('carries table:align into w:jc and leaves a margins table flush', async () => {
    // ODF places a table with table:align; "margins" is the default full-width
    // behaviour, while center/right have a w:jc of their own
    const automatic =
      '<style:style style:name="TC" style:family="table"><style:table-properties table:align="center"/></style:style>' +
      '<style:style style:name="TR" style:family="table"><style:table-properties table:align="right"/></style:style>' +
      '<style:style style:name="TM" style:family="table"><style:table-properties table:align="margins"/></style:style>'
    const table = (style: string, text: string): string =>
      `<table:table table:name="${text}" table:style-name="${style}"><table:table-column/>` +
      `<table:table-row><table:table-cell><text:p>${text}</text:p></table:table-cell></table:table-row></table:table>`
    const body = table('TC', 'centred') + table('TR', 'right') + table('TM', 'margin')
    const result = await convertOdtToDocx(await buildOdt(contentXml(body, automatic)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    // w:jc follows tblW, as CT_TblPrBase orders it
    expect(document).toMatch(/<w:tblW[^>]*\/><w:jc w:val="center"\/><w:tblBorders>/)
    expect(document).toMatch(/<w:tblW[^>]*\/><w:jc w:val="right"\/><w:tblBorders>/)
    expect(document.match(/<w:jc /g) ?? []).toHaveLength(2)
  })

  it('carries a cell style style:vertical-align into w:vAlign', async () => {
    const automatic =
      '<style:style style:name="cm" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffff99" style:vertical-align="middle"/></style:style>' +
      '<style:style style:name="cb" style:family="table-cell"><style:table-cell-properties style:vertical-align="bottom"/></style:style>' +
      '<style:style style:name="ct" style:family="table-cell"><style:table-cell-properties style:vertical-align="top"/></style:style>' +
      // the fo: name is not what a table cell carries, so it must not be read
      '<style:style style:name="cf" style:family="table-cell"><style:table-cell-properties fo:vertical-align="middle"/></style:style>'
    const body =
      '<table:table table:name="T"><table:table-column table:number-columns-repeated="4"/>' +
      '<table:table-row>' +
      '<table:table-cell table:style-name="cm"><text:p>middle</text:p></table:table-cell>' +
      '<table:table-cell table:style-name="cb"><text:p>bottom</text:p></table:table-cell>' +
      '<table:table-cell table:style-name="ct"><text:p>top</text:p></table:table-cell>' +
      '<table:table-cell table:style-name="cf"><text:p>foreign</text:p></table:table-cell>' +
      '</table:table-row></table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body, automatic)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    // middle is Word's "center"; bottom keeps its name; top is the default
    expect(document).toContain('<w:vAlign w:val="center"/>')
    expect(document).toContain('<w:vAlign w:val="bottom"/>')
    expect(document.match(/<w:vAlign /g) ?? []).toHaveLength(2)
    // vAlign follows shd in CT_TcPr
    expect(document).toContain(
      '<w:shd w:val="clear" w:color="auto" w:fill="FFFF99"/><w:vAlign w:val="center"/>',
    )
  })

  it('imports a table with widths, spans and a cell fill', async () => {
    const automatic =
      '<style:style style:name="co1" style:family="table-column"><style:column-properties style:column-width="3cm"/></style:style>' +
      '<style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffff99"/></style:style>'
    const body =
      '<table:table table:name="T" table:style-name="T1">' +
      '<table:table-column table:style-name="co1" table:number-columns-repeated="2"/>' +
      '<table:table-row><table:table-cell table:style-name="ce1" office:value-type="string" table:number-columns-spanned="2"><text:p>wide</text:p></table:table-cell></table:table-row>' +
      '<table:table-row><table:table-cell><text:p>left</text:p></table:table-cell><table:table-cell><text:p>right</text:p></table:table-cell></table:table-row>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body, automatic)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('<w:tbl>')
    expect(document).toContain('<w:gridSpan w:val="2"/>')
    expect(document).toContain('w:fill="FFFF99"')
    expect(document).toContain(
      '<w:tblGrid><w:gridCol w:w="1701"/><w:gridCol w:w="1701"/></w:tblGrid>',
    )
  })

  it('keeps every cell of a table that declares fewer columns than it paints', async () => {
    // The column list is optional; the single-column fallback stopped the grid
    // at one slot, so every cell after the first disappeared with no tally.
    const body =
      '<table:table table:name="T">' +
      '<table:table-row>' +
      '<table:table-cell><text:p>AAACELL</text:p></table:table-cell>' +
      '<table:table-cell><text:p>BBB CELL</text:p></table:table-cell>' +
      '<table:table-cell><text:p>CCC CELL</text:p></table:table-cell>' +
      '</table:table-row>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    for (const text of ['AAACELL', 'BBB CELL', 'CCC CELL']) expect(document).toContain(text)
    expect((document.match(/<w:gridCol /g) ?? []).length).toBe(3)
    expect(result.dropped).toEqual([])
  })

  it('keeps a span that the declared column list cannot hold', async () => {
    // The same narrowing also clamped a cell's colspan to the fallback's one
    // column, so a merged cell came back unmerged.
    const body =
      '<table:table table:name="T">' +
      '<table:table-row><table:table-cell table:number-columns-spanned="3"><text:p>wide</text:p></table:table-cell></table:table-row>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('<w:gridSpan w:val="3"/>')
    expect((document.match(/<w:gridCol /g) ?? []).length).toBe(3)
  })

  it('does not widen a declared grid for a fully merged first row', async () => {
    // LibreOffice's shape for a 3-column table whose first row is merged: the
    // anchored cell plus one covered placeholder per merged-in column. The span
    // does not add grid positions of its own â€” the placeholders are its slots.
    const body =
      '<table:table table:name="T">' +
      '<table:table-column table:number-columns-repeated="3"/>' +
      '<table:table-row>' +
      '<table:table-cell table:number-columns-spanned="3"><text:p>merged</text:p></table:table-cell>' +
      '<table:covered-table-cell/><table:covered-table-cell/>' +
      '</table:table-row>' +
      '<table:table-row>' +
      '<table:table-cell><text:p>cell x</text:p></table:table-cell>' +
      '<table:table-cell><text:p>cell y</text:p></table:table-cell>' +
      '<table:table-cell><text:p>cell z</text:p></table:table-cell>' +
      '</table:table-row>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect((document.match(/<w:gridCol /g) ?? []).length).toBe(3)
    expect(document).toContain('<w:gridSpan w:val="3"/>')
    for (const text of ['merged', 'cell x', 'cell y', 'cell z']) expect(document).toContain(text)
  })

  it('does not widen a declared grid for a cell merging two of its columns', async () => {
    const body =
      '<table:table table:name="T">' +
      '<table:table-column table:number-columns-repeated="3"/>' +
      '<table:table-row>' +
      '<table:table-cell table:number-columns-spanned="2"><text:p>pair</text:p></table:table-cell>' +
      '<table:covered-table-cell/>' +
      '<table:table-cell><text:p>side</text:p></table:table-cell>' +
      '</table:table-row>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect((document.match(/<w:gridCol /g) ?? []).length).toBe(3)
    expect(document).toContain('<w:gridSpan w:val="2"/>')
    expect(document).toContain('pair')
    expect(document).toContain('side')
  })

  it('counts a vertical merge placeholder as the position it covers', async () => {
    // The second row's covered cell stands for the slot under the row span; it
    // must not shift that row's own cell into a fourth column.
    const body =
      '<table:table table:name="T">' +
      '<table:table-column table:number-columns-repeated="2"/>' +
      '<table:table-row>' +
      '<table:table-cell table:number-rows-spanned="2"><text:p>tall</text:p></table:table-cell>' +
      '<table:table-cell><text:p>top right</text:p></table:table-cell>' +
      '</table:table-row>' +
      '<table:table-row>' +
      '<table:covered-table-cell/>' +
      '<table:table-cell><text:p>bottom right</text:p></table:table-cell>' +
      '</table:table-row>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect((document.match(/<w:gridCol /g) ?? []).length).toBe(2)
    expect(document).toContain('bottom right')
  })

  it('repeats a row table:number-rows-repeated asks for', async () => {
    const body =
      '<table:table table:name="T">' +
      '<table:table-column/>' +
      '<table:table-row table:number-rows-repeated="3"><table:table-cell><text:p>once</text:p></table:table-cell></table:table-row>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    // the attribute is shorthand for three copies of the row, not one row
    expect(document.match(/<w:tr>/g) ?? []).toHaveLength(3)
    expect(document.match(/once/g) ?? []).toHaveLength(3)
    // nothing was lost, so nothing is reported as dropped
    expect(result.dropped).toEqual([])
  })

  it('bounds the rows a table repeats, not only one row element', async () => {
    // the per-element cap alone still multiplies: 50 elements each claiming
    // 1048576 copies over 64 columns was 51,200 rows / 3.3M cells / 380 MB
    const row =
      '<table:table-row table:number-rows-repeated="1048576">' +
      '<table:table-cell table:number-columns-repeated="64"><text:p>f</text:p></table:table-cell>' +
      '</table:table-row>'
    const body =
      '<table:table table:name="T">' +
      '<table:table-column table:number-columns-repeated="64"/>' +
      row.repeat(50) +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    const rows = (document.match(/<w:tr>/g) ?? []).length
    // 64 columns divide the 65,536-slot ceiling into 1024 repeat rows; every row
    // element keeps its first copy and the repeats share that one table budget
    expect(rows).toBeLessThanOrEqual(50 + 1024)
    expect(rows).toBeGreaterThan(1024)
    expect(result.dropped).toEqual([{ what: 'repeated-table-rows', count: 50 * 1048576 - rows }])
  })

  it('spends the repeat budget across the whole table, not per row element', async () => {
    // one narrow column: the slot ceiling does not bind, the 4096-row count does
    const row =
      '<table:table-row table:number-rows-repeated="1048576">' +
      '<table:table-cell><text:p>f</text:p></table:table-cell></table:table-row>'
    const body =
      '<table:table table:name="T"><table:table-column/>' + row.repeat(6) + '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    const rows = (document.match(/<w:tr>/g) ?? []).length
    expect(rows).toBeLessThanOrEqual(6 + 4096)
    expect(rows).toBeGreaterThan(1024)
    expect(result.dropped).toEqual([{ what: 'repeated-table-rows', count: 6 * 1048576 - rows }])
  })

  it('bounds a row repeat like the column reader does', async () => {
    const body =
      '<table:table table:name="T">' +
      '<table:table-column/>' +
      '<table:table-row table:number-rows-repeated="100000"><table:table-cell><text:p>filler</text:p></table:table-cell></table:table-row>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    // readColumns caps a repeat at 1024; the row reader uses the same clamp
    expect(document.match(/<w:tr>/g) ?? []).toHaveLength(1024)
    // and the copies the clamp and the budget cut are reported, not lost silently
    expect(result.dropped).toEqual([{ what: 'repeated-table-rows', count: 100000 - 1024 }])
  })

  it('bounds a row span to the rows that exist', async () => {
    // table:number-rows-spanned is producer-controlled: an unclamped 200000
    // allocated that many empty rows, and Infinity never returned.
    const body =
      '<table:table table:name="T">' +
      '<table:table-row><table:table-cell table:number-rows-spanned="200000"><text:p>tall</text:p></table:table-cell></table:table-row>' +
      '<table:table-row><table:table-cell><text:p>next</text:p></table:table-cell></table:table-row>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect((document.match(/<w:tr>/g) ?? []).length).toBe(2)
    expect(document).toContain('<w:vMerge w:val="restart"/>')
    expect(document).toContain('<w:vMerge/>')
  })

  it('marks table:table-header-rows so Word repeats the heading on each page', async () => {
    // ODF splits the two row containers: header-rows repeat at the top of a
    // page, rows do not. readRows flattens them, so the distinction has to
    // ride along as w:tblHeader or the heading scrolls away on page two.
    const body =
      '<table:table table:name="T"><table:table-column/>' +
      '<table:table-header-rows>' +
      '<table:table-row><table:table-cell><text:p>head one</text:p></table:table-cell></table:table-row>' +
      '<table:table-row><table:table-cell><text:p>head two</text:p></table:table-cell></table:table-row>' +
      '</table:table-header-rows>' +
      '<table:table-rows>' +
      '<table:table-row><table:table-cell><text:p>body</text:p></table:table-cell></table:table-row>' +
      '</table:table-rows>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    const rows = [...document.matchAll(/<w:tr>([\s\S]*?)<\/w:tr>/g)].map((match) => match[1])
    expect(rows).toHaveLength(3)
    // both header rows repeat; the ordinary row does not
    expect(rows.map((row) => row.includes('<w:tblHeader/>'))).toEqual([true, true, false])
    // w:trPr is the row's first child, so the marker precedes the first cell
    expect(rows[0].startsWith('<w:trPr><w:tblHeader/></w:trPr><w:tc>')).toBe(true)
  })

  it('does not mark a heading when the table keeps its rows in one container', async () => {
    const body =
      '<table:table table:name="T"><table:table-column/>' +
      '<table:table-row><table:table-cell><text:p>only</text:p></table:table-cell></table:table-row>' +
      '</table:table>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).not.toContain('w:tblHeader')
    expect(document).toContain('<w:tr><w:tc>')
  })

  it('carries the page geometry into the section properties', async () => {
    const body = '<text:p>page</text:p>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body), stylesXml('', PAGE_A4)))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    // 21cm â†’ 11906 twips, 29.7cm â†’ 16838 twips, 2cm â†’ 1134
    expect(document).toContain('w:pgSz w:w="11906" w:h="16838"')
    expect(document).toContain('w:pgMar w:top="1134"')
  })

  it('imports framed and inline images', async () => {
    const body =
      '<text:p>before</text:p>' +
      `<draw:frame svg:width="4cm" svg:height="3cm"><draw:image xlink:href="Pictures/red.png"/></draw:frame>` +
      '<text:p>with <draw:frame svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="2cm"><draw:image xlink:href="Pictures/red.png"/></draw:frame> inline</text:p>'
    const zip = new JSZip()
    zip.file('content.xml', contentXml(body))
    zip.file('Pictures/red.png', PNG_BASE64, { base64: true })
    const result = await convertOdtToDocx(await zip.generateAsync({ type: 'uint8array' }))
    expect(await docxHasMedia(result.docx, '.png')).toBe(true)
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('<w:drawing>')
  })

  it('drops an inline image whose base64 payload is not decodable', async () => {
    const body =
      '<text:p>before<draw:frame svg:width="2cm" svg:height="2cm"><draw:image><office:binary-data>not-base64!!!</office:binary-data></draw:image></draw:frame>after</text:p>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    expect(result.dropped).toContainEqual({ what: 'images', count: 1 })
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('before')
    expect(document).toContain('after')
    expect(document).not.toContain('<w:drawing>')
  })

  it('does not blame the first table for an image a paragraph already emitted', async () => {
    // The paragraph's inline image is collected in the shared walk context; a
    // table cell only has its own images to account for, so a leftover from
    // the paragraph above would be counted as dropped here.
    const body =
      '<text:p>with <draw:frame svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="2cm"><draw:image xlink:href="Pictures/red.png"/></draw:frame> inline</text:p>' +
      '<table:table table:name="T">' +
      '<table:table-row><table:table-cell><text:p>left</text:p></table:table-cell></table:table-row>' +
      '</table:table>'
    const zip = new JSZip()
    zip.file('content.xml', contentXml(body))
    zip.file('Pictures/red.png', PNG_BASE64, { base64: true })
    const result = await convertOdtToDocx(await zip.generateAsync({ type: 'uint8array' }))
    // the image really is in the package, so nothing was dropped
    expect(await docxHasMedia(result.docx, '.png')).toBe(true)
    expect(result.dropped).not.toContainEqual(expect.objectContaining({ what: 'images' }))
    // and the table is still imported
    expect((await readDocxEntry(result.docx, 'word/document.xml')) ?? '').toContain('<w:tbl>')
  })

  it('still counts an image a table cell cannot hold', async () => {
    const body =
      '<table:table table:name="T">' +
      '<table:table-row><table:table-cell><text:p><draw:frame svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="2cm"><draw:image xlink:href="Pictures/red.png"/></draw:frame></text:p></table:table-cell></table:table-row>' +
      '</table:table>'
    const zip = new JSZip()
    zip.file('content.xml', contentXml(body))
    zip.file('Pictures/red.png', PNG_BASE64, { base64: true })
    const result = await convertOdtToDocx(await zip.generateAsync({ type: 'uint8array' }))
    expect(result.dropped).toContainEqual({ what: 'images', count: 1 })
    expect(await docxHasMedia(result.docx, '.png')).toBe(false)
  })

  it('counts constructs it cannot represent instead of guessing', async () => {
    const body =
      '<text:p>keep <text:note text:id="n1" text:note-class="footnote"><text:note-body><text:p>lost</text:p></text:note-body></text:note></text:p>' +
      '<draw:rect svg:width="1cm" svg:height="1cm"/>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    expect(result.dropped).toContainEqual({ what: 'footnotes', count: 1 })
    expect(result.dropped).toContainEqual({ what: 'shapes', count: 1 })
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('keep')
    expect(document).not.toContain('lost')
  })

  it('flattens a text frame into the flow and says so', async () => {
    const body =
      '<draw:frame svg:width="8cm" svg:height="2cm"><draw:text-box><text:p>boxed words</text:p></draw:text-box></draw:frame>'
    const result = await convertOdtToDocx(await buildOdt(contentXml(body)))
    expect(result.dropped).toContainEqual({ what: 'textboxes-flattened', count: 1 })
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('boxed words')
  })

  it('accepts .ott templates (a -template mimetype, same body shape)', async () => {
    const body = '<text:h text:outline-level="1">From a template</text:h><text:p>body</text:p>'
    const zip = new JSZip()
    zip.file('mimetype', 'application/vnd.oasis.opendocument.text-template')
    zip.file('content.xml', contentXml(body))
    const result = await convertOdtToDocx(await zip.generateAsync({ type: 'uint8array' }))
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).toContain('From a template')
  })

  it('rejects bytes that are not an ODF package and zips with no text body', async () => {
    await expect(convertOdtToDocx(new Uint8Array([1, 2, 3]))).rejects.toMatchObject({
      code: 'unreadable',
    })
    await expect(
      convertOdtToDocx(
        await buildOdt(
          contentXml('<text:p>x</text:p>').replace('office:text', 'office:wordprocessing'),
        ),
      ),
    ).rejects.toBeInstanceOf(OdfConvertError)
  })
})
