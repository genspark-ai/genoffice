import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { PAGE_MARK, parseDocx, saveDocx, type HfParagraph } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const BODY = '<w:p><w:r><w:t>正文</w:t></w:r></w:p>'

async function zipText(bytes: Uint8Array, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file(path)
  return file ? file.async('string') : null
}

describe('first-page / even-page headers & footers', () => {
  it('creates typed parts, references and flags; round-trips through parse', async () => {
    const source = await buildDocx({ bodyXml: BODY })
    const parsed = await parseDocx(source)
    expect(parsed.titlePg).toBe(false)
    expect(parsed.evenAndOddHeaders).toBe(false)
    expect(parsed.headerFirst).toBeNull()

    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      header: { text: '默认页眉' },
      headerFirst: { text: '首页页眉' },
      headerEven: { text: '偶数页页眉' },
      footerEven: { text: '偶 #', pageNumber: true },
      titlePg: true,
      evenAndOddHeaders: true,
    })

    const docXml = (await zipText(saved, 'word/document.xml'))!
    expect(docXml).toContain('<w:titlePg/>')
    expect(docXml).toMatch(/<w:headerReference w:type="first"/)
    expect(docXml).toMatch(/<w:headerReference w:type="even"/)
    expect(docXml).toMatch(/<w:footerReference w:type="even"/)
    const settingsXml = (await zipText(saved, 'word/settings.xml'))!
    expect(settingsXml).toContain('<w:evenAndOddHeaders/>')

    const reparsed = await parseDocx(saved)
    expect(reparsed.titlePg).toBe(true)
    expect(reparsed.evenAndOddHeaders).toBe(true)
    expect(reparsed.headerText).toBe('默认页眉')
    expect(reparsed.headerFirst?.text).toBe('首页页眉')
    expect(reparsed.headerEven?.text).toBe('偶数页页眉')
    expect(reparsed.footerEven?.text).toBe(`偶 ${PAGE_MARK}`)
    expect(reparsed.footerEven?.hasPageNumber).toBe(true)
    expect(reparsed.footerFirst).toBeNull()
  })

  it('gives each new part a distinct filename in a single save', async () => {
    const source = await buildDocx({ bodyXml: BODY })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      header: { text: 'a' },
      headerFirst: { text: 'b' },
      headerEven: { text: 'c' },
    })
    const zip = await JSZip.loadAsync(saved)
    const headers = Object.keys(zip.files).filter((n) => /^word\/header\d+\.xml$/.test(n))
    expect(headers.sort()).toEqual(['word/header1.xml', 'word/header2.xml', 'word/header3.xml'])
  })

  it('rewrites only the edited variant; the others stay byte-identical', async () => {
    const source = await buildDocx({ bodyXml: BODY })
    const parsed = await parseDocx(source)
    const withBoth = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      header: { text: '默认' },
      headerFirst: { text: '首页' },
      titlePg: true,
    })
    const p2 = await parseDocx(withBoth)
    const firstPartBefore = await zipText(withBoth, 'word/header2.xml')
    const edited = await saveDocx(p2, [{ kind: 'original', docxIndex: 0 }], {
      header: { text: '默认已改' },
    })
    expect(await zipText(edited, 'word/header2.xml')).toBe(firstPartBefore)
    const p3 = await parseDocx(edited)
    expect(p3.headerText).toBe('默认已改')
    expect(p3.headerFirst?.text).toBe('首页')
    expect(p3.titlePg).toBe(true)
  })

  it('removes titlePg and evenAndOddHeaders when toggled off', async () => {
    const source = await buildDocx({ bodyXml: BODY })
    const parsed = await parseDocx(source)
    const on = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      titlePg: true,
      evenAndOddHeaders: true,
    })
    const p2 = await parseDocx(on)
    const off = await saveDocx(p2, [{ kind: 'original', docxIndex: 0 }], {
      titlePg: false,
      evenAndOddHeaders: false,
    })
    const p3 = await parseDocx(off)
    expect(p3.titlePg).toBe(false)
    expect(p3.evenAndOddHeaders).toBe(false)
    expect(await zipText(off, 'word/document.xml')).not.toContain('<w:titlePg/>')
  })

  it('parses even-header table cells with one line per cell paragraph (same as default)', async () => {
    const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    const hdrXml = (tblCells: string) =>
      XML_DECL +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="500"/><w:gridCol w:w="7500"/></w:tblGrid>' +
      `<w:tr>${tblCells}</w:tr></w:tbl></w:hdr>`
    const cells =
      '<w:tc><w:tcPr><w:tcW w:w="500" w:type="dxa"/><w:shd w:val="clear" w:fill="C00000"/></w:tcPr><w:p/></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="7500" w:type="dxa"/></w:tcPr>' +
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Title line</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Subtitle line</w:t></w:r></w:p></w:tc>'
    const headerType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels:
        '<Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
        '<Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>',
      extraParts: [
        { path: 'word/header1.xml', xml: hdrXml(cells), contentType: headerType },
        { path: 'word/header2.xml', xml: hdrXml(cells), contentType: headerType },
      ],
      sectPrExtra:
        '<w:headerReference w:type="default" r:id="rId60"/>' +
        '<w:headerReference w:type="even" r:id="rId61"/>',
    })
    const parsed = await parseDocx(bytes)
    const cellShape = (row: HfParagraph) =>
      row.cells!.map((c) => ({
        fill: c.fill,
        texts: c.paras.map((rs) => rs.map((r) => r.text).join('')),
      }))
    const evenRow = parsed.headerEven!.paras[0]
    expect(cellShape(evenRow)).toEqual([
      { fill: 'C00000', texts: [''] },
      { fill: undefined, texts: ['Title line', 'Subtitle line'] },
    ])
    expect(cellShape(evenRow)).toEqual(cellShape(parsed.headerParas![0]))
    expect(evenRow.cells![1].paras[0][0]).toMatchObject({ text: 'Title line', bold: true })
  })

  it('keeps titlePg before w:docGrid in schema order', async () => {
    const source = await buildDocx({ bodyXml: BODY })
    // inject a docGrid into the trailing sectPr
    const zip = await JSZip.loadAsync(source)
    const docXml = (await zip.file('word/document.xml')!.async('string')).replace(
      '</w:sectPr>',
      '<w:docGrid w:linePitch="360"/></w:sectPr>',
    )
    zip.file('word/document.xml', docXml)
    const parsed = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], { titlePg: true })
    const out = (await zipText(saved, 'word/document.xml'))!
    expect(out.indexOf('<w:titlePg/>')).toBeGreaterThan(-1)
    expect(out.indexOf('<w:titlePg/>')).toBeLessThan(out.indexOf('<w:docGrid'))
  })
})

const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const FOOTER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const hdrPart = (inner: string) => `${XML_DECL}<w:hdr ${W_NS}>${inner}</w:hdr>`
const ftrPart = (inner: string) => `${XML_DECL}<w:ftr ${W_NS}>${inner}</w:ftr>`
const settingsPart = (inner: string) => `${XML_DECL}<w:settings ${W_NS}>${inner}</w:settings>`

describe('non-schema and non-self-closing variants (POI corpus)', () => {
  it('treats w:type="odd" header/footer references as the default part', async () => {
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels:
        '<Relationship Id="rId70" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
        '<Relationship Id="rId71" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: hdrPart('<w:p><w:r><w:t>EVEN_PAGE_HEADER</w:t></w:r></w:p>'),
          contentType: HEADER_TYPE,
        },
        {
          path: 'word/header2.xml',
          xml: hdrPart('<w:p><w:r><w:t>ODD_PAGE_HEADER</w:t></w:r></w:p>'),
          contentType: HEADER_TYPE,
        },
      ],
      sectPrExtra:
        '<w:headerReference w:type="even" r:id="rId70"/>' +
        '<w:headerReference w:type="odd" r:id="rId71"/>',
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.headerText).toBe('ODD_PAGE_HEADER')
    expect(parsed.headerEven?.text).toBe('EVEN_PAGE_HEADER')

    // editing the default header rewrites the odd part instead of adding a duplicate reference
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      header: { text: 'EDITED_DEFAULT' },
    })
    expect((await zipText(saved, 'word/document.xml'))!.match(/<w:headerReference/g)).toHaveLength(
      2,
    )
    expect(await zipText(saved, 'word/header2.xml')).toContain('EDITED_DEFAULT')
    expect((await parseDocx(saved)).headerText).toBe('EDITED_DEFAULT')
  })

  it('reads evenAndOddHeaders and titlePg written as paired tags or with w:val', async () => {
    const build = (settingsXml: string, titlePgTag: string) =>
      buildDocx({
        bodyXml: BODY,
        sectPrExtra: titlePgTag,
        extraParts: [
          { path: 'word/settings.xml', xml: settingsPart(settingsXml), contentType: SETTINGS_TYPE },
        ],
      })
    const on = await parseDocx(
      await build(
        '<w:evenAndOddHeaders></w:evenAndOddHeaders>',
        '<w:titlePg w:val="true"></w:titlePg>',
      ),
    )
    expect(on.evenAndOddHeaders).toBe(true)
    expect(on.titlePg).toBe(true)
    const off = await parseDocx(
      await build('<w:evenAndOddHeaders w:val="false"/>', '<w:titlePg w:val="0"/>'),
    )
    expect(off.evenAndOddHeaders).toBe(false)
    expect(off.titlePg).toBe(false)
  })

  it('flattens nested layout tables into their cell (banner shading and text kept)', async () => {
    // banner header: outer cell holds a nested single-cell table carrying the
    // fill and the centered title, followed by the mandatory empty paragraph
    const footerXml = ftrPart(
      '<w:tbl><w:tblGrid><w:gridCol w:w="8000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>' +
        '<w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>' +
        '<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr>' +
        '<w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/><w:shd w:val="clear" w:fill="000066"/></w:tcPr>' +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Banner title</w:t></w:r></w:p>' +
        '</w:tc></w:tr></w:tbl><w:p/></w:tc>' +
        '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>' +
        '</w:tr></w:tbl>',
    )
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels:
        '<Relationship Id="rId72" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
      extraParts: [{ path: 'word/footer1.xml', xml: footerXml, contentType: FOOTER_TYPE }],
      sectPrExtra: '<w:footerReference w:type="default" r:id="rId72"/>',
    })
    const parsed = await parseDocx(bytes)
    const row = parsed.footerParas![0]
    expect(row.cells![0]).toMatchObject({ fill: '000066', align: 'center' })
    // nested cell text kept; the trailing empty paragraph after the nested table is trimmed
    expect(row.cells![0].paras.map((rs) => rs.map((r) => r.text).join(''))).toEqual([
      'Banner title',
    ])
    expect(parsed.footerText).toBe('Banner title')
  })

  it('emits a floating table (w:tblpPr) after its anchor paragraph, matching Word order', async () => {
    const footerXml = ftrPart(
      '<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="text" w:tblpY="1"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr>' +
        '<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>' +
        '<w:p><w:r><w:t>Page band</w:t></w:r></w:p></w:tc>' +
        '</w:tr></w:tbl>' +
        '<w:p><w:r><w:t>Simple footer text</w:t></w:r></w:p>',
    )
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraRels:
        '<Relationship Id="rId73" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
      extraParts: [{ path: 'word/footer1.xml', xml: footerXml, contentType: FOOTER_TYPE }],
      sectPrExtra: '<w:footerReference w:type="default" r:id="rId73"/>',
    })
    const parsed = await parseDocx(bytes)
    const paras = parsed.footerParas!
    expect(paras[0].runs.map((r) => r.text).join('')).toBe('Simple footer text')
    expect(paras[1].cells![0].paras[0][0].text).toBe('Page band')
  })
})

describe('header style-inherited rtl', () => {
  it('hf runs pick the Cs property set when their paragraph style sets w:rtl', async () => {
    const headerXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:pPr><w:pStyle w:val="HeaderAr"/></w:pPr>' +
      '<w:r><w:rPr><w:bCs/><w:szCs w:val="36"/><w:sz w:val="20"/></w:rPr><w:t>مرحبا</w:t></w:r>' +
      '</w:p></w:hdr>'
    const source = await buildDocx({
      bodyXml: BODY,
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="HeaderAr"><w:name w:val="Header Ar"/>' +
        '<w:basedOn w:val="Normal"/><w:rPr><w:rtl/></w:rPr></w:style>',
      extraRels:
        '<Relationship Id="rId21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId21"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: headerXml,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
    })
    const parsed = await parseDocx(source)
    const run = parsed.headerParas![0].runs[0]
    // style rtl reaches hf runs: szCs/bCs win over sz (cs property set)
    expect(run).toMatchObject({ text: 'مرحبا', cs: true, bold: true, sizeHalfPoints: 36 })
  })
})
