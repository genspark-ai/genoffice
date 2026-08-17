import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { PAGE_MARK, parseDocx, saveDocx, TOTAL_PAGES_MARK, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/** rich header/footer paragraphs: save (paras) -> reparse (headerParas) round trip */

async function base() {
  const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>正文</w:t></w:r></w:p>' })
  const parsed = await parseDocx(bytes)
  const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: parsed.blocks[0].docxIndex! }]
  return { parsed, saveBlocks }
}

describe('rich header / footer', () => {
  it('keeps an RTL header right-to-left and keeps its alignment side', async () => {
    const { parsed, saveBlocks } = await base()
    const saved = await saveDocx(parsed, saveBlocks, {
      header: {
        text: '',
        paras: [
          { bidi: true, align: 'left', runs: [{ text: 'ترويسة عربية' }] },
          { bidi: true, runs: [{ text: 'بدون محاذاة' }] },
        ],
      },
    })
    const reparsed = await parseDocx(saved)
    expect(reparsed.headerParas![0]).toMatchObject({ bidi: true, align: 'left' })
    expect(reparsed.headerParas![1]?.bidi).toBe(true)
  })

  it('saves multi-paragraph styled header and reparses it as headerParas', async () => {
    const { parsed, saveBlocks } = await base()
    const saved = await saveDocx(parsed, saveBlocks, {
      header: {
        text: '',
        paras: [
          { align: 'right', runs: [{ text: '机密', bold: true, color: 'FF0000' }] },
          { runs: [{ text: '年度报告 ' }, { text: '2026', italic: true }] },
        ],
      },
    })
    const reparsed = await parseDocx(saved)
    expect(reparsed.headerParas).toHaveLength(2)
    expect(reparsed.headerParas![0].align).toBe('right')
    expect(reparsed.headerParas![0].runs[0]).toMatchObject({
      text: '机密',
      bold: true,
      color: 'FF0000',
    })
    expect(reparsed.headerParas![1].runs.map((r) => r.text).join('')).toBe('年度报告 2026')
    expect(reparsed.headerText).toBe('机密年度报告 2026')
  })

  it('footer paras with # keep a real PAGE field', async () => {
    const { parsed, saveBlocks } = await base()
    const saved = await saveDocx(parsed, saveBlocks, {
      footer: {
        text: '',
        pageNumber: true,
        paras: [{ align: 'center', runs: [{ text: '第 # 页' }] }],
      },
    })
    const reparsed = await parseDocx(saved)
    expect(reparsed.footerHasPageNumber).toBe(true)
    expect(reparsed.footerParas![0].runs.map((r) => r.text).join('')).toBe(`第 ${PAGE_MARK} 页`)
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const names = Object.keys(zip.files).filter((n) => /word\/footer\d+\.xml/.test(n))
    const xml = await zip.file(names[0])!.async('string')
    expect(xml).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>')
  })

  it('footer NUMPAGES field parses to TOTAL_PAGES_MARK and saves back as a field', async () => {
    const { parsed, saveBlocks } = await base()
    const saved = await saveDocx(parsed, saveBlocks, {
      footer: {
        text: '',
        pageNumber: true,
        paras: [{ align: 'center', runs: [{ text: `Seite # von ${TOTAL_PAGES_MARK}` }] }],
      },
    })
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const names = Object.keys(zip.files).filter((n) => /word\/footer\d+\.xml/.test(n))
    const xml = await zip.file(names[0])!.async('string')
    expect(xml).toContain('<w:instrText xml:space="preserve"> NUMPAGES </w:instrText>')
    expect(xml).not.toContain(TOTAL_PAGES_MARK)
    // Re-parse: both fields return to their placeholders (no cached result lingers)
    const reparsed = await parseDocx(saved)
    expect(reparsed.footerParas!.flatMap((p) => p.runs.map((r) => r.text)).join('')).toBe(
      `Seite ${PAGE_MARK} von ${TOTAL_PAGES_MARK}`,
    )
  })

  it('paragraphs inside nested w:sdt keep their alignment (POI Bug60341: OpenXML SDK footer)', async () => {
    const FOOTER =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:sdt><w:sdtContent><w:sdt><w:sdtContent>' +
      '<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="right"/></w:pPr>' +
      '<w:r><w:t xml:space="preserve">Page </w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText>PAGE</w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>2</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '</w:p>' +
      '</w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:ftr>'
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId62" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
      extraParts: [
        {
          path: 'word/footer1.xml',
          xml: FOOTER,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
        },
      ],
      sectPrExtra: '<w:footerReference w:type="default" r:id="rId62"/>',
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.footerParas).toHaveLength(1)
    expect(parsed.footerParas![0].align).toBe('right')
    expect(parsed.footerParas![0].runs.map((r) => r.text).join('')).toBe(`Page ${PAGE_MARK}`)
  })

  it('Header style w:tabs and w:jc reach the paragraphs; direct pPr wins (POI ThreeColHead)', async () => {
    const HEADER =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>' +
      '<w:r><w:t xml:space="preserve">Left\tMid\tRight</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="right"/></w:pPr>' +
      '<w:r><w:t>direct wins</w:t></w:r></w:p>' +
      '</w:hdr>'
    const STYLES =
      '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/>' +
      '<w:pPr><w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" w:pos="9026"/></w:tabs>' +
      '<w:jc w:val="left"/></w:pPr></w:style>'
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraStylesXml: STYLES,
      extraRels:
        '<Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: HEADER,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId61"/>',
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.headerParas![0].tabStops).toEqual([
      { pos: 4513, val: 'center' },
      { pos: 9026, val: 'right' },
    ])
    expect(parsed.headerParas![0].align).toBe('left')
    expect(parsed.headerParas![1].align).toBe('right')
    expect(parsed.headerParas![1].tabStops).toEqual([
      { pos: 4513, val: 'center' },
      { pos: 9026, val: 'right' },
    ])
  })

  it('w:framePr xAlign surfaces as frameXAlign (POI WordWithAttachments page number frame)', async () => {
    const HEADER =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:pPr><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="right" w:y="1"/></w:pPr>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve">PAGE  </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
      '<w:p><w:r><w:t>Страница 1</w:t></w:r></w:p>' +
      '</w:hdr>'
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: HEADER,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId61"/>',
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.headerParas![0].frameXAlign).toBe('right')
    expect(parsed.headerParas![0].runs.map((r) => r.text).join('')).toBe(PAGE_MARK)
    expect(parsed.headerParas![1].frameXAlign).toBeUndefined()
  })

  it('w:ptab absolute tabs surface as \\t runs with their alignments (POI ThreeColHead)', async () => {
    const HEADER =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>' +
      '<w:r><w:t>First header column!</w:t></w:r>' +
      '<w:r><w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/></w:r>' +
      '<w:r><w:t>Mid header</w:t></w:r>' +
      '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>' +
      '<w:r><w:t>Right header!</w:t></w:r>' +
      '</w:p></w:hdr>'
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: HEADER,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId61"/>',
    })
    const parsed = await parseDocx(bytes)
    const para = parsed.headerParas![0]
    expect(para.runs.map((r) => r.text).join('')).toBe(
      'First header column!\tMid header\tRight header!',
    )
    expect(para.ptabAligns).toEqual(['center', 'right'])
  })

  it('mixed w:tab / w:ptab keeps ptab alignments on the right tab index', async () => {
    const HEADER =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:r><w:t>a</w:t></w:r>' +
      '<w:r><w:tab/><w:t>b</w:t></w:r>' +
      '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>' +
      '<w:r><w:t>c</w:t></w:r>' +
      '</w:p></w:hdr>'
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: HEADER,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId61"/>',
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.headerParas![0].ptabAligns).toEqual([undefined, 'right'])
  })

  it('literal # before a PAGE field stays literal; the field parses to PAGE_MARK', async () => {
    // regression: the footer "[Course #] | Page <PAGE field>" used to render the
    // page number inside "[Course #]" because '#' doubled as the field marker
    const FOOTER =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p>' +
      '<w:r><w:t xml:space="preserve">[Course #] | Page </w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText>PAGE</w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>1</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '</w:p></w:ftr>'
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId62" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
      extraParts: [
        {
          path: 'word/footer1.xml',
          xml: FOOTER,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
        },
      ],
      sectPrExtra: '<w:footerReference w:type="default" r:id="rId62"/>',
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.footerHasPageNumber).toBe(true)
    const line = parsed.footerParas!.flatMap((p) => p.runs.map((r) => r.text)).join('')
    expect(line).toBe(`[Course #] | Page ${PAGE_MARK}`)

    // saving the unchanged value keeps the field at its position, not at the literal '#'
    const saveBlocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, saveBlocks, {
      footer: { text: parsed.footerText!, pageNumber: true, paras: parsed.footerParas! },
    })
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const xml = await zip.file('word/footer1.xml')!.async('string')
    expect(xml.match(/<w:instrText[^>]*> PAGE <\/w:instrText>/g)).toHaveLength(1)
    expect(xml).toContain('[Course #] | Page ')
    expect(xml).not.toContain(PAGE_MARK)
    const reparsed = await parseDocx(saved)
    expect(reparsed.footerParas!.flatMap((p) => p.runs.map((r) => r.text)).join('')).toBe(
      `[Course #] | Page ${PAGE_MARK}`,
    )
  })

  it('rich header edit keeps the watermark paragraph', async () => {
    const { parsed, saveBlocks } = await base()
    const withWm = await saveDocx(parsed, saveBlocks, {
      header: { text: '旧页眉' },
      watermark: '草稿',
    })
    const reparsed = await parseDocx(withWm)
    expect(reparsed.watermarkText).toBe('草稿')
    // rich rewrite carries the watermark through (options.watermark undefined)
    const again = await saveDocx(
      reparsed,
      reparsed.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! })),
      { header: { text: '', paras: [{ runs: [{ text: '新页眉', bold: true }] }] } },
    )
    const final = await parseDocx(again)
    expect(final.watermarkText).toBe('草稿')
    expect(final.headerParas![0].runs[0]).toMatchObject({ text: '新页眉', bold: true })
    // the watermark paragraph is not part of the text paras
    expect(final.headerText).toBe('新页眉')
  })
})

describe('surgical header rewrite (real Word headers carry tables/logos)', () => {
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
  const HEADER_TBL =
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:shd w:val="clear" w:fill="1F3864"/></w:tcPr>' +
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ACME 公司</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
    '<w:p><w:r><w:t>Logo 占位</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  const RICH_HEADER =
    XML_DECL +
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    HEADER_TBL +
    '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>公司内部资料</w:t></w:r></w:p>' +
    '</w:hdr>'

  async function richHeaderDocx() {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>正文</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: RICH_HEADER,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId60"/>',
    })
    const parsed = await parseDocx(bytes)
    const saveBlocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    return { parsed, saveBlocks }
  }

  it('watermark-only change keeps the header table and paragraphs byte-identical', async () => {
    const { parsed, saveBlocks } = await richHeaderDocx()
    const saved = await saveDocx(parsed, saveBlocks, { watermark: '机密' })
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const hdr = await zip.file('word/header1.xml')!.async('string')
    // Before the surgical fix: the whole part was rebuilt, losing the table and the
    // right-aligned italic paragraph
    expect(hdr).toContain(HEADER_TBL)
    expect(hdr).toContain(
      '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>公司内部资料</w:t></w:r></w:p>',
    )
    expect(hdr).toContain('v:textpath')
    expect(hdr).toContain('string="机密"')
    // Then remove the watermark: the table remains, the watermark paragraph is gone
    const removed = await saveDocx(await parseDocx(saved), saveBlocks, { watermark: null })
    const hdr2 = await (
      await (await import('jszip')).default.loadAsync(removed)
    )
      .file('word/header1.xml')!
      .async('string')
    expect(hdr2).toContain(HEADER_TBL)
    expect(hdr2).not.toContain('v:textpath')
  })

  it('header text edit replaces paragraphs but keeps the table', async () => {
    const { parsed, saveBlocks } = await richHeaderDocx()
    const saved = await saveDocx(parsed, saveBlocks, { header: { text: '新版页眉' } })
    const hdr = await (
      await (await import('jszip')).default.loadAsync(saved)
    )
      .file('word/header1.xml')!
      .async('string')
    expect(hdr).toContain(HEADER_TBL) // the table (invisible in the user's model) is kept
    expect(hdr).toContain('新版页眉')
    expect(hdr).not.toContain('公司内部资料') // the text-paragraph set was replaced
  })

  it('parses the header table as a cells paragraph (columns, widths, cell gap in text)', async () => {
    const { parsed } = await richHeaderDocx()
    const paras = parsed.headerParas!
    expect(paras).toHaveLength(2)
    const row = paras[0]
    expect(row.cells).toHaveLength(2)
    expect(row.cells![0].paras).toHaveLength(1)
    expect(row.cells![0].paras[0][0]).toMatchObject({ text: 'ACME 公司', bold: true })
    expect(row.cells![1].paras[0][0]).toMatchObject({ text: 'Logo 占位' })
    expect(row.cells!.map((c) => Math.round(c.widthPct!))).toEqual([50, 50])
    expect(row.cells![0].fill).toBe('1F3864')
    expect(row.cells![1].fill).toBeUndefined()
    expect(paras[1].runs[0].text).toBe('公司内部资料')
    // legacy plain text separates cell texts instead of gluing them
    expect(parsed.headerText).toBe('ACME 公司 Logo 占位 公司内部资料')
  })

  it('keeps every cell paragraph as its own line (two-paragraph title cell, empty shaded cell)', async () => {
    const twoParaTbl = HEADER_TBL.replace(
      '<w:p><w:r><w:t>Logo 占位</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>Title line 1</w:t></w:r></w:p><w:p><w:r><w:t>Title line 2</w:t></w:r></w:p>',
    ).replace('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ACME 公司</w:t></w:r></w:p>', '<w:p/>')
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>正文</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml:
            XML_DECL +
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            twoParaTbl +
            '</w:hdr>',
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId60"/>',
    })
    const row = (await parseDocx(bytes)).headerParas![0]
    // the shaded cell keeps its empty paragraph (it must still reserve a line/fill)
    expect(row.cells![0].paras).toEqual([[]])
    expect(row.cells![0].fill).toBe('1F3864')
    expect(row.cells![1].paras.map((rs) => rs.map((r) => r.text).join(''))).toEqual([
      'Title line 1',
      'Title line 2',
    ])
  })

  it('saving paras that include a cells paragraph keeps the table bytes and never duplicates them', async () => {
    const { parsed, saveBlocks } = await richHeaderDocx()
    const saved = await saveDocx(parsed, saveBlocks, {
      header: {
        text: '新页眉',
        paras: [parsed.headerParas![0], { align: 'right', runs: [{ text: '新页眉' }] }],
      },
    })
    const hdr = await (
      await (await import('jszip')).default.loadAsync(saved)
    )
      .file('word/header1.xml')!
      .async('string')
    expect(hdr).toContain(HEADER_TBL)
    expect(hdr.split('ACME 公司').length - 1).toBe(1) // table content only inside w:tbl
    expect(hdr).toContain('新页眉')
  })
})

describe('header edits keep image paragraphs', () => {
  it('the logo paragraph with w:drawing keeps its original bytes; text paragraphs are replaced', async () => {
    const LOGO_P =
      '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/></wp:inline></w:drawing></w:r></w:p>'
    const HDR =
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      LOGO_P +
      '<w:p><w:r><w:t>旧页眉文字</w:t></w:r></w:p>' +
      '</w:hdr>'
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>正文</w:t></w:r></w:p>',
      extraRels:
        '<Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: HDR,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId61"/>',
    })
    const parsed = await parseDocx(bytes)
    const saveBlocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, saveBlocks, { header: { text: '新页眉文字' } })
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const hdr = await zip.file('word/header1.xml')!.async('string')
    expect(hdr).toContain(LOGO_P) // the logo paragraph keeps its original bytes
    expect(hdr).toContain('新页眉文字')
    expect(hdr).not.toContain('旧页眉文字')
  })
})

describe('header paragraph borders', () => {
  const HDR_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

  /** a header whose side borders are reset with w:val="nil", as a style-level reset writes them */
  async function withNilBorderHeader() {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      sectPrExtra: '<w:headerReference w:type="default" r:id="rIdHdr"/>',
      extraRels:
        '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml:
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${HDR_NS}>` +
            '<w:p><w:pPr><w:pBdr><w:top w:val="nil"/><w:left w:val="nil"/><w:right w:val="nil"/>' +
            '<w:bottom w:val="single" w:sz="18" w:color="4472C4"/></w:pBdr></w:pPr>' +
            '<w:r><w:t>Quarterly report</w:t></w:r></w:p></w:hdr>',
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
    })
    return parseDocx(bytes)
  }

  it('does not read a nil side as a border', async () => {
    const parsed = await withNilBorderHeader()
    expect(parsed.headerParas?.[0]?.borders).toBe('b')
    expect(parsed.headerParas?.[0]?.borderLines).toEqual({ b: { color: '4472C4', szPt: 2.25 } })
  })

  it('does not stamp rules the header never had when it is saved back', async () => {
    const parsed = await withNilBorderHeader()
    const blocks: SaveBlock[] = [{ kind: 'original', docxIndex: parsed.blocks[0].docxIndex! }]
    const saved = await saveDocx(parsed, blocks, {
      header: { text: '', paras: parsed.headerParas ?? [] },
    })
    const header = await (await JSZip.loadAsync(saved)).file('word/header1.xml')!.async('text')
    const pBdr = /<w:pBdr>[\s\S]*?<\/w:pBdr>/.exec(header)?.[0] ?? ''
    expect(pBdr).toContain('<w:bottom ')
    for (const side of ['<w:top ', '<w:left ', '<w:right ']) {
      expect(pBdr).not.toContain(side)
    }
  })
})
