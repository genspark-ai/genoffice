import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, TOTAL_PAGES_MARK, type SaveBlock } from '../src/index'
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
    expect(reparsed.footerParas![0].runs.map((r) => r.text).join('')).toBe('第 # 页')
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
    // Re-parse: the NUMPAGES field returns to the placeholder (no cached result lingers)
    const reparsed = await parseDocx(saved)
    expect(reparsed.footerParas!.flatMap((p) => p.runs.map((r) => r.text)).join('')).toBe(
      `Seite # von ${TOTAL_PAGES_MARK}`,
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
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
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
