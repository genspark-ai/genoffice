import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  bibliographyLine,
  citationText,
  parseDocx,
  readThemeColors,
  readThemeFonts,
  saveDocx,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const BODY = '<w:p><w:r><w:t>正文</w:t></w:r></w:p>'

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

describe('watermark', () => {
  it('writes a VML text watermark into the header and reads it back', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const out = await saveDocx(doc, originalOrder(doc), { watermark: '机密' })
    const zip = await JSZip.loadAsync(out)
    const header = await zip.file('word/header1.xml')!.async('string')
    expect(header).toContain('<v:textpath')
    expect(header).toContain('string="机密"')

    const reparsed = await parseDocx(out)
    expect(reparsed.watermarkText).toBe('机密')
  })

  it('a header text rewrite carries the existing watermark through', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const withWm = await parseDocx(await saveDocx(doc, originalOrder(doc), { watermark: '草稿' }))
    const out = await parseDocx(
      await saveDocx(withWm, originalOrder(withWm), { header: { text: '公司内部' } }),
    )
    expect(out.headerText).toBe('公司内部')
    expect(out.watermarkText).toBe('草稿')
  })

  it('watermark: null removes it but keeps the header text', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const withBoth = await parseDocx(
      await saveDocx(doc, originalOrder(doc), { watermark: '草稿', header: { text: '页眉' } }),
    )
    const cleared = await parseDocx(
      await saveDocx(withBoth, originalOrder(withBoth), { watermark: null }),
    )
    expect(cleared.watermarkText).toBeNull()
    expect(cleared.headerText).toBe('页眉')
  })
})

describe('theme fonts / colors', () => {
  it('creates word/theme/theme1.xml when the document has none', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    expect(doc.themeFonts).toBeNull()
    const out = await saveDocx(doc, originalOrder(doc), {
      themeFonts: { major: '微软雅黑', minor: '宋体', eastAsia: '宋体' },
      themeColors: { accent1: '1F4E79' },
    })
    const zip = await JSZip.loadAsync(out)
    const theme = await zip.file('word/theme/theme1.xml')!.async('string')
    // buildThemeXml writes the eastAsia face into both font groups
    expect(readThemeFonts(theme)).toEqual({
      major: '微软雅黑',
      minor: '宋体',
      eastAsia: '宋体',
      majorEastAsia: '宋体',
    })
    expect(readThemeColors(theme)?.accent1).toBe('1F4E79')
    const ct = await zip.file('[Content_Types].xml')!.async('string')
    expect(ct).toContain('/word/theme/theme1.xml')
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('theme/theme1.xml')

    const reparsed = await parseDocx(out)
    expect(reparsed.themeFonts).toEqual({
      major: '微软雅黑',
      minor: '宋体',
      eastAsia: '宋体',
      majorEastAsia: '宋体',
    })
  })

  it('patches an existing theme part in place', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const first = await parseDocx(
      await saveDocx(doc, originalOrder(doc), {
        themeFonts: { major: 'Calibri Light', minor: 'Calibri' },
      }),
    )
    const out = await saveDocx(first, originalOrder(first), {
      themeFonts: { major: 'Georgia', minor: 'Georgia' },
      themeColors: { accent1: '70AD47', name: '绿色' },
    })
    const zip = await JSZip.loadAsync(out)
    const theme = await zip.file('word/theme/theme1.xml')!.async('string')
    expect(readThemeFonts(theme)?.major).toBe('Georgia')
    expect(readThemeColors(theme)).toMatchObject({ accent1: '70AD47', name: '绿色' })
  })
})

describe('bibliography sources', () => {
  const source = {
    tag: 'Wang2024',
    type: 'Book',
    author: '王明',
    title: '深度学习实践',
    year: '2024',
    publisher: '清华出版社',
  }

  it('creates the b:Sources customXml part with rels + content type', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    expect(doc.sources).toEqual([])
    const out = await saveDocx(doc, originalOrder(doc), { sources: [source] })
    const zip = await JSZip.loadAsync(out)
    const item = await zip.file('customXml/item1.xml')!.async('string')
    expect(item).toContain('b:Sources')
    expect(item).toContain('深度学习实践')
    expect(zip.file('customXml/itemProps1.xml')).toBeTruthy()
    expect(zip.file('customXml/_rels/item1.xml.rels')).toBeTruthy()
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('../customXml/item1.xml')

    const reparsed = await parseDocx(out)
    expect(reparsed.sources).toEqual([source])
  })

  it('rewrites the existing sources part on later saves', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const first = await parseDocx(await saveDocx(doc, originalOrder(doc), { sources: [source] }))
    const second = await parseDocx(
      await saveDocx(first, originalOrder(first), {
        sources: [
          ...first.sources,
          {
            tag: 'Li2023',
            type: 'JournalArticle',
            author: '李强',
            title: 'LLM 综述',
            year: '2023',
            publisher: '计算机学报',
          },
        ],
      }),
    )
    expect(second.sources).toHaveLength(2)
    expect(second.sources[1].publisher).toBe('计算机学报')
  })

  it('citation and bibliography text formatting', () => {
    expect(citationText(source)).toBe('(王明, 2024)')
    expect(bibliographyLine(source)).toBe('王明. (2024). 深度学习实践. 清华出版社.')
  })

  it('adding a source keeps unmodeled fields and citation style of existing entries (surgical rebuild)', async () => {
    // Original document: GB citation style + a journal entry with Volume/Pages/DOI/
    // multiple authors (all fields the engine does not model)
    const RICH_SOURCES =
      '<b:Sources SelectedStyle="\\GB7714.xsl" StyleName="GB7714" Version="2005"' +
      ' xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"' +
      ' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">' +
      '<b:Source><b:Tag>Zhao2022</b:Tag><b:SourceType>JournalArticle</b:SourceType>' +
      '<b:Author><b:Author><b:NameList>' +
      '<b:Person><b:Last>赵</b:Last><b:First>一</b:First></b:Person>' +
      '<b:Person><b:Last>钱</b:Last><b:First>二</b:First></b:Person>' +
      '</b:NameList></b:Author></b:Author>' +
      '<b:Title>大模型对齐</b:Title><b:Year>2022</b:Year>' +
      '<b:JournalName>软件学报</b:JournalName>' +
      '<b:Volume>33</b:Volume><b:Issue>4</b:Issue><b:Pages>1-20</b:Pages>' +
      '<b:StandardNumber>doi:10.1000/xyz</b:StandardNumber>' +
      '</b:Source></b:Sources>'
    const bytes = await buildDocx({
      bodyXml: BODY,
      extraParts: [
        { path: 'customXml/item1.xml', xml: RICH_SOURCES, contentType: 'application/xml' },
      ],
    })
    const doc = await parseDocx(bytes)
    expect(doc.sources).toHaveLength(1)
    const out = await saveDocx(doc, originalOrder(doc), {
      sources: [...doc.sources, source],
    })
    const item = await (await JSZip.loadAsync(out)).file('customXml/item1.xml')!.async('string')
    // Before the surgical fix: the root element was rewritten to APA6 and Volume/Pages/
    // DOI/the second author were all lost
    expect(item).toContain('SelectedStyle="\\GB7714.xsl"')
    expect(item).toContain('<b:Volume>33</b:Volume>')
    expect(item).toContain('<b:Pages>1-20</b:Pages>')
    expect(item).toContain('doi:10.1000/xyz')
    expect(item).toContain('<b:Last>钱</b:Last>')
    expect(item).toContain('深度学习实践') // the new entry is there too
  })
})

describe('w:themeColor resolution', () => {
  const THEME_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">' +
    '<a:themeElements><a:clrScheme name="Office">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="1A1A1A"/></a:dk1>' +
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
    '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
    '</a:clrScheme><a:fontScheme name="Office">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>' +
    '</a:themeElements></a:theme>'

  const themedDocx = (bodyXml: string) =>
    buildDocx({
      bodyXml,
      extraRels:
        '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
      extraParts: [
        {
          path: 'word/theme/theme1.xml',
          xml: THEME_XML,
          contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
        },
      ],
    })

  it('parseDocx returns the palette, including sysClr dk1/lt1 and hlink', async () => {
    const doc = await parseDocx(await themedDocx(BODY))
    expect(doc.themeColors).toMatchObject({
      dk1: '1A1A1A',
      lt1: 'FFFFFF',
      accent1: '4472C4',
      hlink: '0563C1',
    })
  })

  it('run color resolves w:themeColor against the palette, beating a stale w:val', async () => {
    const doc = await parseDocx(
      await themedDocx(
        '<w:p><w:r><w:rPr><w:color w:val="DEAD00" w:themeColor="accent1"/></w:rPr><w:t>主题色</w:t></w:r></w:p>',
      ),
    )
    expect(doc.blocks[0].runs?.[0].color).toBe('4472C4')
  })

  it('applies themeTint / themeShade (sRGB approximation)', async () => {
    const doc = await parseDocx(
      await themedDocx(
        '<w:p><w:r><w:rPr><w:color w:val="8EAADB" w:themeColor="accent1" w:themeTint="99"/></w:rPr><w:t>浅</w:t></w:r>' +
          '<w:r><w:rPr><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/></w:rPr><w:t>深</w:t></w:r></w:p>',
      ),
    )
    // tint 0x99: c*0.6 + 255*0.4 -> 8FAADC (Word writes 8EAADB; rounding differs by at most 1)
    expect(doc.blocks[0].runs?.[0].color).toBe('8FAADC')
    // shade 0xBF: c*0.75 -> 335593 (Word writes 2F5496 — linear-space difference; visually close enough)
    expect(doc.blocks[0].runs?.[1].color).toBe('335593')
  })

  it('w:val alone still wins when the doc has no theme part', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:rPr><w:color w:val="FF0000" w:themeColor="accent1"/></w:rPr><w:t>x</w:t></w:r></w:p>',
      }),
    )
    expect(doc.themeColors).toBeNull()
    expect(doc.blocks[0].runs?.[0].color).toBe('FF0000')
  })

  it('an untouched themed document still saves byte-identical', async () => {
    const bytes = await themedDocx(
      '<w:p><w:r><w:rPr><w:color w:val="DEAD00" w:themeColor="accent1"/></w:rPr><w:t>主题色</w:t></w:r></w:p>',
    )
    const doc = await parseDocx(bytes)
    const out = await saveDocx(doc, originalOrder(doc), {})
    expect(out).toEqual(bytes)
  })
})
