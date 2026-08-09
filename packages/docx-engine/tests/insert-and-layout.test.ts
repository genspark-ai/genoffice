import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  generateParagraphXml,
  generateTableXml,
  parseDocx,
  readPageColor,
  readSectionSettings,
  saveDocx,
  type GenerateContext,
  type SaveBlock,
} from '../src/index'
import { TINY_PNG_BASE64, buildKitchenSinkDocx } from './helpers/build-docx'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map(),
  allocateHyperlinkRel: () => 'rId999',
}

async function docXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

function originals(indexes: number[]): SaveBlock[] {
  return indexes.map((docxIndex) => ({ kind: 'original', docxIndex }))
}

describe('new paragraph format props', () => {
  it('round-trips spaceBefore/spaceAfter/indentRight/pageBreakBefore', async () => {
    const para = generateParagraphXml(
      {
        type: 'paragraph',
        format: { spaceBefore: 240, spaceAfter: 120, indentRight: 360, pageBreakBefore: true },
        runs: [{ text: 'x' }],
      },
      GEN_CTX,
    )
    expect(para).toContain('<w:pageBreakBefore/>')
    expect(para).toContain('<w:spacing w:before="240" w:after="120"/>')
    expect(para).toContain('<w:ind w:right="360"/>')
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden).map((b) => b.docxIndex!)
    const saved = await saveDocx(doc, [...originals(visible), { kind: 'xml', xml: para }])
    const reparsed = await parseDocx(saved)
    const last = reparsed.blocks.filter((b) => !b.hidden).at(-1)!
    expect(last.format).toEqual({
      spaceBefore: 240,
      spaceAfter: 120,
      indentRight: 360,
      pageBreakBefore: true,
    })
  })
})

describe('inserting new tables (xml save blocks)', () => {
  it('generates a bordered fixed table and reparses it as a table block', async () => {
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden).map((b) => b.docxIndex!)
    const saved = await saveDocx(doc, [
      { kind: 'xml', xml: generateTableXml(3, 4) },
      ...originals(visible),
    ])
    const xml = await docXmlOf(saved)
    expect(xml).toContain('<w:tblBorders>')
    const reparsed = await parseDocx(saved)
    const first = reparsed.blocks.filter((b) => !b.hidden)[0]
    expect(first.type).toBe('table')
    expect(first.label).toBe('Table 3×4')
  })
})

describe('inserting new images', () => {
  it('adds media, relationship and content-type default', async () => {
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden).map((b) => b.docxIndex!)
    const saved = await saveDocx(doc, [
      ...originals(visible),
      {
        kind: 'image',
        image: { base64: TINY_PNG_BASE64, mime: 'image/png', widthPx: 100, heightPx: 60 },
      },
    ])
    const zip = await JSZip.loadAsync(saved)
    const media = Object.keys(zip.files).filter((n) => n.startsWith('word/media/aidocs'))
    expect(media).toHaveLength(1)
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain(`Target="${media[0].replace('word/', '')}"`)
    const xml = await docXmlOf(saved)
    expect(xml).toContain(`<wp:extent cx="${100 * 9525}" cy="${60 * 9525}"/>`)
    const reparsed = await parseDocx(saved)
    const last = reparsed.blocks.filter((b) => !b.hidden).at(-1)!
    expect(last.type).toBe('image')
    expect(last.imageDataUrl).toMatch(/^data:image\/png;base64,/)
  })

  it('records rotation bounding-box overflow in effectExtent for new images', async () => {
    const bytes = await buildKitchenSinkDocx()
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden).map((b) => b.docxIndex!)
    // 200x100px turned 90°: overflows the unrotated footprint by
    // (bboxH - cy)/2 = (200-100)px/2 * 9525 EMU on top/bottom
    const saved = await saveDocx(doc, [
      ...originals(visible),
      {
        kind: 'image',
        image: {
          base64: TINY_PNG_BASE64,
          mime: 'image/png',
          widthPx: 200,
          heightPx: 100,
          rotDeg: 90,
          flipH: true,
        },
      },
    ])
    const xml = await docXmlOf(saved)
    expect(xml).toContain(`<wp:effectExtent l="0" t="${50 * 9525}" r="0" b="${50 * 9525}"/>`)
    expect(xml).toContain(`rot="${90 * 60000}"`)
    expect(xml).toContain('flipH="1"')
    const reparsed = await parseDocx(saved)
    const last = reparsed.blocks.filter((b) => !b.hidden).at(-1)!
    expect(last.imageRotDeg).toBe(90)
    expect(last.imageFlipH).toBe(true)
  })
})

describe('section settings (page setup)', () => {
  it('parses sectPr vAlign and in-paragraph page breaks / hyperlink tooltip', async () => {
    const { sectionSettingsFromXml } = await import('../src/section')
    expect(sectionSettingsFromXml('<w:sectPr><w:vAlign w:val="center"/></w:sectPr>').vAlign).toBe(
      'center',
    )
    expect(sectionSettingsFromXml('<w:sectPr/>').vAlign).toBeUndefined()

    const { buildDocx } = await import('./helpers/build-docx')
    const bodyXml =
      '<w:p><w:r><w:t>上半</w:t><w:br w:type="page"/><w:t>下半</w:t></w:r></w:p>' +
      '<w:p><w:hyperlink r:id="rIdLink" w:tooltip="提示语"><w:r><w:t>链</w:t></w:r></w:hyperlink></w:p>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml,
        extraRels:
          '<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://a.cn" TargetMode="External"/>',
      }),
    )
    expect(doc.blocks[0].runs?.[0].text).toBe('上半\f下半')
    const linkRun = doc.blocks[1].runs?.find((r) => r.link)
    expect(linkRun?.link?.tooltip).toBe('提示语')
  })

  it('reads defaults from the kitchen sink sectPr (A4 portrait)', async () => {
    const doc = await parseDocx(await buildKitchenSinkDocx())
    const section = readSectionSettings(doc)
    expect(section).toEqual({
      pageWidth: 11906,
      pageHeight: 16838,
      orientation: 'portrait',
      marginTop: 1440,
      marginRight: 1440,
      marginBottom: 1440,
      marginLeft: 1440,
      headerDist: 708,
      footerDist: 708,
      pageBorder: false,
      columns: 1,
      colSpace: 720,
    })
  })

  it('rewrites pgSz and pgMar while keeping header/footer attrs', async () => {
    const doc = await parseDocx(await buildKitchenSinkDocx())
    const visible = doc.blocks.filter((b) => !b.hidden).map((b) => b.docxIndex!)
    const saved = await saveDocx(doc, originals(visible), {
      section: {
        pageWidth: 16838,
        pageHeight: 11906,
        orientation: 'landscape',
        marginTop: 720,
        marginRight: 720,
        marginBottom: 720,
        marginLeft: 720,
        pageBorder: false,
        columns: 1,
      },
    })
    const xml = await docXmlOf(saved)
    expect(xml).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>')
    expect(xml).toMatch(
      /<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708"/,
    )
    const reparsed = await parseDocx(saved)
    expect(readSectionSettings(reparsed).orientation).toBe('landscape')
    expect(readSectionSettings(reparsed).marginLeft).toBe(720)
  })
})

describe('page color', () => {
  it('sets, reads back and removes w:background', async () => {
    const doc = await parseDocx(await buildKitchenSinkDocx())
    expect(readPageColor(doc)).toBeNull()
    const visible = doc.blocks.filter((b) => !b.hidden).map((b) => b.docxIndex!)

    const colored = await saveDocx(doc, originals(visible), { pageColor: 'FFF2CC' })
    const coloredXml = await docXmlOf(colored)
    expect(coloredXml).toMatch(/<w:document[^>]*><w:background w:color="FFF2CC"\/>/)
    const reparsedColored = await parseDocx(colored)
    expect(readPageColor(reparsedColored)).toBe('FFF2CC')

    const cleared = await saveDocx(
      reparsedColored,
      reparsedColored.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
      { pageColor: null },
    )
    expect(readPageColor(await parseDocx(cleared))).toBeNull()
  })
})
