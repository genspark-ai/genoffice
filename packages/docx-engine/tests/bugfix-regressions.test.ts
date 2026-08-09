import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  applySectionSettings,
  buildChartWorkbookXlsxBase64,
  generateTableModelXml,
  latexToOmml,
  nextNoteId,
  ommlToLatex,
  parseDocx,
  patchChartWorkbookXlsxBase64,
  patchTableCellTexts,
  readSections,
  saveDocx,
  sectionSettingsFromXml,
  type SaveBlock,
} from '../src/index'
import { patchParagraphTexts } from '../src/text-patch'
import { escapeXmlText } from '../src/xml-utils'
import { buildDocx } from './helpers/build-docx'

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

describe('saveDocx isUnchanged short-circuit', () => {
  it('a sectionStartType-only edit is not silently dropped', async () => {
    const bytes = await buildDocx({ bodyXml: P('a') + P('b') })
    const parsed = await parseDocx(bytes)
    const visible: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, visible, { sectionStartType: 'continuous' })
    expect(saved).not.toBe(bytes)
    const sections = readSections(await parseDocx(saved))
    expect(sections[0].startType).toBe('continuous')
  })
})

describe('applySectionSettings w:cols', () => {
  const SECT_PR =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
    '<w:cols w:num="3" w:space="720">' +
    '<w:col w:w="3000" w:space="720"/><w:col w:w="3600" w:space="720"/><w:col w:w="2000"/>' +
    '</w:cols></w:sectPr>'

  it('keeps explicit column widths when the count is unchanged', () => {
    const settings = sectionSettingsFromXml(SECT_PR)
    expect(settings.columns).toBe(3)
    const out = applySectionSettings(SECT_PR, settings)
    expect(out.match(/<w:cols[\s>]/g)?.length).toBe(1)
    expect(out).toContain('<w:col w:w="3000"')
  })

  it('replaces the open-form element without duplicating it when the count changes', () => {
    const settings = sectionSettingsFromXml(SECT_PR)
    const out = applySectionSettings(SECT_PR, { ...settings, columns: 2 })
    expect(out.match(/<w:cols[\s>]/g)?.length).toBe(1)
    expect(out).toContain('w:num="2"')
    expect(out).not.toContain('<w:col w:w=')
  })
})

describe('patchParagraphTexts with nested textbox paragraphs', () => {
  const ENTRY =
    '<w:comment w:id="1">' +
    '<w:p><w:r><w:t xml:space="preserve">outer </w:t></w:r>' +
    '<w:r><w:drawing><wps:txbx><w:txbxContent>' +
    '<w:p><w:r><w:t>inside box</w:t></w:r></w:p>' +
    '</w:txbxContent></wps:txbx></w:drawing></w:r>' +
    '<w:r><w:t>tail</w:t></w:r></w:p>' +
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>second</w:t></w:r></w:p>' +
    '</w:comment>'

  it('slices the outer paragraph past the embedded w:p and patches surgically', () => {
    const out = patchParagraphTexts(ENTRY, 'outer inside boxtail\nsecond edited')
    expect(out).not.toBeNull()
    // the textbox structure of the untouched paragraph survives byte-for-byte
    expect(out).toContain(
      '<w:txbxContent><w:p><w:r><w:t>inside box</w:t></w:r></w:p></w:txbxContent>',
    )
    expect(out).toContain('second edited')
    expect(out).toContain('<w:rPr><w:b/></w:rPr>')
  })
})

describe('escapeXmlText', () => {
  it('strips control characters that are illegal in XML 1.0 but keeps tab/newline', () => {
    expect(escapeXmlText('a\u0000b\u000Bc\td\ne')).toBe('abc\td\ne')
    expect(escapeXmlText('<a & b>')).toBe('&lt;a &amp; b&gt;')
  })
})

describe('nextNoteId', () => {
  it('starts at 1 for a document with no existing notes', () => {
    expect(nextNoteId([])).toBe('1')
  })
})

describe('ommlToLatex matrix environments', () => {
  it('recovers the named pmatrix environment instead of generic delimiters', () => {
    const omml = `<m:oMath>${latexToOmml('\\begin{pmatrix} 1 & 0 \\\\ 0 & 1 \\end{pmatrix}')}</m:oMath>`
    expect(ommlToLatex(omml)).toContain('\\begin{pmatrix}')
  })
})

describe('patchChartWorkbookXlsxBase64', () => {
  it('preserves unrelated workbook parts and updates Sheet1 in place', async () => {
    const base = await buildChartWorkbookXlsxBase64(['a', 'b'], [{ name: 'S1', values: [1, 2] }])
    const zip = await JSZip.loadAsync(Buffer.from(base, 'base64'))
    zip.file('xl/styles.xml', '<styleSheet/>')
    zip.file('xl/worksheets/sheet2.xml', '<worksheet><sheetData/></worksheet>')
    const withExtras = Buffer.from(await zip.generateAsync({ type: 'uint8array' })).toString(
      'base64',
    )

    const patched = await patchChartWorkbookXlsxBase64(
      withExtras,
      ['x', 'y'],
      [{ name: 'S1', values: [9, 8] }],
    )
    expect(patched).not.toBeNull()
    const outZip = await JSZip.loadAsync(Buffer.from(patched!, 'base64'))
    expect(outZip.file('xl/styles.xml')).toBeTruthy()
    expect(outZip.file('xl/worksheets/sheet2.xml')).toBeTruthy()
    const sheet = await outZip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<v>9</v>')
    expect(sheet).toContain('>x</t>')
  })
})

describe('body-level invisible markers (stray w:bookmarkEnd)', () => {
  const BODY = P('before') + '<w:bookmarkEnd w:id="3"/>' + P('after')

  it('parses as an invisible passthrough block that keeps its position', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const marker = doc.blocks[1]
    expect(marker.type).toBe('passthrough')
    expect(marker.label).toBe('w:bookmarkEnd')
    expect(marker.invisibleMarker).toBe(true)
    // not `hidden`: hidden blocks are moved to the body tail on save
    expect(marker.hidden).toBeUndefined()
  })

  it('saves byte-identically in its original position', async () => {
    const bytes = await buildDocx({ bodyXml: BODY })
    const doc = await parseDocx(bytes)
    const visible: SaveBlock[] = doc.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const saved = await saveDocx(doc, visible)
    expect(saved).toBe(bytes)
  })
})

describe('paragraph with stray text plus content-carrying textboxes', () => {
  const WPS_BOX = (text: string) =>
    '<w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:extent cx="2857500" cy="457200"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'

  it('keeps the textbox content instead of dropping it for the stray run', async () => {
    const bodyXml =
      '<w:p><w:r><w:t>F,</w:t></w:r>' +
      WPS_BOX('REPUBLIQUE DEMOCRATIQUE DU CONGO') +
      WPS_BOX('MINISTERE DES MINES') +
      '</w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Text box')
    expect(block.textboxes).toHaveLength(2)
    expect(block.textboxes![0].paras[0].runs[0].text).toBe('REPUBLIQUE DEMOCRATIQUE DU CONGO')
    // the stray text joins the preview instead of vanishing
    expect(block.previewText).toContain('F,')
    expect(block.previewText).toContain('MINISTERE DES MINES')
  })

  it('a decorative shape without txbxContent still leaves the paragraph editable', async () => {
    const bodyXml =
      '<w:p><w:r><w:t>body text</w:t></w:r>' +
      '<w:r><w:drawing><wp:anchor behindDoc="1"><wp:extent cx="914400" cy="9525"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp>' +
      '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].runs?.map((r) => r.text).join('')).toBe('body text')
  })
})

describe('inline pictures in table cells', () => {
  const IMG_RUN =
    '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  const IMG_TABLE =
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr><w:tc><w:p>${IMG_RUN}</w:p></w:tc>` +
    '<w:tc><w:p><w:r><w:t>notes</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

  it('parses the cell picture into an image run', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: IMG_TABLE, withImage: true }))
    const run = doc.blocks[0].table!.rows[0][0].richParas![0].runs[0]
    expect(run.text).toBe('')
    expect(run.image?.dataUrl).toMatch(/^data:image\/png;base64,/)
    expect(run.image?.widthPx).toBe(96) // 914400 EMU / 9525
    expect(run.image?.heightPx).toBe(48)
    expect(run.image?.xml).toContain('<a:blip r:embed="rId10"/>')
  })

  it('keeps the drawing when the cell text is patched', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: IMG_TABLE, withImage: true }))
    const out = patchTableCellTexts(doc.blocks[0].originalXml!, [[['new text'], null]])
    expect(out).toContain('<w:t xml:space="preserve">new text</w:t>')
    expect(out).toContain('r:embed="rId10"')
    expect(out).toContain('<w:t>notes</w:t>')
  })

  it('re-emits the drawing verbatim on whole-table regeneration', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: IMG_TABLE, withImage: true }))
    const xml = generateTableModelXml(doc.blocks[0].table!)
    expect(xml).toContain('r:embed="rId10"')
    expect(xml).toContain('<w:drawing>')
  })

  it('untouched table with a cell picture saves byte-identically', async () => {
    const bytes = await buildDocx({ bodyXml: IMG_TABLE, withImage: true })
    const doc = await parseDocx(bytes)
    const visible: SaveBlock[] = doc.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, visible)).toBe(bytes)
  })
})

describe('cell-level color needs run consensus (B5)', () => {
  const CELL = (runs: string) =>
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr><w:tc><w:p>${runs}</w:p></w:tc></w:tr></w:tbl>`
  const RED = '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>red run</w:t></w:r>'
  const PLAIN = '<w:r><w:t>plain run</w:t></w:r>'

  it('mixed run colors leave cell.color unset, run colors intact', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: CELL(RED + PLAIN) }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.color).toBeUndefined()
    expect(cell.richParas![0].runs[0].color).toBe('FF0000')
    expect(cell.richParas![0].runs[1].color).toBeUndefined()
  })

  it('a uniformly colored cell still gets cell.color', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: CELL(RED + RED) }))
    expect(doc.blocks[0].table!.rows[0][0].color).toBe('FF0000')
  })
})

describe('table style whole-table rPr and firstCol conditionals (B5b)', () => {
  const STYLE =
    '<w:style w:type="table" w:styleId="LightShading">' +
    '<w:name w:val="Light Shading"/>' +
    '<w:rPr><w:color w:val="365F91"/></w:rPr>' +
    '<w:tblStylePr w:type="firstCol"><w:rPr><w:b/></w:rPr></w:tblStylePr>' +
    '</w:style>'
  const TC = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
  const TBL =
    '<w:tbl><w:tblPr><w:tblStyle w:val="LightShading"/>' +
    '<w:tblLook w:val="04A0" w:firstRow="1" w:firstColumn="1" w:noHBand="0" w:noVBand="1"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr>${TC('A1')}${TC('B1')}</w:tr><w:tr>${TC('A2')}${TC('B2')}</w:tr></w:tbl>`

  it('applies the style rPr color to every cell and firstCol bold to the first column', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TBL, extraStylesXml: STYLE }))
    const rows = doc.blocks[0].table!.rows
    for (const row of rows) for (const cell of row) expect(cell.color).toBe('365F91')
    expect(rows[0][0].bold).toBe(true)
    expect(rows[1][0].bold).toBe(true)
    expect(rows[1][1].bold).toBeUndefined()
  })
})

describe('duplicated w:tcW takes the last occurrence (B7)', () => {
  const TC = (widths: number[], t: string) =>
    `<w:tc><w:tcPr>${widths
      .map((w) => `<w:tcW w:w="${w}" w:type="dxa"/>`)
      .join('')}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
  const TBL =
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr>${TC([4933, 1872], 'A')}${TC([4933, 7776], 'B')}</w:tr></w:tbl>`

  it('column widths come from the second tcW, not the stale first one', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TBL }))
    expect(doc.blocks[0].table!.colWidthsTwips).toEqual([1872, 7776])
  })
})

describe('trHeight hRule round-trip', () => {
  const TBL =
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr><w:trPr><w:trHeight w:val="400" w:hRule="exact"/></w:trPr>' +
    '<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:trPr><w:trHeight w:val="500"/></w:trPr>' +
    '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

  it('parses the rule and writes it back instead of forcing atLeast', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TBL }))
    const table = doc.blocks[0].table!
    expect(table.rowHeightRules).toEqual(['exact', 'atLeast'])
    const xml = generateTableModelXml(table)
    expect(xml).toContain('<w:trHeight w:val="400" w:hRule="exact"/>')
    expect(xml).toContain('<w:trHeight w:val="500" w:hRule="atLeast"/>')
  })
})

describe('picture with a broken rel degrades to an empty frame block', () => {
  const PIC =
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>' +
    '<wp:inline><wp:extent cx="1905000" cy="952500"/>' +
    '<wp:docPr id="5" name="Image 5" descr="photo of the site"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId99"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

  it('exposes brokenImage + alt text + extent instead of a Drawing object chip', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PIC }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.brokenImage).toBe(true)
    expect(block.previewText).toBe('photo of the site')
    expect(block.imageWidthPx).toBe(200)
    expect(block.imageHeightPx).toBe(100)
    expect(block.imageAlign).toBe('center')
  })

  it('falls back to the docPr name when descr is absent', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: PIC.replace(' descr="photo of the site"', '') }),
    )
    expect(doc.blocks[0].brokenImage).toBe(true)
    expect(doc.blocks[0].previewText).toBe('Image 5')
  })

  it('still saves byte-identically', async () => {
    const bytes = await buildDocx({ bodyXml: PIC })
    const doc = await parseDocx(bytes)
    const visible: SaveBlock[] = doc.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, visible)).toBe(bytes)
  })
})

describe('explicitly invisible empty shapes are suppressed', () => {
  const SHAPE = (spPrChildren: string, txbx = '<w:p/>') =>
    '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:extent cx="2934335" cy="1032510"/><wp:wrapNone/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPrChildren}</wps:spPr>` +
    `<wps:txbx><w:txbxContent>${txbx}</w:txbxContent></wps:txbx>` +
    '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

  it('noFill + noFill outline + empty text renders as nothing', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: SHAPE('<a:noFill/><a:ln><a:noFill/></a:ln>') }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.invisibleMarker).toBe(true)
  })

  it('a filled shape is not suppressed', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: SHAPE(
          '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln>',
        ),
      }),
    )
    expect(doc.blocks[0].invisibleMarker).toBeUndefined()
  })

  it('a shape with visible effects is not suppressed', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: SHAPE(
          '<a:noFill/><a:ln><a:noFill/></a:ln>' +
            '<a:effectLst><a:outerShdw dist="28398" dir="3806097"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>',
        ),
      }),
    )
    expect(doc.blocks[0].invisibleMarker).toBeUndefined()
  })

  it('a shape with text is not suppressed', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: SHAPE(
          '<a:noFill/><a:ln><a:noFill/></a:ln>',
          '<w:p><w:r><w:t>hi</w:t></w:r></w:p>',
        ),
      }),
    )
    expect(doc.blocks[0].invisibleMarker).toBeUndefined()
  })
})

describe('decorative rule carries its stroke display', () => {
  const LINE =
    '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:extent cx="5230495" cy="20955"/><wp:wrapNone/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5230495" cy="20955"/></a:xfrm>' +
    '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>' +
    '<a:ln w="38100"><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></a:ln>' +
    '</wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

  it('exposes color / thickness / width from a:ln and wp:extent', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: LINE }))
    const block = doc.blocks[0]
    expect(block.decorative).toBe(true)
    expect(block.ruleColorHex).toBe('0070C0')
    expect(block.ruleThicknessPx).toBe(4)
    expect(block.ruleWidthPx).toBe(549)
  })

  it('a rule without a:ln keeps the fields absent (default 1px line)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: LINE.replace(
          '<a:ln w="38100"><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></a:ln>',
          '',
        ),
      }),
    )
    expect(doc.blocks[0].decorative).toBe(true)
    expect(doc.blocks[0].ruleColorHex).toBeUndefined()
    expect(doc.blocks[0].ruleThicknessPx).toBeUndefined()
  })
})

describe('gradFill textbox approximates the average of all stops as a solid fill', () => {
  const BOX = (fillXml: string) =>
    '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:extent cx="4897120" cy="520700"/><wp:wrapNone/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fillXml}<a:ln><a:noFill/></a:ln></wps:spPr>` +
    '<wps:txbx><w:txbxContent><w:p><w:r><w:t>JUIN 2026</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
    '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

  it('mixes srgbClr stops equally (white + color -> the midpoint tint)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BOX(
          '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="4472C4"/></a:gs>' +
            '<a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst></a:gradFill>',
        ),
      }),
    )
    // (4472C4 + FFFFFF) / 2
    expect(doc.blocks[0].textboxes?.[0].fill).toBe('A2B9E2')
  })

  it('resolves schemeClr stops against the theme with lumMod/lumOff, then mixes', async () => {
    const THEME_XML =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">' +
      '<a:themeElements><a:clrScheme name="T">' +
      '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="1F3864"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>' +
      '<a:accent1><a:srgbClr val="ED7D31"/></a:accent1>' +
      '</a:clrScheme></a:themeElements></a:theme>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BOX(
          '<a:gradFill><a:gsLst>' +
            '<a:gs pos="0"><a:schemeClr val="accent1"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr></a:gs>' +
            '<a:gs pos="100000"><a:schemeClr val="lt1"/></a:gs></a:gsLst></a:gradFill>',
        ),
        extraParts: [
          {
            path: 'word/theme/theme1.xml',
            xml: THEME_XML,
            contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
          },
        ],
      }),
    )
    // (ED7D31 at lumMod 40% + lumOff 60%, i.e. ~F8CBAD) mixed with lt1 white
    expect(doc.blocks[0].textboxes?.[0].fill).toBe('FBE5D6')
  })

  it('an explicit solidFill still wins over the approximation path', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: BOX('<a:solidFill><a:srgbClr val="00B050"/></a:solidFill>') }),
    )
    expect(doc.blocks[0].textboxes?.[0].fill).toBe('00B050')
  })
})
