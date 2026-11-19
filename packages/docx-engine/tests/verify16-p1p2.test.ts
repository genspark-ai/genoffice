/**
 * verify16-p1p2.test.ts
 * P1+P2 end-to-end verification (pure engine layer, no Electron)
 * Covers: chart workbook / character styles / image posOffset / tab stops / drop caps /
 * SDT / full revision semantics
 * Run: cd packages/docx-engine && npx vitest run tests/verify16-p1p2.test.ts
 */

import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  buildChartWorkbookXlsxBase64,
  patchChartWorkbookXlsxBase64,
  parseDocx,
  saveDocx,
  stripPPrChange,
} from '../src/index'
import type { GeneratedBlock } from '../src/types'
import { buildDocx } from './helpers/build-docx'

// ═══════════════════════════════════════════════════════════════════════════
// T1 – chart embedded workbook (embeddings/*.xlsx + sheetData consistent)
// ═══════════════════════════════════════════════════════════════════════════

const CHART_CATEGORIES = ['A', 'B', 'C']
const CHART_SERIES = [{ name: 'S1', values: [10, 20, 30] as (number | null)[] }]
const CHART_SPEC = {
  kind: 'bar' as const,
  title: '验证图表',
  categories: CHART_CATEGORIES,
  series: CHART_SERIES,
}

describe('T1: chart embedded workbook', () => {
  it('buildChartWorkbookXlsxBase64 produces a valid xlsx whose sheetData contains the chart data', async () => {
    // buildChartWorkbookXlsxBase64 takes (categories, series) separately
    const b64 = await buildChartWorkbookXlsxBase64(CHART_CATEGORIES, CHART_SERIES)
    expect(b64.length).toBeGreaterThan(100)
    const buf = Buffer.from(b64, 'base64')
    const zip = await JSZip.loadAsync(buf)
    // xl/workbook.xml exists
    expect(zip.file('xl/workbook.xml')).toBeTruthy()
    // xl/worksheets/sheet1.xml contains sheetData
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<sheetData')
    // contains categories and values (via sharedStrings)
    const ss = zip.file('xl/sharedStrings.xml')
    if (ss) {
      const ssXml = await ss.async('string')
      expect(ssXml).toContain('>A<')
    }
    // numeric values appear inline in sheet
    expect(sheet).toContain('>10<')
    expect(sheet).toContain('>20<')
  })

  it('saveDocx kind:chart creates the chart part + ContentType + rels', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>前</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'chart', chart: CHART_SPEC },
    ])
    const zip = await JSZip.loadAsync(saved)
    // chart part exists
    expect(zip.file('word/charts/chart1.xml')).toBeTruthy()
    const chartXml = await zip.file('word/charts/chart1.xml')!.async('string')
    expect(chartXml).toContain('验证图表')
    // ContentTypes includes the chart
    const ct = await zip.file('[Content_Types].xml')!.async('string')
    expect(ct).toContain('/word/charts/chart1.xml')
    // document rels include the chart
    const docRels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(docRels).toContain('charts/chart1.xml')
  })

  it('saveDocx kind:chart generates the embeddings xlsx + c:externalData', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>前</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'chart', chart: CHART_SPEC },
    ])
    const zip = await JSZip.loadAsync(saved)
    const chartXml = await zip.file('word/charts/chart1.xml')!.async('string')
    // c:externalData is written (workbook rel)
    expect(chartXml).toContain('externalData')
    // workbook xlsx exists in the zip (path: word/charts/embeddings/workbook1.xlsx)
    const wb = zip.file(/word\/charts\/embeddings\/.*\.xlsx/)
    expect(wb.length).toBeGreaterThan(0)
    // chart rels include the workbook
    const chartRels = await zip.file('word/charts/_rels/chart1.xml.rels')!.async('string')
    expect(chartRels).toContain('embeddings/')
    // the workbook xlsx contains sheetData
    const wbBytes = await wb[0].async('uint8array')
    const wbZip = await JSZip.loadAsync(wbBytes)
    const sheet = await wbZip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<sheetData')
    expect(sheet).toContain('>10<')
  })

  it('patchChartWorkbookXlsxBase64 updates the values', async () => {
    const b64 = await buildChartWorkbookXlsxBase64(CHART_CATEGORIES, CHART_SERIES)
    const newSeries = [{ name: 'S1', values: [99, 88, 77] as (number | null)[] }]
    const patched = await patchChartWorkbookXlsxBase64(b64, CHART_CATEGORIES, newSeries)
    expect(patched).not.toBeNull()
    const zip = await JSZip.loadAsync(Buffer.from(patched!, 'base64'))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('>99<')
    expect(sheet).toContain('>88<')
    expect(sheet).not.toContain('>10<') // old value gone
  })
})

// ═══════════════════════════════════════════════════════════════════════════
// T2 – character style rStyle preservation
// ═══════════════════════════════════════════════════════════════════════════

const CHAR_STYLE_P =
  '<w:p>' +
  '<w:r>' +
  '<w:rPr><w:rStyle w:val="Emphasis"/></w:rPr>' +
  '<w:t>emphasized text</w:t>' +
  '</w:r>' +
  '</w:p>'

describe('T2: character style rStyle preservation', () => {
  it('parsing yields runs[0] with styleId=Emphasis', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: CHAR_STYLE_P,
        extraStylesXml:
          '<w:style w:type="character" w:styleId="Emphasis">' +
          '<w:name w:val="Emphasis"/>' +
          '<w:rPr><w:i/></w:rPr>' +
          '</w:style>',
      }),
    )
    const run = doc.blocks[0].runs?.[0]
    expect(run?.styleId).toBe('Emphasis')
  })

  it('rStyle survives saveDocx and reparse', async () => {
    const source = await buildDocx({
      bodyXml: CHAR_STYLE_P,
      extraStylesXml:
        '<w:style w:type="character" w:styleId="Emphasis">' +
        '<w:name w:val="Emphasis"/>' +
        '<w:rPr><w:i/></w:rPr>' +
        '</w:style>',
    })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }])
    const zip = await JSZip.loadAsync(saved)
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml).toContain('w:val="Emphasis"')
  })

  it('character-type styles from the style sheet are parsed into the styles Map', async () => {
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      extraStylesXml:
        '<w:style w:type="character" w:styleId="StrongEmph">' +
        '<w:name w:val="Strong Emphasis"/>' +
        '<w:rPr><w:b/><w:i/></w:rPr>' +
        '</w:style>',
    })
    const parsed = await parseDocx(source)
    expect(parsed.styles).toBeDefined()
    const style = parsed.styles?.get('StrongEmph')
    expect(style?.type).toBe('character')
    expect(style?.display?.bold).toBe(true)
  })
})

// ═══════════════════════════════════════════════════════════════════════════
// T3 – image posOffset write
// ═══════════════════════════════════════════════════════════════════════════

/** Floating image XML with posOffset */
const FLOAT_IMG_P =
  '<w:p><w:r><w:rPr/><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" ' +
  'simplePos="0" relativeHeight="251658240" behindDoc="0" ' +
  'locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>457200</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>914400</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1905000" cy="1905000"/>' +
  '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  '<wp:wrapSquare wrapText="bothSides"/>' +
  '<wp:docPr id="1" name="Image1"/>' +
  '<wp:cNvGraphicFramePr/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:nvPicPr><pic:cNvPr id="1" name="Image1"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill>' +
  '<a:blip r:embed="rId10"/>' +
  '<a:stretch><a:fillRect/></a:stretch>' +
  '</pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
  '</pic:pic>' +
  '</a:graphicData>' +
  '</a:graphic>' +
  '</wp:anchor>' +
  '</w:drawing></w:r></w:p>'

describe('T3: image posOffset read/write', () => {
  it('parsing a floating image reads imageOffsetXEmu / imageOffsetYEmu', async () => {
    const doc = await parseDocx(
      // withImage:true adds rId10 rel + word/media/image1.png so extractImage succeeds
      await buildDocx({ bodyXml: FLOAT_IMG_P, withImage: true }),
    )
    const block = doc.blocks.find((b) => b.imageOffsetXEmu !== undefined)
    expect(block).toBeDefined()
    expect(block!.imageOffsetXEmu).toBe(457200)
    expect(block!.imageOffsetYEmu).toBe(914400)
  })

  it('wp:posOffset values are preserved after saveDocx', async () => {
    const source = await buildDocx({ bodyXml: FLOAT_IMG_P, withImage: true })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }])
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('<wp:posOffset>457200</wp:posOffset>')
    expect(xml).toContain('<wp:posOffset>914400</wp:posOffset>')
  })
})

// ═══════════════════════════════════════════════════════════════════════════
// T4 – tab stops w:tabs save
// ═══════════════════════════════════════════════════════════════════════════

const TABS_P =
  '<w:p>' +
  '<w:pPr>' +
  '<w:tabs>' +
  '<w:tab w:val="left" w:pos="1440"/>' +
  '<w:tab w:val="center" w:pos="2880"/>' +
  '</w:tabs>' +
  '</w:pPr>' +
  '<w:r><w:t>a\tb\tc</w:t></w:r>' +
  '</w:p>'

describe('T4: tab stops w:tabs', () => {
  it('parses the tabStops array', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TABS_P }))
    const block = doc.blocks[0]
    expect(block.format?.tabStops).toBeDefined()
    expect(block.format!.tabStops!.length).toBe(2)
    expect(block.format!.tabStops![0]).toEqual({ pos: 1440, val: 'left' })
    expect(block.format!.tabStops![1]).toEqual({ pos: 2880, val: 'center' })
  })

  it('w:tabs is written back correctly after saveDocx', async () => {
    const source = await buildDocx({ bodyXml: TABS_P })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }])
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:tabs>')
    expect(xml).toContain('w:val="left"')
    expect(xml).toContain('w:pos="1440"')
    expect(xml).toContain('w:val="center"')
    expect(xml).toContain('w:pos="2880"')
  })
})

// ═══════════════════════════════════════════════════════════════════════════
// T5 – drop cap w:framePr dropCap
// ═══════════════════════════════════════════════════════════════════════════

const DROPCAP_P =
  '<w:p>' +
  '<w:pPr>' +
  '<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>' +
  '</w:pPr>' +
  '<w:r><w:t>T</w:t></w:r>' +
  '</w:p>'

describe('T5: drop cap w:framePr dropCap', () => {
  it('parses dropCap attributes', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: DROPCAP_P }))
    const block = doc.blocks[0]
    expect(block.format?.dropCap).toBeDefined()
    expect(block.format!.dropCap!.type).toBe('drop')
    expect(block.format!.dropCap!.lines).toBe(3)
  })

  it('w:framePr dropCap is written back after saveDocx', async () => {
    const source = await buildDocx({ bodyXml: DROPCAP_P })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }])
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('w:dropCap="drop"')
    expect(xml).toContain('w:lines="3"')
  })

  it('margin dropCap parses as type=margin', async () => {
    const p =
      '<w:p><w:pPr>' +
      '<w:framePr w:dropCap="margin" w:lines="2" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>' +
      '</w:pPr><w:r><w:t>M</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: p }))
    expect(doc.blocks[0].format?.dropCap?.type).toBe('margin')
  })
})

// ═══════════════════════════════════════════════════════════════════════════
// T6 – SDT content-control shell preservation
// ═══════════════════════════════════════════════════════════════════════════

const SDT_P =
  '<w:sdt>' +
  '<w:sdtPr>' +
  '<w:alias w:val="AuthorField"/>' +
  '<w:tag w:val="author_tag"/>' +
  '<w:text/>' +
  '</w:sdtPr>' +
  '<w:sdtEndPr/>' +
  '<w:sdtContent>' +
  '<w:p><w:r><w:t>Initial Value</w:t></w:r></w:p>' +
  '</w:sdtContent>' +
  '</w:sdt>'

describe('T6: SDT content controls', () => {
  it('parses the sdt block and exposes sdtShell', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: SDT_P }))
    const block = doc.blocks[0]
    expect(block.sdtShell).toBeDefined()
    expect(block.sdtShell!.alias).toBe('AuthorField')
    expect(block.sdtShell!.tag).toBe('author_tag')
    expect(block.sdtShell!.controlType).toBe('text')
    expect(block.runs?.[0].text).toBe('Initial Value')
  })

  it('passthrough: the w:sdt shell is kept when saveDocx does not touch the sdt content', async () => {
    const source = await buildDocx({ bodyXml: SDT_P })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }])
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:sdt>')
    expect(xml).toContain('w:val="AuthorField"')
    expect(xml).toContain('Initial Value')
  })

  it('save after editing: sdt shell kept, sdtContent updated', async () => {
    const source = await buildDocx({ bodyXml: SDT_P })
    const parsed = await parseDocx(source)
    const block = parsed.blocks[0]
    expect(block.sdtShell).toBeDefined()
    // Build a generated SaveBlock that edits the text but preserves the sdtShell
    const genBlock: GeneratedBlock = {
      type: 'paragraph',
      runs: [{ text: 'Updated Value' }],
      sdtShell: block.sdtShell,
    }
    const saved = await saveDocx(parsed, [{ kind: 'generated', block: genBlock }])
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('word/document.xml')!.async('string')
    // SDT shell preserved
    expect(xml).toContain('<w:sdt>')
    expect(xml).toContain('w:val="AuthorField"')
    // content updated
    expect(xml).toContain('Updated Value')
    expect(xml).not.toContain('Initial Value')
  })
})

// ═══════════════════════════════════════════════════════════════════════════
// T7 – full revision semantics (pPrChange / move revisions)
// ═══════════════════════════════════════════════════════════════════════════

const PPRCHANGE_P =
  '<w:p><w:pPr>' +
  '<w:jc w:val="center"/>' +
  '<w:pPrChange w:id="51" w:author="Alice" w:date="2026-07-01T10:00:00Z">' +
  '<w:pPr><w:jc w:val="left"/></w:pPr>' +
  '</w:pPrChange>' +
  '</w:pPr>' +
  '<w:r><w:t>段落修订文字</w:t></w:r>' +
  '</w:p>'

const MOVEFROM_P =
  '<w:p><w:moveFrom w:id="41" w:author="Alice" w:name="move1">' +
  '<w:r><w:delText>moved away text</w:delText></w:r>' +
  '</w:moveFrom></w:p>'

const MOVETO_P =
  '<w:p><w:moveTo w:id="42" w:author="Bob" w:name="move2">' +
  '<w:r><w:t>moved here text</w:t></w:r>' +
  '</w:moveTo></w:p>'

describe('T7: full tracked-revision semantics', () => {
  it('pPrChange: parses pPrChangeInfo.author and keeps rawPPr', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PPRCHANGE_P }))
    const block = doc.blocks[0]
    expect(block.pPrChangeInfo).toBeDefined()
    expect(block.pPrChangeInfo!.author).toBe('Alice')
    expect(block.pPrChangeInfo!.id).toBe('51')
    expect(block.rawPPr).toContain('<w:pPrChange')
  })

  it('pPrChange: still present after saveDocx passthrough', async () => {
    const source = await buildDocx({ bodyXml: PPRCHANGE_P })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }])
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:pPrChange')
    expect(xml).toContain('w:author="Alice"')
  })

  it('stripPPrChange removes the pPrChange tag', () => {
    const rawPPr =
      '<w:pPr><w:jc w:val="center"/><w:pPrChange w:id="51" w:author="Alice"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>'
    const stripped = stripPPrChange(rawPPr)
    expect(stripped).not.toContain('<w:pPrChange')
    expect(stripped).toContain('<w:jc w:val="center"')
  })

  it('moveFrom: parses as del marks + moveRevision=from', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: MOVEFROM_P }))
    const block = doc.blocks[0]
    expect(block.moveRevision).toBe('from')
    const delRun = block.runs?.find((r) => r.del)
    expect(delRun).toBeDefined()
    expect(delRun?.text).toContain('moved away text')
  })

  it('moveTo: parses as ins marks + moveRevision=to', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: MOVETO_P }))
    const block = doc.blocks[0]
    expect(block.moveRevision).toBe('to')
    const insRun = block.runs?.find((r) => r.ins)
    expect(insRun).toBeDefined()
    expect(insRun?.text).toContain('moved here text')
  })
})
