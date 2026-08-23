import { describe, expect, it } from 'vitest'
import { buildAnchoredTextboxParagraphXml } from '../src/generate'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

/**
 * P20 region container: a WPS text box with real w:p content, anchored either
 * to its holder paragraph (card regions that ride the flow, wrap topAndBottom)
 * or to the page (newsletter regions at absolute coordinates).
 */

const CARD_PARAS = [
  {
    runs: [{ text: 'FINAL TAKEAWAY', bold: true, color: 'FFFFFF' }],
    format: { spaceAfter: 60, lineRule: 'exact' as const, lineRawTwips: 320 },
  },
  {
    runs: [{ text: '结论正文第二行', color: 'FFFFFF' }],
    format: {
      spaceAfter: 0,
      lineRule: 'exact' as const,
      lineRawTwips: 280,
      align: 'left' as const,
    },
  },
]

const cardXml = () =>
  buildAnchoredTextboxParagraphXml({
    anchor: 'paragraph',
    xEmu: 91440,
    yEmu: 12700,
    widthEmu: 5486400,
    heightEmu: 1371600,
    fillHex: '1A1A2E',
    insetsEmu: { l: 182880, t: 91440, r: 182880, b: 91440 },
    zOrder: 7,
    id: 42,
    paragraphs: CARD_PARAS,
    holderLineTwips: 20,
  })

describe('buildAnchoredTextboxParagraphXml (P20 region container)', () => {
  it('paragraph anchor floats with the flow and wraps topAndBottom', () => {
    const xml = cardXml()
    expect(xml).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>91440</wp:posOffset></wp:positionH>',
    )
    expect(xml).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>12700</wp:posOffset></wp:positionV>',
    )
    expect(xml).toContain('<wp:wrapTopAndBottom/>')
    expect(xml).not.toContain('<wp:wrapSquare')
    // CT_Anchor child order: wrap element sits between effectExtent and docPr
    const wrapAt = xml.indexOf('<wp:wrapTopAndBottom/>')
    expect(wrapAt).toBeGreaterThan(xml.indexOf('<wp:effectExtent'))
    expect(wrapAt).toBeLessThan(xml.indexOf('<wp:docPr'))
    expect(xml).toContain('<wp:extent cx="5486400" cy="1371600"/>')
  })

  it('page anchor pins both axes to the page and wraps none', () => {
    const xml = buildAnchoredTextboxParagraphXml({
      anchor: 'page',
      xEmu: 457200,
      yEmu: 914400,
      widthEmu: 2743200,
      heightEmu: 3657600,
      paragraphs: [{ runs: [{ text: 'column A' }] }],
    })
    expect(xml).toContain(
      '<wp:positionH relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionH>',
    )
    expect(xml).toContain(
      '<wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>',
    )
    expect(xml).toContain('<wp:wrapNone/>')
  })

  it('writes fill, z-order and behindDoc discipline', () => {
    const xml = cardXml()
    expect(xml).toContain('<a:solidFill><a:srgbClr val="1A1A2E"/></a:solidFill>')
    expect(xml).toContain(`relativeHeight="${251658240 + 7}"`)
    expect(xml).toContain('behindDoc="0"')
    // no border unless asked for
    expect(xml).toContain('<a:ln><a:noFill/></a:ln>')
  })

  it('omitted fill stays transparent', () => {
    const xml = buildAnchoredTextboxParagraphXml({
      anchor: 'page',
      xEmu: 0,
      yEmu: 0,
      widthEmu: 914400,
      heightEmu: 914400,
      paragraphs: [{ runs: [{ text: 'x' }] }],
    })
    expect(xml).toContain('<a:noFill/>')
    expect(xml).not.toContain('<a:solidFill>')
  })

  it('writes measured insets on wps:bodyPr', () => {
    const xml = cardXml()
    const bodyPr = /<wps:bodyPr[^>]*>/.exec(xml)?.[0] ?? ''
    expect(bodyPr).toContain('lIns="182880"')
    expect(bodyPr).toContain('tIns="91440"')
    expect(bodyPr).toContain('rIns="182880"')
    expect(bodyPr).toContain('bIns="91440"')
    // fixed box: content must not auto-grow the region
    expect(xml).toContain('<a:noAutofit/>')
  })

  it('rounded corners map to roundRect with the measured radius', () => {
    const xml = buildAnchoredTextboxParagraphXml({
      anchor: 'paragraph',
      xEmu: 0,
      yEmu: 0,
      widthEmu: 4000000,
      heightEmu: 1000000,
      cornerRadiusEmu: 100000,
      paragraphs: [{ runs: [{ text: 'rounded' }] }],
    })
    // adj = radius / min(w,h) * 100000 = 100000/1000000*100000 = 10000
    expect(xml).toContain(
      '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>',
    )
  })

  it('content is real editable w:p paragraphs with line structure', () => {
    const xml = cardXml()
    const txbx =
      /<wps:txbx><w:txbxContent>([\s\S]*?)<\/w:txbxContent><\/wps:txbx>/.exec(xml)?.[1] ?? ''
    expect(txbx.match(/<w:p>/g)?.length).toBe(2)
    expect(txbx).toContain('FINAL TAKEAWAY')
    expect(txbx).toContain('结论正文第二行')
    expect(txbx).toContain('<w:spacing w:after="60" w:line="320" w:lineRule="exact"/>')
    expect(txbx).toContain('<w:color w:val="FFFFFF"/>')
  })

  it('holder paragraph keeps an exact tight line so the anchor row adds no height', () => {
    const xml = cardXml()
    expect(
      xml.startsWith(
        '<w:p><w:pPr><w:spacing w:after="0" w:line="20" w:lineRule="exact"/></w:pPr><w:r>',
      ),
    ).toBe(true)
  })

  it('holder paragraph carries spacing-before and page-break-before when asked', () => {
    const xml = buildAnchoredTextboxParagraphXml({
      anchor: 'paragraph',
      xEmu: 0,
      yEmu: 0,
      widthEmu: 914400,
      heightEmu: 914400,
      paragraphs: [{ runs: [{ text: 'x' }] }],
      holderLineTwips: 20,
      holderSpacingBeforeTwips: 240,
      holderPageBreakBefore: true,
    })
    expect(
      xml.startsWith(
        '<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="240" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr><w:r>',
      ),
    ).toBe(true)
  })

  it('declares Requires prefixes at AlternateContent scope and pt-sized VML fallback', () => {
    const xml = cardXml()
    const acTag = /<mc:AlternateContent[^>]*>/.exec(xml)?.[0] ?? ''
    for (const m of xml.matchAll(/Requires="([^"]+)"/g)) {
      for (const prefix of m[1]!.split(/\s+/)) {
        expect(acTag).toContain(`xmlns:${prefix}=`)
      }
    }
    const style = /<v:(?:round)?rect[^>]*style="([^"]*)"/.exec(xml)?.[1] ?? ''
    expect(style).toContain(`width:${5486400 / 12700}pt`)
    expect(style).toContain(`height:${1371600 / 12700}pt`)
    expect(xml).toContain('fillcolor="#1A1A2E"')
  })

  it('round-trips through parseDocx as an editable textbox', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: cardXml() }))
    const box = parsed.blocks.flatMap((b) => b.textboxes ?? [])[0]
    expect(box).toBeDefined()
    expect(box!.fill?.toUpperCase()).toBe('1A1A2E')
    expect(box!.paras.map((p) => p.runs.map((r) => r.text).join(''))).toEqual([
      'FINAL TAKEAWAY',
      '结论正文第二行',
    ])
    expect(box!.readOnly).not.toBe(true)
    expect(box!.widthPx).toBe(Math.round(5486400 / 9525))
    expect(box!.insetLeftPx).toBeCloseTo(182880 / 9525, 1)
  })
})
