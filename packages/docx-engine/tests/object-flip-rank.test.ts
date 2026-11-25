/**
 * What the reader takes from a drawing for Arrange: mirror flips on every
 * kind of box, the stacking rank of a floating picture among text, a
 * saveable source for a shape anchored in a table cell, and a group placed by
 * alignment split from the corner the page drew.
 */
import { describe, expect, it } from 'vitest'
import { splitGroupParagraph } from '../src/group-drawings'
import { parseDocx } from '../src/parse'
import { patchRunPictureXml } from '../src/run-picture'
import { buildDocx } from './helpers/build-docx'

const EMU = 9525

const anchorOpen = (rank: number, behind = false) =>
  `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${rank}" ` +
  `behindDoc="${behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  `<wp:extent cx="${100 * EMU}" cy="${50 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`

const shape = (prst: string, xfrmAttrs: string, id = 5, text = '') =>
  `<w:drawing>${anchorOpen(251659264)}<wp:docPr id="${id}" name="Shape ${id}"/><wp:cNvGraphicFramePr/>` +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm${xfrmAttrs}><a:off x="0" y="0"/><a:ext cx="${100 * EMU}" cy="${50 * EMU}"/></a:xfrm>` +
  `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>` +
  '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>' +
  (text
    ? `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>`
    : '') +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>'

const picture = (rank: number, xfrmAttrs = '', id = 9, behind = false) =>
  `<w:drawing>${anchorOpen(rank, behind)}<wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  '<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  `<pic:spPr><a:xfrm${xfrmAttrs}><a:off x="0" y="0"/><a:ext cx="${100 * EMU}" cy="${50 * EMU}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing>'

const run = (inner: string) => `<w:r>${inner}</w:r>`
const text = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`

async function boxesOf(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  return parsed.blocks.flatMap((block) => block.textboxes ?? [])
}

describe('mirror flips of drawn boxes', () => {
  it('reads flipH / flipV of a preset shape, with or without text', async () => {
    const [bare] = await boxesOf(`<w:p>${run(shape('rtTriangle', ' flipH="1"'))}</w:p>`)
    expect(bare).toMatchObject({ prst: 'rtTriangle', flipH: true })
    expect(bare.flipV).toBeUndefined()
    const [worded] = await boxesOf(
      `<w:p>${run(shape('rightArrow', ' rot="5400000" flipV="1"', 6, 'Go'))}</w:p>`,
    )
    expect(worded).toMatchObject({ prst: 'rightArrow', flipV: true, rotDeg: 90 })
    expect(worded.paras[0].runs[0].text).toBe('Go')
  })

  it('reads the flips of a bent connector, which a straight line already had', async () => {
    const [bent] = await boxesOf(`<w:p>${run(shape('bentConnector3', ' flipH="1"'))}</w:p>`)
    expect(bent).toMatchObject({ prst: 'lineBent', flipH: true })
    const [plain] = await boxesOf(`<w:p>${run(shape('bentConnector3', ''))}</w:p>`)
    expect(plain.flipH).toBeUndefined()
  })

  it('reads the flips of a picture drawn as a frame beside a shape', async () => {
    const boxes = await boxesOf(
      `<w:p>${run(shape('ellipse', ''))}${run(picture(251659264, ' flipH="1" flipV="1"'))}</w:p>`,
    )
    const frame = boxes.find((box) => box.picture)
    expect(frame).toMatchObject({ flipH: true, flipV: true })
    expect(boxes.find((box) => !box.picture)?.flipH).toBeUndefined()
  })
})

describe('stacking rank of a floating picture among text', () => {
  const imagesOf = async (bodyXml: string) => {
    const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    return parsed.blocks.flatMap((b) => (b.runs ?? []).flatMap((r) => (r.image ? [r.image] : [])))
  }

  it('is the relativeHeight above Word’s base; the base itself is no rank', async () => {
    const images = await imagesOf(
      `<w:p>${text('A ')}${run(picture(251658240))}${run(picture(251660288, '', 10))}` +
        `${run(picture(251659264, '', 11, true))}${text(' Z')}</w:p>`,
    )
    expect(images.map((image) => [image.wrap, image.zOrder])).toEqual([
      ['front', undefined],
      ['front', 2048],
      ['behind', 1024],
    ])
    expect(images.some((image) => image.zOrderNormalized)).toBe(false)
  })

  it('wild values are compressed together with the other anchored drawings, in order', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p>${text('A ')}${run(picture(3))}${run(picture(1, '', 10))}${text(' Z')}</w:p>` +
          `<w:p>${run(shape('ellipse', '').replace('relativeHeight="251659264"', 'relativeHeight="2"'))}</w:p>`,
        withImage: true,
      }),
    )
    const images = parsed.blocks.flatMap((b) =>
      (b.runs ?? []).flatMap((r) => (r.image ? [r.image] : [])),
    )
    const [box] = parsed.blocks.flatMap((b) => b.textboxes ?? [])
    // the shape's block and its box each hold the shape's rank
    expect(images.map((image) => image.zOrder ?? 0)).toEqual([3, 0])
    expect(box.z).toBeGreaterThan(0)
    expect(box.z).toBeLessThan(3)
    expect(images.every((image) => image.zOrderNormalized)).toBe(true)
  })

  it('a rank is written as relativeHeight and nothing else of the anchor changes', () => {
    const xml = picture(251659264)
    const geometry = { wrap: 'front' as const, offsetXEmu: 0, offsetYEmu: 0 }
    expect(patchRunPictureXml(xml, geometry)).toBe(xml)
    expect(patchRunPictureXml(xml, { ...geometry, zOrder: 7 })).toBe(
      xml.replace('relativeHeight="251659264"', 'relativeHeight="251658247"'),
    )
  })
})

describe('a shape anchored in a table cell', () => {
  const table = (cell: string) =>
    '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
    `${cell}</w:tc></w:tr></w:tbl><w:p/>`

  it('keeps a paragraph of its drawings alone to be saved from', async () => {
    const drawing = shape('ellipse', '')
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: table(`<w:p>${text('Cell ')}${run(drawing)}</w:p>`) }),
    )
    const cell = parsed.blocks[0].table!.rows[0][0]
    expect(cell.anchoredBoxes).toHaveLength(1)
    expect(cell.anchoredBoxXml).toEqual([`<w:p><w:r>${drawing}</w:r></w:p>`])
    expect(cell.paras).toEqual(['Cell '])
  })

  it('has none where a box is not an anchored drawing of the paragraph', async () => {
    const vml =
      '<w:r><w:pict><v:rect style="position:absolute;margin-left:0;margin-top:0;width:60pt;height:30pt">' +
      '<v:textbox><w:txbxContent><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r>'
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: table(`<w:p>${text('Cell')}${vml}</w:p>`) }),
    )
    const cell = parsed.blocks[0].table!.rows[0][0]
    for (const source of cell.anchoredBoxXml ?? []) expect(source).toBeNull()
  })
})

describe('a group placed by alignment', () => {
  const member = (id: number, x: number, y: number) =>
    `<wps:wsp><wps:cNvPr id="${id}" name="Child ${id}"/><wps:cNvSpPr/><wps:spPr><a:xfrm>` +
    `<a:off x="${x * EMU}" y="${y * EMU}"/><a:ext cx="${40 * EMU}" cy="${30 * EMU}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>'
  const group =
    '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="5" ' +
    'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>' +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${20 * EMU}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${200 * EMU}" cy="${100 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
    '<wp:docPr id="50" name="Group 50"/><wp:cNvGraphicFramePr/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${200 * EMU}" cy="${100 * EMU}"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="${200 * EMU}" cy="${100 * EMU}"/></a:xfrm></wpg:grpSpPr>` +
    `${member(1, 0, 0)}${member(2, 150, 60)}</wpg:wgp>` +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

  it('is not split without the corner the page drew', () => {
    expect(splitGroupParagraph(group, 0, 100)).toBeNull()
  })

  it('splits from that corner: the aligned axis becomes an offset from the column, the other keeps its own', () => {
    const split = splitGroupParagraph(group, 0, 100, { xEmu: 300 * EMU, yEmu: 999 * EMU })!
    expect(split.fromCorner).toEqual({ h: true, v: false })
    expect(split.members.map((m) => [m.xEmu / EMU, m.yEmu / EMU])).toEqual([
      [300, 20],
      [450, 80],
    ])
    expect(split.members.map((m) => m.box)).toEqual([
      { x: 0, y: 0, cx: 40 * EMU, cy: 30 * EMU },
      { x: 150 * EMU, y: 60 * EMU, cx: 40 * EMU, cy: 30 * EMU },
    ])
    const all = split.paragraphs.join('')
    expect(all).not.toContain('<wp:align>')
    expect(all).toContain(
      `<wp:positionH relativeFrom="column"><wp:posOffset>${450 * EMU}</wp:posOffset></wp:positionH>`,
    )
    expect(all).toContain(
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${80 * EMU}</wp:posOffset></wp:positionV>`,
    )
  })

  it('an axis placed by a percentage stays unsplit', () => {
    const pct = group.replace(
      /<wp:positionV[^>]*>.*?<\/wp:positionV>/,
      '<wp:positionV relativeFrom="page"><wp14:pctPosVOffset>10000</wp14:pctPosVOffset></wp:positionV>',
    )
    expect(splitGroupParagraph(pct, 0, 100, { xEmu: 0, yEmu: 0 })).toBeNull()
  })
})
