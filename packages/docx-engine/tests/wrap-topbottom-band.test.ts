import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'

function textWsp(opts: {
  cx: number
  cy: number
  text: string
  offY?: number
  autoFit?: boolean
}): string {
  return (
    `<wps:wsp ${WPS_NS}><wps:cNvSpPr txBox="1"/><wps:spPr>` +
    `<a:xfrm><a:off x="0" y="${opts.offY ?? 0}"/><a:ext cx="${opts.cx}" cy="${opts.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
    `</wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>${opts.text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr>${opts.autoFit ? '<a:spAutoFit/>' : ''}</wps:bodyPr></wps:wsp>`
  )
}

function drawing(opts: {
  offXEmu: number
  offYEmu: number
  cx: number
  cy: number
  text: string
  wrap: string
  relV?: string
  inner?: string
}): string {
  return (
    `<w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${opts.offXEmu}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="${opts.relV ?? 'paragraph'}"><wp:posOffset>${opts.offYEmu}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${opts.cx}" cy="${opts.cy}"/>${opts.wrap}` +
    `<wp:docPr id="1" name="Box"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    (opts.inner ?? textWsp(opts)) +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing>`
  )
}

describe('wrapTopAndBottom band reservation', () => {
  it('marks both boxes of a two-anchor paragraph floating with their band bottoms', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 0,
        offYEmu: 0,
        cx: 1905000,
        cy: 257175, // 27 px
        text: 'card',
        wrap: '<wp:wrapTopAndBottom/>',
      }) +
      drawing({
        offXEmu: 2857500,
        offYEmu: 914400, // 96 px
        cx: 1905000,
        cy: 952500, // 100 px
        text: 'chip',
        wrap: '<wp:wrapTopAndBottom/>',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const boxes = doc.blocks[0].textboxes
    expect(boxes?.length).toBe(2)
    const [card, chip] = boxes!
    expect(card.floating).toBe(true)
    expect(card.bandTopPx).toBe(0)
    expect(card.bandBottomPx).toBe(27)
    expect(chip.floating).toBe(true)
    expect(chip.offsetXEmu).toBe(2857500)
    expect(chip.bandTopPx).toBe(96)
    expect(chip.bandBottomPx).toBe(196)
  })

  it('floats a lone topAndBottom box and reserves offset + height', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 91440,
        offYEmu: 190500, // 20 px
        cx: 1905000,
        cy: 476250, // 50 px
        text: 'solo',
        wrap: '<wp:wrapTopAndBottom/>',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBe(true)
    expect(box.bandTopPx).toBe(20)
    expect(box.bandBottomPx).toBe(70)
  })

  it('skips the extent fallback for a group child without its own height', async () => {
    const group =
      `<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">` +
      `<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="3810000" cy="1905000"/></a:xfrm></wpg:grpSpPr>` +
      textWsp({ cx: 1905000, cy: 476250, text: 'fixed' }) +
      textWsp({ cx: 1905000, cy: 476250, text: 'grows', offY: 476250, autoFit: true }) +
      `</wpg:wgp>`
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 0,
        offYEmu: 190500, // 20 px
        cx: 3810000,
        cy: 1905000, // 200 px: whole-group extent, not any child's height
        text: '',
        wrap: '<wp:wrapTopAndBottom/>',
        inner: group,
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const boxes = doc.blocks[0].textboxes
    expect(boxes?.length).toBe(2)
    const [fixed, grows] = boxes!
    expect(fixed.bandTopPx).toBe(20)
    expect(fixed.bandBottomPx).toBe(70)
    expect(grows.floating).toBe(true)
    expect(grows.bandTopPx).toBeUndefined()
    expect(grows.bandBottomPx).toBeUndefined()
  })

  it('clamps a negative offset band to the below-anchor extent', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 0,
        offYEmu: -95250, // -10 px
        cx: 1905000,
        cy: 476250, // 50 px
        text: 'raised',
        wrap: '<wp:wrapTopAndBottom/>',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBe(true)
    expect(box.bandTopPx).toBe(-10)
    expect(box.bandBottomPx).toBe(40)
  })

  it('leaves page-relative topAndBottom anchors on their previous path', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 0,
        offYEmu: 914400,
        cx: 1905000,
        cy: 476250,
        text: 'pagebox',
        wrap: '<wp:wrapTopAndBottom/>',
        relV: 'page',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBeUndefined()
    expect(box.bandBottomPx).toBeUndefined()
  })

  it('keeps wrapNone semantics unchanged (floating, no band)', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 0,
        offYEmu: 190500,
        cx: 1905000,
        cy: 476250,
        text: 'overlay',
        wrap: '<wp:wrapNone/>',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBe(true)
    expect(box.bandBottomPx).toBeUndefined()
  })
})
