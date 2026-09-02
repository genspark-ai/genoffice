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
  relH?: string
  relV?: string
  inner?: string
}): string {
  return (
    `<w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="${opts.relH ?? 'column'}"><wp:posOffset>${opts.offXEmu}</wp:posOffset></wp:positionH>` +
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

  // body width in the test sectPr: 11906 − 2×1440 = 9026 twips = 5731510 EMU
  it('reserves the band for a column-spanning wrapSquare box in a multi-drawing paragraph', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 60000,
        offYEmu: 190500, // 20 px
        cx: 5600000, // leaves < 36px beside the box
        cy: 476250, // 50 px
        text: 'wide',
        wrap: '<wp:wrapSquare wrapText="bothSides"/>',
      }) +
      drawing({
        offXEmu: 0,
        offYEmu: 1905000,
        cx: 1905000,
        cy: 476250,
        text: 'other',
        wrap: '<wp:wrapNone/>',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [wide] = doc.blocks[0].textboxes!
    expect(wide.floating).toBe(true)
    expect(wide.bandTopPx).toBe(20)
    expect(wide.bandBottomPx).toBe(70)
  })

  it('keeps a narrow wrapSquare box bandless in a multi-drawing paragraph', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 0,
        offYEmu: 190500,
        cx: 1905000, // wide sliver remains: text wraps beside it in Word
        cy: 476250,
        text: 'narrow',
        wrap: '<wp:wrapSquare wrapText="bothSides"/>',
      }) +
      drawing({
        offXEmu: 0,
        offYEmu: 1905000,
        cx: 1905000,
        cy: 476250,
        text: 'other',
        wrap: '<wp:wrapNone/>',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [narrow] = doc.blocks[0].textboxes!
    expect(narrow.floating).toBe(true)
    expect(narrow.bandBottomPx).toBeUndefined()
  })

  // page-relative posOffset already lands in column space by parse time; the
  // span test must not subtract the margin a second time
  it('bands a page-relative wrapSquare box that covers the column', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 914400 + 60000, // margin (1440tw) + 60000 from the column left
        offYEmu: 190500, // 20 px
        cx: 5600000,
        cy: 476250, // 50 px
        text: 'page wide',
        wrap: '<wp:wrapSquare wrapText="bothSides"/>',
        relH: 'page',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBe(true)
    expect(box.bandTopPx).toBe(20)
    expect(box.bandBottomPx).toBe(70)
  })

  it('unions mixed page- and column-relative anchors in one coordinate space', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 914400, // page-relative left half: column x = 0
        offYEmu: 190500, // 20 px
        cx: 2900000,
        cy: 476250, // 50 px
        text: 'left half',
        wrap: '<wp:wrapSquare wrapText="bothSides"/>',
        relH: 'page',
      }) +
      drawing({
        offXEmu: 2900000,
        offYEmu: 190500,
        cx: 2831510, // right half: no sliver remains
        cy: 476250,
        text: 'right half',
        wrap: '<wp:wrapSquare wrapText="bothSides"/>',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [left, right] = doc.blocks[0].textboxes!
    expect(left.bandTopPx).toBe(20)
    expect(right.bandTopPx).toBe(20)
    // a union band never overflows the page bottom
    expect(left.bandOverflow).toBeUndefined()
    expect(right.bandOverflow).toBeUndefined()
  })

  it('floats a lone column-spanning wrapSquare box with its band', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 60000,
        offYEmu: 190500, // 20 px
        cx: 5600000,
        cy: 476250, // 50 px
        text: 'solo wide',
        wrap: '<wp:wrapSquare wrapText="bothSides"/>',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBe(true)
    expect(box.bandTopPx).toBe(20)
    expect(box.bandBottomPx).toBe(70)
  })

  it('bands a wrapSquare box when the paragraph union of anchors covers the column', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 60000,
        offYEmu: 190500, // 20 px
        cx: 2800000, // left half
        cy: 476250, // 50 px
        text: 'left frame',
        wrap: '<wp:wrapSquare wrapText="bothSides"/>',
      }) +
      drawing({
        offXEmu: 2860000,
        offYEmu: 190500,
        cx: 2800000, // right half: union leaves < 36px of column
        cy: 476250,
        text: 'right frame',
        wrap: '<wp:wrapNone/>',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [left, right] = doc.blocks[0].textboxes!
    expect(left.floating).toBe(true)
    expect(left.bandTopPx).toBe(20)
    expect(left.bandBottomPx).toBe(70)
    // union spanning is not own spanning: the band is pushed whole, not overflowed
    expect(left.bandOverflow).toBeUndefined()
    expect(right.floating).toBe(true)
    expect(right.bandBottomPx).toBeUndefined()
  })

  it('flags the overflow band only for a box spanning the column by itself', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 60000,
        offYEmu: 190500,
        cx: 5600000,
        cy: 476250,
        text: 'solo wide',
        wrap: '<wp:wrapSquare wrapText="bothSides"/>',
      }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.bandOverflow).toBe(true)
  })

  it('keeps a lone narrow wrapSquare box on the flow path', async () => {
    const para =
      `<w:p><w:r>` +
      drawing({
        offXEmu: 60000,
        offYEmu: 190500,
        cx: 1905000, // wide sliver remains beside the box
        cy: 476250,
        text: 'solo narrow',
        wrap: '<wp:wrapSquare wrapText="bothSides"/>',
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
