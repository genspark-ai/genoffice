import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

// test sectPr: A4 with 1440-twip margins → column = 9026 twips = 5731510 EMU
const COL_EMU = 5731510
const MARGIN_EMU = 914400

function callout(o: { x: number; cx: number; relH?: string; y?: number }): string {
  return (
    `<w:p><w:r><w:drawing>` +
    `<wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="${o.relH ?? 'column'}"><wp:posOffset>${o.x}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${o.y ?? -900000}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${o.cx}" cy="1000000"/><wp:wrapSquare wrapText="bothSides"/>` +
    `<wp:docPr id="7" name="Zone de texte 2"/><a:graphic>` +
    `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>` +
    `<a:ext cx="${o.cx}" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>` +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>Callout</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr wrap="square" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>` +
    `<w:r><w:t>Body text of the anchor paragraph.</w:t></w:r></w:p>`
  )
}

describe('in-column wrapSquare textbox: side wrap', () => {
  it('floats a right-hand box on the right, keeping its page-relative X', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: callout({ relH: 'page', x: MARGIN_EMU + 3500000, cx: 2000000 }) }),
    )
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBeUndefined()
    expect(box.bandBottomPx).toBeUndefined()
    expect(box.wrapSide).toBe('right')
    expect(box.wrapEdgePx).toBeCloseTo((COL_EMU - 3500000 - 2000000) / 9525, 1)
    expect(box.wrapGapPx).toBeCloseTo(12, 1)
  })

  it('floats a left-hand box on the left', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: callout({ x: 0, cx: 2000000 }) }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.wrapSide).toBe('left')
    expect(box.wrapEdgePx).toBe(0)
  })

  it('leaves a column-spanning box to the band path', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: callout({ x: 0, cx: 5600000 }) }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.wrapSide).toBeUndefined()
    expect(box.floating).toBe(true)
  })
})
