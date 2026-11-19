import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
// build-docx sectPr: 1440-twip margins
const MARGIN_EMU = 1440 * 635

function anchorPara(relV: string, offY: number): string {
  return (
    `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>28575</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="${relV}"><wp:posOffset>${offY}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="1800000" cy="400000"/><wp:wrapNone/><wp:docPr id="1" name="Box 1"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp ${WPS_NS}><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="400000"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>` +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>band</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r><w:r><w:t>anchor text</w:t></w:r></w:p>`
  )
}

describe('page/margin-relative posOffset V keeps its reference edge', () => {
  it('page anchors: content offset from the pgMar top, flagged pageRelVFrom=page', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: anchorPara('page', 1075478) }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBe(true)
    expect(box.pageRelV).toBe(true)
    expect(box.pageRelVFrom).toBe('page')
    expect(box.offsetYEmu).toBe(1075478 - MARGIN_EMU)
  })

  it('margin anchors: raw offset, flagged pageRelVFrom=margin', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: anchorPara('margin', 300000) }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.pageRelV).toBe(true)
    expect(box.pageRelVFrom).toBe('margin')
    expect(box.offsetYEmu).toBe(300000)
  })

  it('flags a snapToGrid=0 anchor paragraph so its own line stays natural', async () => {
    const nosnap = anchorPara('page', 15875).replace(
      '<w:p><w:r>',
      '<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr><w:r>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: nosnap + anchorPara('page', 818678) }))
    expect(doc.blocks[0].anchorSnapToGrid).toBe(false)
    expect(doc.blocks[1].anchorSnapToGrid).toBeUndefined()
  })

  it('paragraph anchors are not page-relative', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: anchorPara('paragraph', 153247) }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.pageRelV).toBeUndefined()
    expect(box.pageRelVFrom).toBeUndefined()
  })
})
