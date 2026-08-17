import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
const WPG_NS = 'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"'
const MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
const WP14_NS = 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"'

// build-docx sectPr: A4 (11906x16838 twips), 1440-twip margins
const PAGE_W = 11906 * 635
const PAGE_H = 16838 * 635
const MARGIN = 1440 * 635

function textWsp(opts: {
  offX: number
  offY: number
  cx: number
  cy: number
  text: string
  bodyPrAttrs?: string
}): string {
  return (
    `<wps:wsp ${WPS_NS}><wps:cNvSpPr txBox="1"/><wps:spPr>` +
    `<a:xfrm><a:off x="${opts.offX}" y="${opts.offY}"/><a:ext cx="${opts.cx}" cy="${opts.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
    `</wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>${opts.text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr${opts.bodyPrAttrs ?? ''}/></wps:wsp>`
  )
}

function anchorPara(
  positionHXml: string,
  positionVXml: string,
  inner: string,
  ext: string,
): string {
  return (
    `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>${positionHXml}${positionVXml}${ext}` +
    `<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Box 1"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    inner +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
  )
}

describe('page/margin-anchored textboxes outside the body column', () => {
  it('floats a page-pct-anchored sidebar group at the resolved page position', async () => {
    // mirrors Word resume templates: pctPosHOffset in mc:Choice, posOffset in
    // the stripped mc:Fallback, positionV margin/align=top, full-height group
    const posH =
      `<mc:AlternateContent ${MC_NS}><mc:Choice Requires="wp14">` +
      `<wp:positionH relativeFrom="page"><wp14:pctPosHOffset ${WP14_NS}>2000</wp14:pctPosHOffset></wp:positionH>` +
      `</mc:Choice><mc:Fallback><wp:positionH relativeFrom="page"><wp:posOffset>151206</wp:posOffset></wp:positionH></mc:Fallback></mc:AlternateContent>`
    const posV = `<wp:positionV relativeFrom="margin"><wp:align>top</wp:align></wp:positionV>`
    const group =
      `<wpg:wgp ${WPG_NS} ${WPS_NS}><wpg:cNvGrpSpPr/><wpg:grpSpPr>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="600000" cy="8000000"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="600000" cy="8000000"/></a:xfrm></wpg:grpSpPr>` +
      textWsp({ offX: 0, offY: 0, cx: 600000, cy: 4000000, text: 'name' }) +
      textWsp({
        offX: 0,
        offY: 4400000,
        cx: 600000,
        cy: 3600000,
        text: 'address',
        bodyPrAttrs: ' anchor="b"',
      }) +
      `</wpg:wgp>`
    const para = anchorPara(posH, posV, group, '<wp:extent cx="600000" cy="8000000"/>').replace(
      'wordprocessingShape">',
      'wordprocessingGroup">',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const boxes = doc.blocks[0].textboxes
    expect(boxes?.length).toBe(2)
    const [name, address] = boxes!
    const anchorX = Math.round((PAGE_W * 2000) / 100000) - MARGIN
    expect(name.floating).toBe(true)
    expect(name.offsetXEmu).toBe(anchorX)
    expect(name.offsetYEmu).toBe(0)
    expect(address.floating).toBe(true)
    expect(address.offsetXEmu).toBe(anchorX)
    expect(address.offsetYEmu).toBe(4400000)
    expect(address.vAlign).toBe('bottom')
    expect(name.vAlign).toBeUndefined()
  })

  it('keeps in-column wrapSquare boxes in the flow', async () => {
    const posH = `<wp:positionH relativeFrom="page"><wp14:pctPosHOffset ${WP14_NS}>20000</wp14:pctPosHOffset></wp:positionH>`
    const posV = `<wp:positionV relativeFrom="margin"><wp:align>top</wp:align></wp:positionV>`
    const para = anchorPara(
      posH,
      posV,
      textWsp({ offX: 0, offY: 0, cx: 600000, cy: 400000, text: 'callout' }),
      '<wp:extent cx="600000" cy="400000"/>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBeUndefined()
    expect(box.offsetXEmu).toBeUndefined()
  })

  it('resolves page-relative wp:align right/bottom against the page geometry', async () => {
    const posH = `<wp:positionH relativeFrom="page"><wp:align>right</wp:align></wp:positionH>`
    const posV = `<wp:positionV relativeFrom="margin"><wp:align>bottom</wp:align></wp:positionV>`
    const para = anchorPara(
      posH,
      posV,
      textWsp({ offX: 0, offY: 0, cx: 600000, cy: 1000000, text: 'side' }),
      '<wp:extent cx="600000" cy="1000000"/>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBe(true)
    expect(box.offsetXEmu).toBe(PAGE_W - 600000 - MARGIN)
    expect(box.offsetYEmu).toBe(PAGE_H - 2 * MARGIN - 1000000)
  })

  it('leaves plain posOffset anchors on the legacy paragraph-relative path', async () => {
    const posH = `<wp:positionH relativeFrom="page"><wp:posOffset>100000</wp:posOffset></wp:positionH>`
    const posV = `<wp:positionV relativeFrom="paragraph"><wp:posOffset>50000</wp:posOffset></wp:positionV>`
    const para = anchorPara(
      posH,
      posV,
      textWsp({ offX: 0, offY: 0, cx: 600000, cy: 400000, text: 'legacy' }),
      '<wp:extent cx="600000" cy="400000"/>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const [box] = doc.blocks[0].textboxes!
    expect(box.floating).toBeUndefined()
    expect(box.offsetXEmu).toBe(100000)
    expect(box.offsetYEmu).toBe(50000)
  })

  it('resolves SDT-wrapped anchors against their own section, not the first one', async () => {
    // section 1: wide page with narrow margins — an anchor resolved against it
    // would land inside the column and stay in the flow
    const sect1Para =
      `<w:p><w:pPr><w:sectPr>` +
      `<w:pgSz w:w="20000" w:h="16838"/>` +
      `<w:pgMar w:top="1440" w:right="720" w:bottom="1440" w:left="720" w:header="708" w:footer="708" w:gutter="0"/>` +
      `</w:sectPr></w:pPr></w:p>`
    const posH = `<wp:positionH relativeFrom="page"><wp14:pctPosHOffset ${WP14_NS}>2000</wp14:pctPosHOffset></wp:positionH>`
    const posV = `<wp:positionV relativeFrom="margin"><wp:align>top</wp:align></wp:positionV>`
    const anchored = anchorPara(
      posH,
      posV,
      textWsp({ offX: 0, offY: 0, cx: 600000, cy: 400000, text: 'sidebar' }),
      '<wp:extent cx="600000" cy="400000"/>',
    )
    const bodyXml =
      `<w:p><w:r><w:t>section one</w:t></w:r></w:p>${sect1Para}` +
      `<w:sdt><w:sdtPr><w:id w:val="42"/></w:sdtPr><w:sdtContent>${anchored}</w:sdtContent></w:sdt>`
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    const box = doc.blocks.find((b) => b.textboxes)?.textboxes?.[0]
    // section 2 (trailing A4 sectPr) geometry, not section 1's
    expect(box?.floating).toBe(true)
    expect(box?.offsetXEmu).toBe(Math.round((PAGE_W * 2000) / 100000) - MARGIN)
  })
})
