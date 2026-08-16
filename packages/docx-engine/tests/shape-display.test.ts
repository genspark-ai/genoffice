import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
const WPG_NS = 'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"'

const STYLE_REFS =
  '<wps:style>' +
  '<a:lnRef idx="2"><a:srgbClr val="123A5B"/></a:lnRef>' +
  '<a:fillRef idx="1"><a:srgbClr val="2E6E9E"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:srgbClr val="000000"/></a:effectRef>' +
  '<a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef>' +
  '</wps:style>'

function wsp(opts: {
  prst: string
  offX?: number
  offY?: number
  cx: number
  cy: number
  xfrmAttrs?: string
  ln?: string
  text?: string
}): string {
  const txbx = opts.text
    ? `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${opts.text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>`
    : ''
  return (
    `<wps:wsp><wps:cNvSpPr/><wps:spPr>` +
    `<a:xfrm${opts.xfrmAttrs ?? ''}><a:off x="${opts.offX ?? 0}" y="${opts.offY ?? 0}"/>` +
    `<a:ext cx="${opts.cx}" cy="${opts.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${opts.prst}"><a:avLst/></a:prstGeom>${opts.ln ?? ''}` +
    `</wps:spPr>${STYLE_REFS}${txbx}<wps:bodyPr/></wps:wsp>`
  )
}

function anchorParagraph(
  inner: string,
  opts?: { posX?: number; posY?: number; extCx?: number; extCy?: number },
): string {
  return (
    `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${opts?.posX ?? 0}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${opts?.posY ?? 0}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${opts?.extCx ?? 952500}" cy="${opts?.extCy ?? 952500}"/><wp:wrapNone/>` +
    `<wp:docPr id="9" name="Shape 9"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    inner.replace('<wps:wsp>', `<wps:wsp ${WPS_NS}>`) +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
  )
}

describe('textless preset shapes render as display boxes', () => {
  it('a star5 with wps:style colors becomes a floating readOnly box', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(wsp({ prst: 'star5', cx: 952500, cy: 857250 }), {
          posX: 190500,
          posY: 95250,
        }),
      }),
    )
    const box = doc.blocks[0].textboxes?.[0]
    expect(box).toBeTruthy()
    expect(box!.prst).toBe('star5')
    expect(box!.readOnly).toBe(true)
    expect(box!.floating).toBe(true)
    expect(box!.fill).toBe('2E6E9E')
    expect(box!.borderColor).toBe('123A5B')
    expect(box!.widthPx).toBe(100)
    expect(box!.heightPx).toBe(90)
    expect(box!.offsetXEmu).toBe(190500)
    expect(box!.offsetYEmu).toBe(95250)
  })

  it('a triangle keeps its preset and a text-bearing ellipse keeps its paras', async () => {
    const body =
      anchorParagraph(wsp({ prst: 'triangle', cx: 952500, cy: 857250 })) +
      anchorParagraph(wsp({ prst: 'ellipse', cx: 1857375, cy: 876300, text: 'inside' }))
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(doc.blocks[0].textboxes?.[0]?.prst).toBe('triangle')
    const ellipse = doc.blocks[1].textboxes?.[0]
    expect(ellipse?.prst).toBe('ellipse')
    expect(ellipse?.paras[0]?.runs[0]?.text).toBe('inside')
    expect(ellipse?.readOnly).toBeUndefined()
  })
})

describe('anchored connectors', () => {
  it('a flipV straightConnector with a tail arrow renders as a diagonal lineArrow', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(
          wsp({
            prst: 'straightConnector1',
            cx: 2362200,
            cy: 390525,
            xfrmAttrs: ' flipV="1"',
            ln: '<a:ln><a:tailEnd type="triangle"/></a:ln>',
          }),
        ),
      }),
    )
    const box = doc.blocks[0].textboxes?.[0]
    expect(box).toBeTruthy()
    expect(box!.prst).toBe('lineArrow')
    expect(box!.lineDiag).toBe(true)
    expect(box!.flipV).toBe(true)
    // stroke falls back to the wps:style lnRef color
    expect(box!.borderColor).toBe('123A5B')
  })

  it('a head-only arrow reverses the segment so the arrowhead sits at the start', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(
          wsp({
            prst: 'straightConnector1',
            cx: 2362200,
            cy: 390525,
            ln: '<a:ln><a:headEnd type="triangle"/></a:ln>',
          }),
        ),
      }),
    )
    const box = doc.blocks[0].textboxes?.[0]
    expect(box!.prst).toBe('lineArrow')
    // (0,0)→(w,h) with the arrow at (0,0) = the reversed segment
    expect(box!.flipH).toBe(true)
    expect(box!.flipV).toBe(true)
  })

  it('head-only on a flipV connector reverses onto the other diagonal', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(
          wsp({
            prst: 'straightConnector1',
            cx: 2362200,
            cy: 390525,
            xfrmAttrs: ' flipV="1"',
            ln: '<a:ln><a:headEnd type="triangle"/></a:ln>',
          }),
        ),
      }),
    )
    const box = doc.blocks[0].textboxes?.[0]
    expect(box!.flipH).toBe(true)
    expect(box!.flipV).toBeUndefined()
    expect(box!.lineDiag).toBe(true)
  })

  it('a shallow flipped connector still renders as a line box', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(
          wsp({ prst: 'straightConnector1', cx: 5230495, cy: 20955, xfrmAttrs: ' flipV="1"' }),
          { extCx: 5230495, extCy: 20955 },
        ),
      }),
    )
    const box = doc.blocks[0].textboxes?.[0]
    expect(box).toBeTruthy()
    expect(box!.prst).toBe('line')
    expect(box!.flipV).toBe(true)
    expect(box!.lineDiag).toBe(true)
  })

  it('a shallow arrowed connector still renders as a line box', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(
          wsp({
            prst: 'straightConnector1',
            cx: 5230495,
            cy: 20955,
            ln: '<a:ln><a:tailEnd type="triangle"/></a:ln>',
          }),
          { extCx: 5230495, extCy: 20955 },
        ),
      }),
    )
    expect(doc.blocks[0].textboxes?.[0]?.prst).toBe('lineArrow')
  })

  it('a zero-height flipped arrow keeps its flip (level line, arrow at the left tip)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(
          wsp({
            prst: 'straightConnector1',
            cx: 2362200,
            cy: 0,
            xfrmAttrs: ' flipH="1"',
            ln: '<a:ln><a:tailEnd type="triangle"/></a:ln>',
          }),
          { extCx: 2362200, extCy: 0 },
        ),
      }),
    )
    const box = doc.blocks[0].textboxes?.[0]
    expect(box).toBeTruthy()
    expect(box!.prst).toBe('lineArrow')
    expect(box!.heightPx).toBe(12)
    expect(box!.flipH).toBe(true)
    // level line: no diagonal, the renderer reverses via flipH alone
    expect(box!.lineDiag).toBeUndefined()
  })

  it('head-only on a zero-height flipH arrow cancels the reversal (arrow at the right tip)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(
          wsp({
            prst: 'straightConnector1',
            cx: 2362200,
            cy: 0,
            xfrmAttrs: ' flipH="1"',
            ln: '<a:ln><a:headEnd type="triangle"/></a:ln>',
          }),
          { extCx: 2362200, extCy: 0 },
        ),
      }),
    )
    const box = doc.blocks[0].textboxes?.[0]
    expect(box!.prst).toBe('lineArrow')
    // headEnd sits at the segment start; flipH puts that start at the right
    // tip, so the reversed-segment representation has no flipH left
    expect(box!.flipH).toBeUndefined()
    expect(box!.flipV).toBe(true)
  })

  it('a near-flat anchored line stays on the decorative thin-rule path', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(wsp({ prst: 'line', cx: 5230495, cy: 20955 }), {
          extCx: 5230495,
          extCy: 20955,
        }),
      }),
    )
    expect(doc.blocks[0].textboxes).toBeUndefined()
    expect(doc.blocks[0].decorative).toBe(true)
  })
})

describe('wpg group children', () => {
  const GROUP_PARA =
    `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>100000</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>200000</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="1476375" cy="723900"/><wp:wrapNone/>` +
    `<wp:docPr id="14" name="Group 13"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">` +
    `<wpg:wgp ${WPG_NS} ${WPS_NS}><wpg:cNvGrpSpPr/><wpg:grpSpPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="1476375" cy="723900"/>` +
    `<a:chOff x="1323975" y="4733925"/><a:chExt cx="1476375" cy="723900"/></a:xfrm>` +
    `</wpg:grpSpPr>` +
    wsp({ prst: 'star5', offX: 1323975, offY: 4733925, cx: 762000, cy: 723900 }) +
    wsp({ prst: 'rightArrow', offX: 2038350, offY: 4733925, cx: 762000, cy: 723900, text: 'go' }) +
    `</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`

  it('maps child offsets through chOff and adds the anchor position', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: GROUP_PARA }))
    const boxes = doc.blocks[0].textboxes
    expect(boxes?.length).toBe(2)
    const [star, arrow] = boxes!
    // both children float at anchor + (childOff - chOff)
    expect(star.floating).toBe(true)
    expect(star.offsetXEmu).toBe(100000)
    expect(star.offsetYEmu).toBe(200000)
    expect(arrow.prst).toBe('rightArrow')
    expect(arrow.offsetXEmu).toBe(100000 + (2038350 - 1323975))
    expect(arrow.offsetYEmu).toBe(200000)
    // document order preserved for the patch-save w:txbxContent mapping
    expect(arrow.paras[0]?.runs[0]?.text).toBe('go')
  })

  it('scales child geometry when ext differs from chExt', async () => {
    const scaled = GROUP_PARA.replace(
      '<a:ext cx="1476375" cy="723900"/><a:chOff',
      '<a:ext cx="738187" cy="361950"/><a:chOff',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: scaled }))
    const [star, arrow] = doc.blocks[0].textboxes!
    const sx = 738187 / 1476375
    const sy = 361950 / 723900
    expect(star.widthPx).toBe(Math.round(Math.round(762000 / 9525) * sx))
    expect(star.heightPx).toBe(Math.round(Math.round(723900 / 9525) * sy))
    expect(arrow.offsetXEmu).toBe(100000 + Math.round((2038350 - 1323975) * sx))
  })
})
