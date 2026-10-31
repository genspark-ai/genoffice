import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

// IeumOffice divergence (genoffice #2331): Word places a page/margin-relative
// anchor at its page position whatever the wrap kind; wrap only decides how
// the body text flows around it. Cover pages built from wrapTopAndBottom /
// wrapSquare rectangles were stacked into the flow, shifted by the margins.

const WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'

// build-docx sectPr: A4 (11906x16838 twips), 1440-twip margins
const PAGE_W = 11906 * 635
const PAGE_H = 16838 * 635
const MARGIN = 1440 * 635
const EMU_PER_PX = 9525

function rect(opts: {
  posH: string
  posV: string
  cx: number
  cy: number
  wrap: string
  relH?: string
  relV?: string
  behind?: boolean
}): string {
  const pos = (dir: 'H' | 'V', rel: string, v: string) =>
    `<wp:position${dir} relativeFrom="${rel}">${
      /^-?\d+$/.test(v) ? `<wp:posOffset>${v}</wp:posOffset>` : `<wp:align>${v}</wp:align>`
    }</wp:position${dir}>`
  return (
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="251658240" behindDoc="${opts.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    pos('H', opts.relH ?? 'page', opts.posH) +
    pos('V', opts.relV ?? 'page', opts.posV) +
    `<wp:extent cx="${opts.cx}" cy="${opts.cy}"/>${opts.wrap}<wp:docPr id="1" name="Rect"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp ${WPS_NS}><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${opts.cx}" cy="${opts.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></wps:spPr>` +
    `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
  )
}

const para = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`
const TOP_AND_BOTTOM = '<wp:wrapTopAndBottom/>'
const SQUARE = '<wp:wrapSquare wrapText="bothSides"/>'

async function boxesOf(bodyXml: string) {
  const doc = await parseDocx(await buildDocx({ bodyXml }))
  return doc.blocks.map((b) => b.textboxes ?? [])
}

describe('page-anchored wrapped shapes keep their page position', () => {
  it('floats a wrapTopAndBottom page band at the page top and reserves its body part', async () => {
    const h = 2000000
    const [[band]] = await boxesOf(
      para(rect({ posH: '0', posV: '0', cx: PAGE_W, cy: h, wrap: TOP_AND_BOTTOM })),
    )
    expect(band.floating).toBe(true)
    // column/body-top space: the page edge sits one margin up-left
    expect(band.offsetXEmu).toBe(-MARGIN)
    expect(band.offsetYEmu).toBe(-MARGIN)
    expect(band.pageRelX).toBe(true)
    expect(band.pageRelV).toBe(true)
    expect(band.pageRelVFrom).toBe('page')
    // text resumes below the band's bottom, not below band height + margin
    expect(band.bandBottomPx).toBe(Math.round(-MARGIN / EMU_PER_PX) + Math.round(h / EMU_PER_PX))
  })

  for (const [kind, wrap] of [
    ['wrapSquare', SQUARE],
    [
      'wrapTight',
      '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/></wp:wrapPolygon></wp:wrapTight>',
    ],
    [
      'wrapThrough',
      '<wp:wrapThrough wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/></wp:wrapPolygon></wp:wrapThrough>',
    ],
    ['wrapTopAndBottom', TOP_AND_BOTTOM],
  ] as const) {
    it(`floats a narrow ${kind} page-anchored box in a later paragraph at its page offset`, async () => {
      const blocks = await boxesOf(
        para(`<w:r><w:t>lead</w:t></w:r>`) +
          para(rect({ posH: '914400', posV: '2743200', cx: 4572000, cy: 914400, wrap })),
      )
      const [box] = blocks[1]
      expect(box.floating).toBe(true)
      expect(box.pagePinned).toBeUndefined()
      expect(box.offsetXEmu).toBe(914400 - MARGIN)
      expect(box.offsetYEmu).toBe(2743200 - MARGIN)
      expect(box.pageRelX).toBe(true)
      expect(box.pageRelV).toBe(true)
      // the box lies below the body top: no band blanking the lines above it
      expect(box.bandBottomPx).toBeUndefined()
      expect(box.wrapSide).toBeUndefined()
    })
  }

  it('floats a page-bottom wrapSquare band without pushing the flow off the page', async () => {
    const h = 1500000
    const [[box]] = await boxesOf(
      para(rect({ posH: '0', posV: String(PAGE_H - h), cx: PAGE_W, cy: h, wrap: SQUARE })),
    )
    expect(box.floating).toBe(true)
    expect(box.offsetYEmu).toBe(PAGE_H - h - MARGIN)
    expect(box.pageRelV).toBe(true)
    expect(box.bandBottomPx).toBeUndefined()
  })

  it('resolves wp:align left/top against the page for a first-block anchor', async () => {
    const [[box]] = await boxesOf(
      para(
        rect({
          posH: 'left',
          posV: 'top',
          cx: PAGE_W,
          cy: 2000000,
          wrap: '<wp:wrapNone/>',
          behind: true,
        }),
      ),
    )
    expect(box.floating).toBe(true)
    expect(box.offsetXEmu).toBe(-MARGIN)
    // was missing: the box hung from the anchor paragraph, a margin too low
    expect(box.offsetYEmu).toBe(-MARGIN)
    expect(box.pageRelV).toBe(true)
    expect(box.pageRelVFrom).toBe('page')
  })

  it('resolves wp:align top/bottom with a posOffset X on pinned first-page cover art', async () => {
    const h = 1500000
    const blocks = await boxesOf(
      para(`<w:r><w:t>lead</w:t></w:r>`) +
        para(
          rect({
            posH: '0',
            posV: 'top',
            cx: PAGE_W,
            cy: 2000000,
            wrap: '<wp:wrapNone/>',
            behind: true,
          }),
        ) +
        para(
          rect({
            posH: '0',
            posV: 'bottom',
            cx: PAGE_W,
            cy: h,
            wrap: '<wp:wrapNone/>',
            behind: true,
          }),
        ),
    )
    const [top] = blocks[1]
    const [bottom] = blocks[2]
    expect(top.pagePinned).toBe(true)
    expect(top.offsetYEmu).toBe(0)
    expect(bottom.pagePinned).toBe(true)
    expect(bottom.offsetYEmu).toBe(PAGE_H - h)
  })

  it('resolves wp:align bottom against the margin box for a wrapped box', async () => {
    const h = 914400
    const [[box]] = await boxesOf(
      para(rect({ posH: '0', posV: 'bottom', cx: 1828800, cy: h, wrap: SQUARE, relV: 'margin' })),
    )
    expect(box.floating).toBe(true)
    expect(box.offsetYEmu).toBe(PAGE_H - 2 * MARGIN - h)
    expect(box.pageRelV).toBe(true)
    expect(box.pageRelVFrom).toBe('margin')
  })

  it('leaves a paragraph-anchored wrapSquare box on the flow path (guard)', async () => {
    const [[box]] = await boxesOf(
      para(
        rect({
          posH: '0',
          posV: '0',
          cx: 1828800,
          cy: 914400,
          wrap: SQUARE,
          relH: 'column',
          relV: 'paragraph',
        }),
        `<w:r><w:t>text beside the box</w:t></w:r>`,
      ),
    )
    expect(box.floating).toBeUndefined()
    expect(box.pageRelV).toBeUndefined()
    expect(box.wrapSides).toBe(true)
    expect(box.wrapSide).toBe('left')
  })

  it('leaves a paragraph-anchored wrapTopAndBottom band unchanged (guard)', async () => {
    const [[box]] = await boxesOf(
      para(
        rect({
          posH: '0',
          posV: '190500',
          cx: 1828800,
          cy: 476250,
          wrap: TOP_AND_BOTTOM,
          relH: 'column',
          relV: 'paragraph',
        }),
      ),
    )
    expect(box.floating).toBe(true)
    expect(box.pageRelV).toBeUndefined()
    expect(box.offsetYEmu).toBe(190500)
    expect(box.bandTopPx).toBe(20)
    expect(box.bandBottomPx).toBe(70)
  })
})
