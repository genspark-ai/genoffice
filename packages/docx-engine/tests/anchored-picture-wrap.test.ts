import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

// test sectPr: A4 with 1440-twip margins → column = 9026 twips = 5731510 EMU
const TIGHT =
  '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/>' +
  '<wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/>' +
  '<wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'

function picture(o: {
  x: number
  y: number
  cx: number
  cy: number
  behind?: boolean
  wrap?: string
  relH?: string
}): string {
  return (
    `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="251658240" behindDoc="${o.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="${o.relH ?? 'column'}"><wp:posOffset>${o.x}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${o.y}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${o.cx}" cy="${o.cy}"/>${o.wrap ?? TIGHT}<wp:docPr id="1" name="Picture 1"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${o.cx}" cy="${o.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing>`
  )
}

const cell = (w: number, text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`

/** w:tblpPr table hugging the left margin: page X 689 twips, 5766 twips wide */
const FLOAT_TABLE =
  '<w:tbl><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" ' +
  'w:horzAnchor="page" w:tblpX="689" w:tblpY="331"/><w:tblW w:w="5766" w:type="dxa"/>' +
  '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2883"/><w:gridCol w:w="2883"/></w:tblGrid>' +
  `<w:tr>${cell(2883, 'a')}${cell(2883, 'b')}</w:tr></w:tbl>`

const PLAIN_TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="5766" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2883"/><w:gridCol w:w="2883"/></w:tblGrid>' +
  `<w:tr>${cell(2883, 'a')}${cell(2883, 'b')}</w:tr></w:tbl>`

describe('anchored picture wrapping', () => {
  it('bands behind-text wrapTight pictures whose union covers the column', async () => {
    const para =
      `<w:p><w:r>` +
      picture({ x: -857250, y: 365125, cx: 2868295, cy: 2487930, behind: true }) +
      picture({ x: 1692275, y: 408940, cx: 2861945, cy: 2423795, behind: true }) +
      picture({ x: 4088765, y: 512445, cx: 2832735, cy: 2183765, behind: true }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para, withImage: true }))
    const boxes = doc.blocks[0].textboxes!
    expect(boxes.length).toBe(3)
    for (const box of boxes) {
      expect(box.floating).toBe(true)
      expect(box.behind).toBe(true)
      expect(box.bandBottomPx).toBeGreaterThan(0)
    }
    expect(boxes[0].bandTopPx).toBe(38)
    expect(boxes[0].bandBottomPx).toBe(38 + 261)
  })

  it('bands a column-spanning photo row that shares its paragraph with text', async () => {
    const para =
      `<w:p><w:r><w:t>Preview of flipchart:</w:t></w:r><w:r>` +
      picture({ x: -209550, y: 412115, cx: 3054350, cy: 2293212, behind: true }) +
      picture({ x: 3003550, y: 405765, cx: 3215640, cy: 2280285, behind: true, relH: 'margin' }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para, withImage: true }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.textboxes?.length).toBe(2)
    expect(block.textboxes?.every((b) => b.floating && b.bandBottomPx! > 0)).toBe(true)
    expect(block.strayRuns?.map((r) => r.text).join('')).toBe('Preview of flipchart:')

    const narrow =
      `<w:p><w:r><w:t>Logos:</w:t></w:r><w:r>` +
      picture({ x: 0, y: 0, cx: 914400, cy: 914400 }) +
      picture({ x: 4800000, y: 0, cx: 914400, cy: 914400 }) +
      `</w:r></w:p>`
    const editable = await parseDocx(await buildDocx({ bodyXml: narrow, withImage: true }))
    expect(editable.blocks[0].type).toBe('paragraph')
    expect(editable.blocks[0].runs?.filter((r) => r.image).length).toBe(2)
  })

  it('keeps behind-text wrapNone pictures as zero-footprint overlays', async () => {
    const para =
      `<w:p><w:r>` +
      picture({ x: 0, y: 0, cx: 5731510, cy: 952500, behind: true, wrap: '<wp:wrapNone/>' }) +
      picture({ x: 0, y: 1905000, cx: 952500, cy: 952500, behind: true, wrap: '<wp:wrapNone/>' }) +
      `</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para, withImage: true }))
    for (const box of doc.blocks[0].textboxes!) {
      expect(box.floating).toBe(true)
      expect(box.bandBottomPx).toBeUndefined()
    }
  })

  it('reserves the band of side-wrapped pictures when a table follows them', async () => {
    const pics =
      `<w:p><w:r>` +
      picture({ x: 634365, y: 5080, cx: 1828800, cy: 1654175, behind: true }) +
      picture({ x: 3565525, y: 247650, cx: 1944370, cy: 1293495, behind: true }) +
      `</w:r></w:p>`
    const withTable = await parseDocx(
      await buildDocx({ bodyXml: pics + '<w:p/><w:p/>' + PLAIN_TABLE, withImage: true }),
    )
    const [left, right] = withTable.blocks[0].textboxes!
    expect(left.bandTopPx).toBe(1)
    expect(left.bandBottomPx).toBe(1 + 174)
    expect(right.bandBottomPx).toBe(26 + 136)

    const withText = await parseDocx(
      await buildDocx({ bodyXml: pics + '<w:p><w:r><w:t>text</w:t></w:r></w:p>', withImage: true }),
    )
    for (const box of withText.blocks[0].textboxes!) expect(box.bandBottomPx).toBeUndefined()
  })

  it('bands a side-wrapped picture that cannot sit beside a floating table', async () => {
    const wide = `<w:p><w:r>${picture({ x: 3307080, y: 410210, cx: 3159760, cy: 3214370 })}</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: FLOAT_TABLE + wide, withImage: true }))
    expect(doc.blocks[1].type).toBe('image')
    expect(doc.blocks[1].imageWrap).toBe('tight-right')
    expect(doc.blocks[1].imageBand).toBe(true)

    const narrow = `<w:p><w:r>${picture({ x: 4500000, y: 410210, cx: 800000, cy: 800000 })}</w:r></w:p>`
    const roomy = await parseDocx(
      await buildDocx({ bodyXml: FLOAT_TABLE + narrow, withImage: true }),
    )
    expect(roomy.blocks[1].imageBand).toBeUndefined()

    const inFlow = await parseDocx(
      await buildDocx({ bodyXml: PLAIN_TABLE + wide, withImage: true }),
    )
    expect(inFlow.blocks[1].imageBand).toBeUndefined()

    // omitted w:horzAnchor is page-relative like an explicit one
    const implicitPage = FLOAT_TABLE.replace(' w:horzAnchor="page"', '')
    const implicit = await parseDocx(
      await buildDocx({ bodyXml: implicitPage + wide, withImage: true }),
    )
    expect(implicit.blocks[1].imageBand).toBe(true)

    // a content control between the floating table and the picture ends the pairing
    const sdt = '<w:sdt><w:sdtContent><w:p><w:r><w:t>note</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    const gapped = await parseDocx(
      await buildDocx({ bodyXml: FLOAT_TABLE + sdt + wide, withImage: true }),
    )
    expect(gapped.blocks[2].imageBand).toBeUndefined()
  })
})

describe('anchor paragraph line', () => {
  const SQUARE = '<wp:wrapSquare wrapText="bothSides"/>'
  const pic = () => picture({ x: 5214099, y: -899793, cx: 2143760, cy: 1614170, wrap: SQUARE })

  it('keeps the empty anchor line of a side-wrapped picture, sized by the mark only', async () => {
    const para =
      `<w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/><w:rPr/></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="36"/></w:rPr>${pic()}</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para, withImage: true }))
    const block = doc.blocks[0]
    expect(block.type).toBe('image')
    expect(block.imageWrap).toMatch(/^square-(?:left|right)$/)
    expect(block.anchorLine?.format?.lineSpacing).toBeCloseTo(1.15, 3)
    expect(block.anchorLine?.format?.emptyRunSizeHalfPoints).toBeUndefined()
  })

  it('takes the paragraph-mark w:sz and style for the anchor line', async () => {
    const para =
      `<w:p><w:pPr><w:pStyle w:val="Caption"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="36"/></w:rPr>${pic()}</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para, withImage: true }))
    expect(doc.blocks[0].anchorLine).toMatchObject({
      styleId: 'Caption',
      format: { emptyRunSizeHalfPoints: 20 },
    })
  })

  it('adds no anchor line for behind-text or inline pictures', async () => {
    const behind =
      `<w:p><w:r>` +
      picture({ x: 0, y: 0, cx: 914400, cy: 914400, behind: true, wrap: '<wp:wrapNone/>' }) +
      `</w:r></w:p>`
    const inline =
      `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="914400" cy="914400"/><wp:docPr id="2" name="Picture 2"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture 2"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: behind + inline, withImage: true }))
    expect(doc.blocks[0].anchorLine).toBeUndefined()
    expect(doc.blocks[1].anchorLine).toBeUndefined()
  })

  it('keeps the line of a paragraph holding only VML shapetype definitions', async () => {
    const para =
      `<w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr><w:r><w:pict>` +
      `<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" path="m@4@5l@4@11@9@11@9@5xe">` +
      `<v:stroke joinstyle="miter"/></v:shapetype></w:pict></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.invisibleMarker).toBe(true)
    expect(block.anchorLine?.format?.lineSpacing).toBeCloseTo(1.15, 3)
  })
})
