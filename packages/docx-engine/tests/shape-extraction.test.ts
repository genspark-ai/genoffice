import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
const WPG_NS = 'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"'
const MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'

/** rectangle custGeom in a w×h path space (the shape Word writes for cards/rules) */
function custGeomXml(w: number, h: number): string {
  return (
    `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>` +
    `<a:pathLst><a:path w="${w}" h="${h}">` +
    `<a:moveTo><a:pt x="0" y="${h}"/></a:moveTo>` +
    `<a:lnTo><a:pt x="${w}" y="${h}"/></a:lnTo>` +
    `<a:lnTo><a:pt x="${w}" y="0"/></a:lnTo>` +
    `<a:lnTo><a:pt x="0" y="0"/></a:lnTo>` +
    `<a:close/></a:path></a:pathLst></a:custGeom>`
  )
}

function custWsp(opts: {
  offX?: number
  offY?: number
  cx: number
  cy: number
  fill?: string
}): string {
  const fill = opts.fill === 'grp' ? '<a:grpFill/>' : opts.fill
  return (
    `<wps:wsp ${WPS_NS}><wps:cNvSpPr/><wps:spPr>` +
    `<a:xfrm><a:off x="${opts.offX ?? 0}" y="${opts.offY ?? 0}"/>` +
    `<a:ext cx="${opts.cx}" cy="${opts.cy}"/></a:xfrm>` +
    custGeomXml(opts.cx, opts.cy) +
    (fill ?? '') +
    `</wps:spPr><wps:bodyPr/></wps:wsp>`
  )
}

function anchorParagraph(inner: string, uri: string): string {
  return (
    `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>100000</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>200000</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="952500" cy="952500"/><wp:wrapNone/>` +
    `<wp:docPr id="7" name="Shape 7"/>` +
    `<a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing></w:r></w:p>`
  )
}

const SHAPE_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
const GROUP_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'

describe('shape-style fontRef supplies the text color', () => {
  /** Gallery shape: accent fill from spPr, text color left to the style's fontRef. */
  const styledWsp = (fontRefClr: string) =>
    `<wps:wsp ${WPS_NS}><wps:cNvSpPr/><wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>` +
    `<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>` +
    `<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>` +
    `<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>` +
    `<a:fontRef idx="minor"><a:schemeClr val="${fontRefClr}"/></a:fontRef></wps:style>` +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>Label</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr anchor="ctr"/></wps:wsp>`

  const boxOf = async (xml: string) =>
    (await parseDocx(await buildDocx({ bodyXml: anchorParagraph(xml, SHAPE_URI) }))).blocks[0]
      .textboxes?.[0]

  // Word writes no color on the runs of a gallery shape; lt1 is why the default
  // blue shape shows white text in Word and PowerPoint alike.
  it('fontRef lt1 becomes the box text color', async () => {
    expect((await boxOf(styledWsp('lt1')))?.textColor).toBe('FFFFFF')
  })

  it('fontRef tx1 gives dark text instead', async () => {
    expect((await boxOf(styledWsp('tx1')))?.textColor).toBe('000000')
  })

  it('a shape with no style block has no box-level color to impose', async () => {
    const plain = styledWsp('lt1').replace(/<wps:style>[\s\S]*?<\/wps:style>/, '')
    expect((await boxOf(plain))?.textColor).toBeUndefined()
  })

  it('anchor="ctr" still reaches vAlign alongside the color', async () => {
    expect((await boxOf(styledWsp('lt1')))?.vAlign).toBe('center')
  })
})

describe('textless custGeom shapes', () => {
  it('a solid-filled custGeom becomes a readOnly box with normalized pathData', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(
          custWsp({
            cx: 952500,
            cy: 476250,
            fill: '<a:solidFill><a:srgbClr val="1A295D"/></a:solidFill>',
          }),
          SHAPE_URI,
        ),
      }),
    )
    const box = doc.blocks[0].textboxes?.[0]
    expect(box).toBeTruthy()
    expect(box!.readOnly).toBe(true)
    expect(box!.fill).toBe('1A295D')
    expect(box!.pathData?.path).toBe('M 0 1 L 1 1 L 1 0 L 0 0 Z')
    expect(box!.widthPx).toBe(100)
    expect(box!.heightPx).toBe(50)
  })

  it('a custGeom with no fill and no outline still renders nothing', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: anchorParagraph(custWsp({ cx: 952500, cy: 476250 }), SHAPE_URI),
      }),
    )
    expect(doc.blocks[0].textboxes).toBeUndefined()
  })

  it('a:grpFill inherits the enclosing wpg group fill through nesting', async () => {
    const group =
      `<wpg:wgp ${WPG_NS} ${WPS_NS}><wpg:cNvGrpSpPr/><wpg:grpSpPr>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm>` +
      `<a:solidFill><a:srgbClr val="0070C0"/></a:solidFill>` +
      `</wpg:grpSpPr>` +
      custWsp({ offX: 190500, offY: 95250, cx: 476250, cy: 476250, fill: 'grp' }) +
      `</wpg:wgp>`
    const doc = await parseDocx(await buildDocx({ bodyXml: anchorParagraph(group, GROUP_URI) }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box).toBeTruthy()
    expect(box!.fill).toBe('0070C0')
    expect(box!.floating).toBe(true)
    expect(box!.offsetXEmu).toBe(100000 + 190500)
    expect(box!.offsetYEmu).toBe(200000 + 95250)
  })

  it('a page-relative posOffset group anchors its children column-relative', async () => {
    const group =
      `<wpg:wgp ${WPG_NS} ${WPS_NS}><wpg:cNvGrpSpPr/><wpg:grpSpPr>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="5867400" cy="4785360"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="5867400" cy="4785360"/></a:xfrm>` +
      `</wpg:grpSpPr>` +
      custWsp({
        offX: 1353311,
        offY: 0,
        cx: 1899285,
        cy: 1179830,
        fill: '<a:solidFill><a:srgbClr val="1A295D"/></a:solidFill>',
      }) +
      `</wpg:wgp>`
    const para = anchorParagraph(group, GROUP_URI)
      .replace(
        '<wp:positionH relativeFrom="column"><wp:posOffset>100000</wp:posOffset></wp:positionH>',
        '<wp:positionH relativeFrom="page"><wp:posOffset>1213103</wp:posOffset></wp:positionH>',
      )
      .replace(
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>200000</wp:posOffset></wp:positionV>',
        '<wp:positionV relativeFrom="page"><wp:posOffset>3465576</wp:posOffset></wp:positionV>',
      )
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box).toBeTruthy()
    // page coords lose the 1440-twip margins: X exact, Y assumes the anchor
    // paragraph at body top (exact for a first-block anchor)
    expect(box!.pagePinned).toBeUndefined()
    expect(box!.offsetXEmu).toBe(1213103 - 1440 * 635 + 1353311)
    expect(box!.offsetYEmu).toBe(3465576 - 1440 * 635)
  })

  it('a later-block page-anchored group pins to the page box (raw page coordinates)', async () => {
    const group =
      `<wpg:wgp ${WPG_NS} ${WPS_NS}><wpg:cNvGrpSpPr/><wpg:grpSpPr>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="5867400" cy="4785360"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="5867400" cy="4785360"/></a:xfrm>` +
      `</wpg:grpSpPr>` +
      custWsp({
        offX: 1353311,
        offY: 0,
        cx: 1899285,
        cy: 1179830,
        fill: '<a:solidFill><a:srgbClr val="1A295D"/></a:solidFill>',
      }) +
      `</wpg:wgp>`
    const para =
      '<w:p><w:r><w:t>heading above</w:t></w:r></w:p>' +
      anchorParagraph(group, GROUP_URI)
        .replace(
          '<wp:positionH relativeFrom="column"><wp:posOffset>100000</wp:posOffset></wp:positionH>',
          '<wp:positionH relativeFrom="page"><wp:posOffset>1213103</wp:posOffset></wp:positionH>',
        )
        .replace(
          '<wp:positionV relativeFrom="paragraph"><wp:posOffset>200000</wp:posOffset></wp:positionV>',
          '<wp:positionV relativeFrom="page"><wp:posOffset>3465576</wp:posOffset></wp:positionV>',
        )
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const box = doc.blocks[1].textboxes?.[0]
    expect(box).toBeTruthy()
    expect(box!.pagePinned).toBe(true)
    expect(box!.floating).toBe(true)
    expect(box!.offsetXEmu).toBe(1213103 + 1353311)
    expect(box!.offsetYEmu).toBe(3465576)
  })
})

describe('grouped and sibling pictures', () => {
  const GROUP_WITH_PIC =
    `<wpg:wgp ${WPG_NS} ${WPS_NS}><wpg:cNvGrpSpPr/><wpg:grpSpPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm>` +
    `</wpg:grpSpPr>` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="Image 2"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="95250" y="190500"/><a:ext cx="476250" cy="238125"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    custWsp({
      offX: 0,
      offY: 0,
      cx: 476250,
      cy: 476250,
      fill: '<a:solidFill><a:srgbClr val="C8A24B"/></a:solidFill>',
    }) +
    `</wpg:wgp>`

  it('a pic:pic inside a wpg group becomes a photo box at the mapped offset', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: anchorParagraph(GROUP_WITH_PIC, GROUP_URI), withImage: true }),
    )
    const boxes = doc.blocks[0].textboxes
    expect(boxes?.length).toBe(2)
    const [photo, shape] = boxes!
    expect(photo.fillImageDataUrl).toContain('data:image/png')
    expect(photo.readOnly).toBe(true)
    expect(photo.floating).toBe(true)
    expect(photo.widthPx).toBe(50)
    expect(photo.heightPx).toBe(25)
    expect(photo.offsetXEmu).toBe(100000 + 95250)
    expect(photo.offsetYEmu).toBe(200000 + 190500)
    // document order preserved: the shape paints over the photo
    expect(shape.fill).toBe('C8A24B')
  })

  it('a rotated grouped pic:pic carries its rotation onto the photo box', async () => {
    const rotated = GROUP_WITH_PIC.replace('<pic:spPr><a:xfrm>', '<pic:spPr><a:xfrm rot="5400000">')
    const doc = await parseDocx(
      await buildDocx({ bodyXml: anchorParagraph(rotated, GROUP_URI), withImage: true }),
    )
    expect(doc.blocks[0].textboxes?.[0]?.rotDeg).toBe(90)
  })

  it('an inline picture sharing its paragraph with a text-bearing textbox becomes a photo box', async () => {
    const textbox = anchorParagraph(
      `<wps:wsp ${WPS_NS}><wps:cNvSpPr/><wps:spPr>` +
        `<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `</wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>card</w:t></w:r></w:p>` +
        `</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`,
      SHAPE_URI,
    ).replace(/<\/w:p>$/, '')
    const inlinePic =
      `<w:r><w:drawing><wp:inline><wp:extent cx="476250" cy="238125"/>` +
      `<wp:docPr id="8" name="Image 8"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="476250" cy="238125"/></a:xfrm></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: textbox + inlinePic, withImage: true }))
    // the inline pic flows as a stray run image: a non-floating photo box
    // would knock the block out of the zero-height floating overlay
    const boxes = doc.blocks[0].textboxes
    expect(boxes?.length).toBe(1)
    expect(boxes![0].paras[0]?.runs[0]?.text).toBe('card')
    const img = doc.blocks[0].strayRuns?.find((r) => r.image)?.image
    expect(img?.dataUrl).toContain('data:image/png')
    expect(img?.widthPx).toBe(50)
    expect(img?.heightPx).toBe(25)
  })

  it('an anchored sibling picture becomes a floating photo box carrying its rotation', async () => {
    const textbox = anchorParagraph(
      `<wps:wsp ${WPS_NS}><wps:cNvSpPr/><wps:spPr>` +
        `<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `</wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>card</w:t></w:r></w:p>` +
        `</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`,
      SHAPE_URI,
    ).replace(/<\/w:p>$/, '')
    const anchoredPic = anchorParagraph(
      `<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>` +
        `<pic:spPr><a:xfrm rot="-2700000"><a:off x="0" y="0"/>` +
        `<a:ext cx="476250" cy="238125"/></a:xfrm></pic:spPr></pic:pic>`,
      'http://schemas.openxmlformats.org/drawingml/2006/picture',
    ).replace('<w:p>', '')
    const doc = await parseDocx(
      await buildDocx({ bodyXml: textbox + anchoredPic, withImage: true }),
    )
    const boxes = doc.blocks[0].textboxes
    expect(boxes?.length).toBe(2)
    const photo = boxes![1]
    expect(photo.fillImageDataUrl).toContain('data:image/png')
    expect(photo.floating).toBe(true)
    expect(photo.rotDeg).toBe(-45)
    expect(doc.blocks[0].strayRuns).toBeUndefined()
  })

  it('a drawing inside mc:AlternateContent/mc:Choice still yields a run image', async () => {
    const body =
      `<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r>` +
      `<w:r><mc:AlternateContent ${MC_NS}><mc:Choice Requires="wps">` +
      `<w:drawing><wp:inline><wp:extent cx="476250" cy="238125"/>` +
      `<wp:docPr id="9" name="Image 9"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>` +
      `</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body, withImage: true }))
    const block = doc.blocks[0]
    expect(block.type).toBe('paragraph')
    const imageRun = block.runs?.find((r) => r.image)
    expect(imageRun?.image?.dataUrl).toContain('data:image/png')
    expect(imageRun?.image?.widthPx).toBe(50)
  })
})

describe('zero-height horizontal lines', () => {
  it('a cy="0" anchored line renders as a decorative rule, not a chip', async () => {
    const body =
      `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="page"><wp:posOffset>735965</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>95250</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="6724015" cy="0"/><wp:wrapNone/>` +
      `<wp:docPr id="2" name="Line 26"/>` +
      `<a:graphic><a:graphicData uri="${SHAPE_URI}">` +
      `<wps:wsp ${WPS_NS}><wps:cNvCnPr/><wps:spPr>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="6724015" cy="0"/></a:xfrm>` +
      `<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/>` +
      `<a:ln w="9144"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:round/><a:headEnd/><a:tailEnd/></a:ln>` +
      `</wps:spPr><wps:bodyPr/></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const block = doc.blocks[0]
    expect(block.decorative).toBe(true)
    expect(block.ruleColorHex).toBe('000000')
    expect(block.ruleThicknessPx).toBe(1)
    expect(block.ruleWidthPx).toBe(706)
  })

  it('a zero-extent drawing does not become a rule', async () => {
    const body =
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="0" cy="0"/>` +
      `<a:graphic><a:graphicData uri="${SHAPE_URI}"><wps:wsp ${WPS_NS}><wps:spPr/></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(doc.blocks[0].decorative).toBeFalsy()
  })
})

describe('text-empty filled textboxes', () => {
  it('a white full-page box with empty paragraphs keeps its extent as a fill-only box', async () => {
    const body = anchorParagraph(
      `<wps:wsp ${WPS_NS}><wps:cNvSpPr txBox="1"/><wps:spPr>` +
        `<a:xfrm><a:off x="0" y="0"/><a:ext cx="7078980" cy="9629775"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>` +
        `<a:ln w="9525"><a:noFill/></a:ln>` +
        `</wps:spPr><wps:txbx><w:txbxContent><w:p/><w:p/></w:txbxContent></wps:txbx>` +
        `<wps:bodyPr><a:noAutofit/></wps:bodyPr></wps:wsp>`,
      SHAPE_URI,
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box).toBeTruthy()
    expect(box!.fill).toBe('FFFFFF')
    // display keeps the pure-ink extent, but the w:txbxContent stays
    // addressable so typing into the empty box still commits
    expect(box!.readOnly).toBeUndefined()
    expect(box!.txbxIndex).toBe(0)
    expect(box!.paras).toEqual([])
    expect(box!.heightPx).toBe(Math.round(9629775 / 9525))
  })

  it('a text-empty box without visible ink still renders nothing', async () => {
    const body = anchorParagraph(
      `<wps:wsp ${WPS_NS}><wps:cNvSpPr txBox="1"/><wps:spPr>` +
        `<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
        `</wps:spPr><wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>` +
        `<wps:bodyPr/></wps:wsp>`,
      SHAPE_URI,
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(doc.blocks[0].textboxes).toBeUndefined()
  })
})

describe('centered wrapSquare pictures', () => {
  it('positionH align=center + wrapSquare maps to the centered topBottom slot', async () => {
    const body =
      `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>133350</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="3293110" cy="594360"/>` +
      `<wp:wrapSquare wrapText="bothSides"/>` +
      `<wp:docPr id="1" name="Image 1"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>` +
      `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body, withImage: true }))
    const block = doc.blocks[0]
    expect(block.type).toBe('image')
    expect(block.imageWrap).toBe('topBottom')
    expect(block.imageOffsetXEmu).toBeUndefined()
    expect(block.imageOffsetYEmu).toBe(133350)
  })

  it('a left-offset wrapSquare picture keeps the square-left float', async () => {
    const body =
      `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="3293110" cy="594360"/>` +
      `<wp:wrapSquare wrapText="bothSides"/>` +
      `<wp:docPr id="1" name="Image 1"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>` +
      `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body, withImage: true }))
    expect(doc.blocks[0].imageWrap).toBe('square-left')
  })
})
