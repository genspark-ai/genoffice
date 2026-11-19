import { describe, expect, it } from 'vitest'
import {
  groupAnchoredDrawings,
  maxDrawingId,
  parseDocx,
  regroupDrawings,
  saveDocx,
  splitGroupParagraph,
  ungroupAnchoredDrawing,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const EMU = 9525
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

const anchorHead = (id: number, x: number, y: number, w: number, h: number): string =>
  `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" ` +
  `behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="column"><wp:posOffset>${x * EMU}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${y * EMU}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${w * EMU}" cy="${h * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:wrapNone/><wp:docPr id="${id}" name="Object ${id}"/><wp:cNvGraphicFramePr/>`

/** a rectangle in front of the text, with `text` inside, px */
const shape = (id: number, x: number, y: number, w: number, h: number, text = ''): string =>
  `<w:drawing>${anchorHead(id, x, y, w, h)}<a:graphic ${A}>` +
  `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * EMU}" cy="${h * EMU}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>` +
  (text
    ? `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>`
    : '') +
  `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`

const picture = (id: number, x: number, y: number, w: number, h: number): string =>
  `<w:drawing>${anchorHead(id, x, y, w, h)}<a:graphic ${A}>` +
  `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * EMU}" cy="${h * EMU}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:anchor></w:drawing>`

const para = (inner: string): string => `<w:p>${inner}</w:p>`
const run = (inner: string): string => `<w:r>${inner}</w:r>`
const text = (t: string): string => run(`<w:t xml:space="preserve">${t}</w:t>`)

/** three drawings in three paragraphs: a shape with text, a shape beside text, a picture among text */
const BLOCKS = [
  para(text('Lead')),
  para(run(shape(1, 40, 30, 120, 60, 'Inside'))),
  para(
    `<mc:AlternateContent><mc:Choice Requires="wps">${run(shape(2, 200, 10, 80, 90))}</mc:Choice>` +
      `<mc:Fallback>${run('<w:pict><v:rect/></w:pict>')}</mc:Fallback></mc:AlternateContent>`,
  ),
  para(text('Before ') + run(picture(3, 300, 5, 50, 50)) + text(' after')),
]

/** the three as one group: boxes inside it, EMU */
const GROUP = {
  members: [
    { block: 1, drawing: 0, xEmu: 0, yEmu: 0, widthEmu: 120 * EMU, heightEmu: 60 * EMU },
    {
      block: 2,
      drawing: 0,
      xEmu: 160 * EMU,
      yEmu: 40 * EMU,
      widthEmu: 80 * EMU,
      heightEmu: 90 * EMU,
    },
    {
      block: 3,
      drawing: 0,
      xEmu: 260 * EMU,
      yEmu: 100 * EMU,
      widthEmu: 50 * EMU,
      heightEmu: 50 * EMU,
    },
  ],
}

const offsets = (xml: string, axis: 'H' | 'V'): number[] =>
  Array.from(
    xml.matchAll(new RegExp(`<wp:position${axis}[^>]*><wp:posOffset>(-?\\d+)</wp:posOffset>`, 'g')),
  ).map((m) => Number(m[1]) / EMU)

describe('grouping anchored drawings', () => {
  it('moves the members into one group anchor and leaves their paragraphs behind', () => {
    const out = groupAnchoredDrawings(BLOCKS, GROUP, 100)!
    expect(out).toHaveLength(4)
    expect(out[0]).toBe(BLOCKS[0])
    // the group takes the first member's place and spans all three
    expect(out[1].match(/<wp:anchor /g)).toHaveLength(1)
    expect(out[1]).toContain(`<wp:extent cx="${310 * EMU}" cy="${150 * EMU}"/>`)
    expect(out[1]).toContain('<wpg:wgp ')
    expect(out[1]).toContain(`<a:chExt cx="${310 * EMU}" cy="${150 * EMU}"/>`)
    expect(out[1].match(/<wps:wsp>/g)).toHaveLength(2)
    expect(out[1].match(/<pic:pic /g)).toHaveLength(1)
    // each child sits at its box in the group and keeps its size and what it holds
    expect(out[1]).toContain(
      `<a:off x="${160 * EMU}" y="${40 * EMU}"/><a:ext cx="${80 * EMU}" cy="${90 * EMU}"/>`,
    )
    expect(out[1]).toContain(
      `<a:off x="${260 * EMU}" y="${100 * EMU}"/><a:ext cx="${50 * EMU}" cy="${50 * EMU}"/>`,
    )
    expect(out[1]).toContain('<w:t>Inside</w:t>')
    expect(out[1]).toContain('r:embed="rId10"')
    // a grouped shape carries its own name
    expect(out[1]).toContain('<wps:wsp><wps:cNvPr id="100" name="Object 1"/><wps:cNvSpPr/>')
    // the others are gone from their paragraphs, fallback and empty run included; the text stays
    expect(out[2]).toBe('<w:p></w:p>')
    expect(out[3]).toBe(para(text('Before ') + text(' after')))
  })

  it("places the group where its first member was, less that member's place inside it", () => {
    const shifted = {
      members: GROUP.members.map((m) => ({
        ...m,
        xEmu: m.xEmu + 15 * EMU,
        yEmu: m.yEmu + 7 * EMU,
      })),
    }
    const out = groupAnchoredDrawings(BLOCKS, shifted, 100)!
    expect(offsets(out[1], 'H')).toEqual([25])
    expect(offsets(out[1], 'V')).toEqual([23])
  })

  it('refuses a member that is missing or is not a shape or a picture', () => {
    const missing = { members: [GROUP.members[0], { ...GROUP.members[1], drawing: 1 }] }
    expect(groupAnchoredDrawings(BLOCKS, missing, 100)).toBeNull()
    const grouped = groupAnchoredDrawings(BLOCKS, GROUP, 100)!
    const again = { members: [GROUP.members[0], { ...GROUP.members[1], block: 0 }] }
    expect(groupAnchoredDrawings(grouped, again, 200)).toBeNull()
    expect(groupAnchoredDrawings(BLOCKS, { members: [GROUP.members[0]] }, 100)).toBeNull()
  })

  it('keeps the group inside a run when the run of the first member sat inside its fallback wrapper', () => {
    // block 2's wrapper holds the run; as the first member its place must stay run content
    const lead = { members: [GROUP.members[1], GROUP.members[0], GROUP.members[2]] }
    const out = groupAnchoredDrawings(BLOCKS, lead, 100)!
    expect(out[2]).toMatch(/^<w:p><w:r><w:drawing>.*<\/w:drawing><\/w:r><\/w:p>$/)
    const split = ungroupAnchoredDrawing(out[2], 0, 500)!
    expect(split).toMatch(/^<w:p><w:r>(?:<w:drawing>.*?<\/w:drawing>){3}<\/w:r><\/w:p>$/)
  })

  it('keeps the frame locks Word writes ahead of the graphic', () => {
    const locks =
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
    const locked = BLOCKS.map((xml) => xml.replace('<wp:cNvGraphicFramePr/>', locks))
    const out = groupAnchoredDrawings(locked, GROUP, 100)!
    expect(out[1]).toContain(`${locks}<a:graphic `)
    expect(out[1].match(/<a:graphicFrameLocks /g)).toHaveLength(1)
    const split = ungroupAnchoredDrawing(out[1], 0, 500)!
    expect(split.match(new RegExp(`${locks}<a:graphic `, 'g'))).toHaveLength(3)
  })

  it('takes two drawings of one paragraph', () => {
    const blocks = [para(run(shape(1, 0, 0, 40, 40)) + text('mid') + run(shape(2, 100, 0, 40, 40)))]
    const group = {
      members: [
        { block: 0, drawing: 0, xEmu: 0, yEmu: 0, widthEmu: 40 * EMU, heightEmu: 40 * EMU },
        { block: 0, drawing: 1, xEmu: 100 * EMU, yEmu: 0, widthEmu: 40 * EMU, heightEmu: 40 * EMU },
      ],
    }
    const [out] = groupAnchoredDrawings(blocks, group, 100)!
    expect(out.match(/<wp:anchor /g)).toHaveLength(1)
    expect(out.match(/<wps:wsp>/g)).toHaveLength(2)
    expect(out.endsWith(`${text('mid')}</w:p>`)).toBe(true)
  })

  it('leaves the formatted text ahead of a member whose own run has properties', () => {
    const bold = '<w:r><w:rPr><w:b/></w:rPr><w:t>Kept</w:t></w:r>'
    const noProof = (drawing: string): string => `<w:r><w:rPr><w:noProof/></w:rPr>${drawing}</w:r>`
    const member = (block: number, drawing: number, x: number) => ({
      block,
      drawing,
      xEmu: x * EMU,
      yEmu: 0,
      widthEmu: 40 * EMU,
      heightEmu: 40 * EMU,
    })
    const apart = [
      para(run(shape(1, 0, 0, 40, 40))),
      para(bold + noProof(shape(2, 100, 0, 40, 40))),
    ]
    const out = groupAnchoredDrawings(
      apart,
      { members: [member(0, 0, 0), member(1, 0, 100)] },
      100,
    )!
    expect(out[1]).toBe(para(bold))

    const together = [
      para(bold + noProof(shape(1, 0, 0, 40, 40)) + bold + noProof(shape(2, 100, 0, 40, 40))),
    ]
    const [one] = groupAnchoredDrawings(
      together,
      { members: [member(0, 0, 0), member(0, 1, 100)] },
      100,
    )!
    expect(one.match(/<wp:anchor /g)).toHaveLength(1)
    expect(one.match(/<wps:wsp>/g)).toHaveLength(2)
    expect(one.match(/<w:t>Kept<\/w:t>/g)).toHaveLength(2)
    expect(one.endsWith(`${bold}</w:p>`)).toBe(true)
  })

  it('reads back as one drawing whose boxes are where the members were', async () => {
    const out = groupAnchoredDrawings(BLOCKS, GROUP, 100)!
    const parsed = await parseDocx(await buildDocx({ bodyXml: out.join(''), withImage: true }))
    const block = parsed.blocks.find((b) => b.textboxes?.length)!
    const boxes = block.textboxes!
    expect(boxes).toHaveLength(3)
    expect(boxes.every((b) => b.grouped && b.floating)).toBe(true)
    // the group anchor is at (40, 30): each box is its offset in the group from there
    expect(
      boxes.map((b) => [b.offsetXEmu! / EMU, b.offsetYEmu! / EMU, b.widthPx, b.heightPx]),
    ).toEqual([
      [40, 30, 120, 60],
      [200, 70, 80, 90],
      [300, 130, 50, 50],
    ])
    expect(boxes[0].paras[0].runs[0].text).toBe('Inside')
    expect(boxes[2].picture).toBe(true)
  })

  it('is what the save writes for a group of body blocks', async () => {
    const source = await buildDocx({ bodyXml: BLOCKS.join(''), withImage: true })
    const parsed = await parseDocx(source)
    const blocks = parsed.blocks.filter((b) => !b.hidden && b.docxIndex !== null)
    const saved = await saveDocx(
      parsed,
      blocks.map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
      { groups: [GROUP] },
    )
    const reread = await parseDocx(saved)
    const withBoxes = reread.blocks.filter((b) => b.textboxes?.length)
    expect(withBoxes).toHaveLength(1)
    expect(withBoxes[0].textboxes).toHaveLength(3)
  })
})

describe('several groups in one save', () => {
  const small = (block: number, drawing: number, x: number, size: number) => ({
    block,
    drawing,
    xEmu: x * EMU,
    yEmu: 0,
    widthEmu: size * EMU,
    heightEmu: size * EMU,
  })

  it('merges two groups that each have a member in the same paragraph', () => {
    // four shapes: the first and last alone in their paragraphs, the middle two sharing one
    const blocks = [
      para(run(shape(1, 0, 0, 40, 40))),
      para(run(shape(2, 100, 0, 40, 40)) + text('mid') + run(shape(3, 200, 0, 40, 40))),
      para(run(shape(4, 300, 0, 40, 40))),
    ]
    // the first group takes a drawing out of the shared paragraph: the other is still its second
    const out = regroupDrawings(
      blocks,
      [
        { members: [small(0, 0, 0, 40), small(1, 0, 100, 40)] },
        { members: [small(1, 1, 0, 40), small(2, 0, 100, 40)] },
      ],
      [],
      100,
    )
    expect(out.map((xml) => xml.match(/<wpg:wgp /g)?.length ?? 0)).toEqual([1, 1, 0])
    expect(out.map((xml) => xml.match(/<wps:wsp>/g)?.length ?? 0)).toEqual([2, 2, 0])
    expect(out[1]).toContain(text('mid'))
    expect(out[2]).toBe('<w:p></w:p>')
  })

  it('leaves out a group whose member another group already took', () => {
    const out = regroupDrawings(
      BLOCKS,
      [GROUP, { members: [GROUP.members[1], { ...GROUP.members[0], block: 3 }] }],
      [],
      100,
    )
    expect(out.join('').match(/<wpg:wgp /g)).toHaveLength(1)
  })
})

describe('ungrouping an anchored drawing', () => {
  it("gives every child its own anchor at its place, in the group's run", () => {
    const grouped = groupAnchoredDrawings(BLOCKS, GROUP, 100)!
    const out = ungroupAnchoredDrawing(grouped[1], 0, maxDrawingId(grouped.join('')) + 1)!
    expect(out.match(/<wp:anchor /g)).toHaveLength(3)
    expect(out).not.toContain('wpg:wgp')
    expect(offsets(out, 'H')).toEqual([40, 200, 300])
    expect(offsets(out, 'V')).toEqual([30, 70, 130])
    expect(
      Array.from(out.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)).map((m) => [
        Number(m[1]) / EMU,
        Number(m[2]) / EMU,
      ]),
    ).toEqual([
      [120, 60],
      [80, 90],
      [50, 50],
    ])
    // every anchor has an id of its own
    const ids = Array.from(out.matchAll(/<wp:docPr id="(\d+)"/g)).map((m) => m[1])
    expect(new Set(ids).size).toBe(3)
    expect(out).toContain('<w:t>Inside</w:t>')
    expect(out.match(/<a:off x="0" y="0"\/>/g)).toHaveLength(3)
  })

  it('scales the children of a resized group', () => {
    const grouped = groupAnchoredDrawings(BLOCKS, GROUP, 100)![1]
    // the group drawn at half its child space
    const half = grouped.replace(
      `<a:ext cx="${310 * EMU}" cy="${150 * EMU}"/><a:chOff`,
      `<a:ext cx="${155 * EMU}" cy="${75 * EMU}"/><a:chOff`,
    )
    const out = ungroupAnchoredDrawing(half, 0, 500)!
    expect(offsets(out, 'H')).toEqual([40, 120, 170])
    expect(out).toContain(`<wp:extent cx="${60 * EMU}" cy="${30 * EMU}"/>`)
  })

  it('reads back as three separate floating boxes', async () => {
    const grouped = groupAnchoredDrawings(BLOCKS, GROUP, 100)!
    grouped[1] = ungroupAnchoredDrawing(grouped[1], 0, 500)!
    const parsed = await parseDocx(await buildDocx({ bodyXml: grouped.join(''), withImage: true }))
    const boxes = parsed.blocks.find((b) => b.textboxes?.length)!.textboxes!
    expect(boxes).toHaveLength(3)
    expect(boxes.some((b) => b.grouped)).toBe(false)
  })

  it('refuses what is not a group, and a group placed by alignment', () => {
    expect(ungroupAnchoredDrawing(BLOCKS[1], 0, 100)).toBeNull()
    const grouped = groupAnchoredDrawings(BLOCKS, GROUP, 100)![1]
    expect(ungroupAnchoredDrawing(grouped, 1, 100)).toBeNull()
    const aligned = grouped.replace(
      /<wp:positionH[^>]*>.*?<\/wp:positionH>/,
      '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>',
    )
    expect(ungroupAnchoredDrawing(aligned, 0, 100)).toBeNull()
  })
})

/** a child shape of a group, px in the group's child space */
const child = (
  id: number,
  x: number,
  y: number,
  w: number,
  h: number,
  attrs = '',
  txt = '',
): string =>
  `<wps:wsp><wps:cNvPr id="${id}" name="Child ${id}"/><wps:cNvSpPr/><wps:spPr><a:xfrm${attrs}>` +
  `<a:off x="${x * EMU}" y="${y * EMU}"/><a:ext cx="${w * EMU}" cy="${h * EMU}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>` +
  (txt
    ? `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${txt}</w:t></w:r></w:p></w:txbxContent></wps:txbx>`
    : '') +
  `<wps:bodyPr/></wps:wsp>`

const innerGroup = (
  id: number,
  box: [number, number, number, number],
  attrs: string,
  kids: string,
): string =>
  `<wpg:grpSp><wpg:cNvPr id="${id}" name="Inner ${id}"/><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm${attrs}>` +
  `<a:off x="${box[0] * EMU}" y="${box[1] * EMU}"/><a:ext cx="${box[2] * EMU}" cy="${box[3] * EMU}"/>` +
  `<a:chOff x="${box[0] * EMU}" y="${box[1] * EMU}"/><a:chExt cx="${box[2] * EMU}" cy="${box[3] * EMU}"/>` +
  `</a:xfrm></wpg:grpSpPr>${kids}</wpg:grpSp>`

/** a paragraph holding one group at (x, y), w by h, its child space the same size */
const groupParagraph = (
  x: number,
  y: number,
  w: number,
  h: number,
  attrs: string,
  kids: string,
  after = '',
): string =>
  para(
    run(
      `<w:drawing>${anchorHead(50, x, y, w, h)}<a:graphic ${A}>` +
        `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">` +
        `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm${attrs}><a:off x="0" y="0"/>` +
        `<a:ext cx="${w * EMU}" cy="${h * EMU}"/><a:chOff x="0" y="0"/><a:chExt cx="${w * EMU}" cy="${h * EMU}"/>` +
        `</a:xfrm></wpg:grpSpPr>${kids}</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing>`,
    ) + after,
  )

const KIDS =
  child(1, 0, 0, 100, 60) + child(2, 140, 20, 80, 80) + child(3, 60, 110, 120, 40, ' rot="1800000"')

/** each box as [left, top, width, height, turn], px and degrees */
const placedBoxes = async (bodyXml: string): Promise<number[][]> => {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  return parsed.blocks
    .flatMap((b) => b.textboxes ?? [])
    .map((b) => [
      Math.round(b.offsetXEmu! / EMU),
      Math.round(b.offsetYEmu! / EMU),
      b.widthPx!,
      b.heightPx!,
      b.rotDeg ?? 0,
    ])
}

const xfrmAttrs = (xml: string): string[] =>
  Array.from(xml.matchAll(/<wps:spPr><a:xfrm([^>]*)>/g)).map((m) => m[1].trim())

describe('ungrouping a turned, mirrored or nested group', () => {
  it('turns the children about the middle of a group turned a quarter', () => {
    // a 220 x 160 group at (60, 20): its middle is (170, 100) on the page
    const out = ungroupAnchoredDrawing(
      groupParagraph(60, 20, 220, 160, ' rot="5400000"', KIDS),
      0,
      500,
    )!
    expect(out).not.toContain('wpg:wgp')
    // the first child's middle (50, 30) is (-60, -50) from the group's: a quarter turn puts it at (50, -60)
    expect(offsets(out, 'H')).toEqual([170, 150, 60])
    expect(offsets(out, 'V')).toEqual([10, 130, 90])
    expect(xfrmAttrs(out)).toEqual(['rot="5400000"', 'rot="5400000"', 'rot="7200000"'])
    // sizes are the children's own
    expect(out).toContain(`<wp:extent cx="${100 * EMU}" cy="${60 * EMU}"/>`)
  })

  it('mirrors the places and the turn of the children of a flipped group', () => {
    const out = ungroupAnchoredDrawing(
      groupParagraph(60, 20, 220, 160, ' flipH="1"', KIDS),
      0,
      500,
    )!
    expect(offsets(out, 'H')).toEqual([180, 60, 100])
    expect(offsets(out, 'V')).toEqual([20, 40, 130])
    expect(xfrmAttrs(out)).toEqual(['flipH="1"', 'flipH="1"', 'rot="19800000" flipH="1"'])
  })

  it('a group flipped both ways turns its children half way round without reversing them', () => {
    const out = ungroupAnchoredDrawing(
      groupParagraph(60, 20, 220, 160, ' flipH="1" flipV="1"', KIDS),
      0,
      500,
    )!
    expect(offsets(out, 'V')).toEqual([120, 80, 30])
    expect(xfrmAttrs(out)[2]).toBe('rot="1800000" flipH="1" flipV="1"')
  })

  it('takes one level off a nested group: the inner group is a group of its own', () => {
    const nested = groupParagraph(
      60,
      20,
      320,
      160,
      '',
      child(1, 0, 0, 100, 60) +
        innerGroup(
          9,
          [150, 30, 170, 130],
          '',
          child(2, 150, 30, 80, 80) + child(3, 240, 110, 80, 50),
        ),
    )
    const once = ungroupAnchoredDrawing(nested, 0, 500)!
    expect(once.match(/<wp:anchor /g)).toHaveLength(2)
    expect(once.match(/<wpg:wgp[ >]/g)).toHaveLength(1)
    expect(once).not.toContain('wpg:grpSp>')
    expect(once).not.toContain('<wpg:cNvPr')
    expect(offsets(once, 'H')).toEqual([60, 210])
    expect(offsets(once, 'V')).toEqual([20, 50])
    // the inner group keeps its child space and is named by its anchor
    expect(once).toContain(
      `<a:off x="0" y="0"/><a:ext cx="${170 * EMU}" cy="${130 * EMU}"/>` +
        `<a:chOff x="${150 * EMU}" y="${30 * EMU}"/><a:chExt cx="${170 * EMU}" cy="${130 * EMU}"/>`,
    )
    expect(once).toContain('name="Inner 9"')
    const twice = ungroupAnchoredDrawing(once, 1, 600)!
    expect(twice.match(/<wp:anchor /g)).toHaveLength(3)
    expect(twice).not.toContain('wpg:')
    expect(offsets(twice, 'H')).toEqual([60, 210, 300])
    expect(offsets(twice, 'V')).toEqual([20, 50, 130])
  })

  it('hands the turn of the outer group to an inner group', () => {
    const nested = groupParagraph(
      0,
      0,
      200,
      200,
      ' rot="5400000"',
      innerGroup(9, [100, 0, 100, 50], ' rot="2700000"', child(2, 100, 0, 100, 50)),
    )
    const once = ungroupAnchoredDrawing(nested, 0, 500)!
    expect(once).toContain('<wpg:grpSpPr><a:xfrm rot="8100000">')
    expect(offsets(once, 'H')).toEqual([125])
    expect(offsets(once, 'V')).toEqual([125])
  })

  for (const [name, attrs] of [
    ['turned', ' rot="2700000"'],
    ['mirrored', ' flipV="1"'],
    ['turned and mirrored', ' rot="16200000" flipH="1"'],
  ]) {
    it(`a ${name} group reads at the places its children get when it is taken apart`, async () => {
      const grouped = groupParagraph(60, 20, 220, 160, attrs, KIDS)
      const before = await placedBoxes(grouped)
      expect(before).toHaveLength(3)
      expect(await placedBoxes(ungroupAnchoredDrawing(grouped, 0, 500)!)).toEqual(before)
    })
  }

  it('a nested turned group reads at the same places after each level is taken off', async () => {
    const nested = groupParagraph(
      60,
      20,
      320,
      160,
      ' rot="1800000"',
      child(1, 0, 0, 100, 60) +
        innerGroup(
          9,
          [150, 30, 170, 130],
          ' rot="3600000"',
          child(2, 150, 30, 80, 80) + child(3, 240, 110, 80, 50, ' rot="900000"'),
        ),
    )
    const before = await placedBoxes(nested)
    const once = ungroupAnchoredDrawing(nested, 0, 500)!
    expect(await placedBoxes(once)).toEqual(before)
    expect(await placedBoxes(ungroupAnchoredDrawing(once, 1, 600)!)).toEqual(before)
  })

  it('tells each box the way down to it from the group', async () => {
    const nested = groupParagraph(
      60,
      20,
      320,
      160,
      '',
      child(1, 0, 0, 100, 60) +
        innerGroup(
          9,
          [150, 30, 170, 130],
          '',
          child(2, 150, 30, 80, 80) + child(3, 240, 110, 80, 50),
        ),
    )
    const parsed = await parseDocx(await buildDocx({ bodyXml: nested }))
    const boxes = parsed.blocks.find((b) => b.textboxes?.length)!.textboxes!
    expect(boxes.map((b) => b.groupPath)).toEqual([[0], [1, 0], [1, 1]])
  })
})

describe('a group paragraph taken apart into paragraphs', () => {
  const grouped = groupParagraph(
    60,
    20,
    220,
    160,
    '',
    child(1, 0, 0, 100, 60, '', 'One') +
      child(2, 140, 20, 80, 80) +
      child(3, 60, 110, 120, 40, '', 'Three'),
    text('beside'),
  )

  it('gives every child a paragraph, the first one the paragraph the group was in', () => {
    const split = splitGroupParagraph(grouped, 0, 500)!
    expect(split.paragraphs).toHaveLength(3)
    expect(split.paragraphs[2]).toMatch(/^<w:p><w:r><w:drawing>.*<\/w:drawing><\/w:r><\/w:p>$/)
    expect(split.paragraphs[0]).toContain('<w:t>One</w:t>')
    expect(split.paragraphs[2]).toContain('<w:t>Three</w:t>')
    expect(split.paragraphs[0].endsWith(`${text('beside')}</w:p>`)).toBe(true)
    expect(split.paragraphs.map((xml) => xml.match(/<wp:anchor /g)?.length)).toEqual([1, 1, 1])
    expect(
      split.members.map((m) => [m.kind, m.xEmu / EMU, m.yEmu / EMU, m.textBoxesBefore]),
    ).toEqual([
      ['shape', 60, 20, 0],
      ['shape', 200, 40, 1],
      ['shape', 120, 130, 1],
    ])
  })

  it('saves the carried paragraphs inside the paragraph before them: the file is the group ungrouped in place', async () => {
    const body = [para(text('Lead')), grouped, para(text('Tail'))]
    const split = splitGroupParagraph(grouped, 0, 500)!
    const out = regroupDrawings([body[0], ...split.paragraphs, body[2]], [], [2, 3], 600)
    expect(out[2]).toBe('')
    expect(out[3]).toBe('')
    // exactly what ungrouping inside the paragraph gives, but for the text now ahead of the later drawings
    expect(out[1].match(/<wp:anchor /g)).toHaveLength(3)
    expect(offsets(out[1], 'H')).toEqual([60, 200, 120])
    expect(out.join('').match(/<w:p>/g)).toHaveLength(body.join('').match(/<w:p>/g)!.length)
    expect(await placedBoxes(out.join(''))).toEqual(await placedBoxes(body.join('')))
  })

  it('leaves a carried paragraph alone when no paragraph comes before it', () => {
    const split = splitGroupParagraph(grouped, 0, 500)!
    const blocks = ['<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>', split.paragraphs[1]]
    expect(regroupDrawings(blocks, [], [1], 600)).toEqual(blocks)
    expect(regroupDrawings([split.paragraphs[1]], [], [0], 600)).toEqual([split.paragraphs[1]])
  })

  it('a group made of carried drawings ends up in the paragraph before them', () => {
    const split = splitGroupParagraph(grouped, 0, 500)!
    const member = (block: number, x: number, y: number, w: number, h: number) => ({
      block,
      drawing: 0,
      xEmu: x * EMU,
      yEmu: y * EMU,
      widthEmu: w * EMU,
      heightEmu: h * EMU,
    })
    const out = regroupDrawings(
      split.paragraphs,
      [{ members: [member(1, 0, 0, 80, 80), member(2, 0, 90, 120, 40)] }],
      [1, 2],
      600,
    )
    expect(out[1]).toBe('')
    expect(out[2]).toBe('')
    expect(out[0].match(/<wpg:wgp /g)).toHaveLength(1)
    expect(out[0].match(/<wp:anchor /g)).toHaveLength(2)
    expect(out[0].endsWith('</w:drawing></w:r></w:p>')).toBe(true)
  })

  it('is what the save writes for carried blocks', async () => {
    const source = await buildDocx({ bodyXml: para(text('Lead')) + grouped + para(text('Tail')) })
    const parsed = await parseDocx(source)
    const split = splitGroupParagraph(grouped, 0, 500)!
    const saved = await saveDocx(
      parsed,
      [
        { kind: 'original', docxIndex: 0 },
        ...split.paragraphs.map((xml) => ({ kind: 'xml' as const, xml })),
        { kind: 'original', docxIndex: 2 },
      ],
      {
        carried: [2, 3],
      },
    )
    const reread = await parseDocx(saved)
    const shown = reread.blocks.filter((b) => !b.hidden)
    expect(shown).toHaveLength(3)
    expect(shown[1].textboxes).toHaveLength(3)
    expect(shown[1].textboxes!.some((b) => b.grouped)).toBe(false)
  })
})
