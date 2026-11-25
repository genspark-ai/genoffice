import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import {
  patchRunPictureXml,
  runPictureAltOf,
  setRunPictureAlt,
  type RunPictureGeometry,
} from '../src/run-picture'
import type { Run } from '../src/types'
import { buildDocx, IMAGE_PARAGRAPH_XML } from './helpers/build-docx'

const SQUARE = '<wp:wrapSquare wrapText="bothSides"/>'
const TIGHT =
  '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/>' +
  '<wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/>' +
  '<wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'

function anchoredPicture(o: {
  posH: string
  posV: string
  relH?: string
  relV?: string
  wrap?: string
  cx?: number
  cy?: number
}): string {
  const cx = o.cx ?? 1905000
  const cy = o.cy ?? 1270000
  return (
    `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="${o.relH ?? 'column'}">${o.posH}</wp:positionH>` +
    `<wp:positionV relativeFrom="${o.relV ?? 'paragraph'}">${o.posV}</wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/>${o.wrap ?? SQUARE}<wp:docPr id="1" name="Picture 1"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing>`
  )
}

const posOffset = (emu: number) => `<wp:posOffset>${emu}</wp:posOffset>`

/** the run picture of a paragraph that shares its line with text */
async function runPicture(drawing: string): Promise<NonNullable<Run['image']>> {
  const source = await buildDocx({
    bodyXml: `<w:p><w:r><w:t>Text beside the picture.</w:t></w:r><w:r>${drawing}</w:r></w:p>`,
    withImage: true,
  })
  const parsed = await parseDocx(source)
  const runs = parsed.blocks.flatMap((block) => ('runs' in block ? (block.runs ?? []) : []))
  const image = runs.find((run) => run.image)?.image
  if (!image) throw new Error('no run picture parsed')
  return image
}

/** what the editor holds for a run picture before anyone edits it */
const geometryOf = (image: NonNullable<Run['image']>): RunPictureGeometry => ({
  widthPx: image.widthPx,
  heightPx: image.heightPx,
  wrap: image.wrap ?? null,
  offsetXEmu: image.offsetXEmu ?? null,
  offsetYEmu: image.offsetYEmu ?? null,
  relH: image.relH ?? null,
  relV: image.relV ?? null,
  lineCenterV: image.lineCenterV ?? false,
})

/** re-parse a patched fragment the way the importer reads it */
const reparse = (xml: string) => runPicture(xml)

describe('run-level picture geometry on save', () => {
  it('keeps an untouched floating picture byte-identical', async () => {
    for (const drawing of [
      anchoredPicture({ posH: posOffset(457200), posV: posOffset(95250) }),
      anchoredPicture({ posH: '<wp:align>right</wp:align>', posV: posOffset(0), wrap: TIGHT }),
      anchoredPicture({
        posH: posOffset(914400),
        posV: posOffset(1828800),
        relH: 'page',
        relV: 'page',
        wrap: '<wp:wrapNone/>',
      }),
    ]) {
      const image = await runPicture(drawing)
      expect(patchRunPictureXml(image.xml, geometryOf(image))).toBe(image.xml)
    }
  })

  it('writes a moved picture as numeric offsets on its own bases', async () => {
    const image = await runPicture(anchoredPicture({ posH: posOffset(457200), posV: posOffset(0) }))
    const moved = patchRunPictureXml(image.xml, {
      ...geometryOf(image),
      offsetXEmu: 1143000,
      offsetYEmu: 381000,
    })
    expect(moved).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>1143000</wp:posOffset></wp:positionH>',
    )
    expect(moved).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>381000</wp:posOffset></wp:positionV>',
    )
    const again = await reparse(moved)
    expect([again.offsetXEmu, again.offsetYEmu, again.wrap]).toEqual([
      1143000,
      381000,
      'square-left',
    ])
    // everything else of the drawing keeps its bytes
    expect(moved.replace(/<wp:position[HV][\s\S]*?<\/wp:position[HV]>/g, '')).toBe(
      image.xml.replace(/<wp:position[HV][\s\S]*?<\/wp:position[HV]>/g, ''),
    )
  })

  it('keeps a page-relative base when the picture moves on it', async () => {
    const image = await runPicture(
      anchoredPicture({
        posH: posOffset(914400),
        posV: posOffset(1828800),
        relH: 'page',
        relV: 'page',
        wrap: '<wp:wrapNone/>',
      }),
    )
    const moved = patchRunPictureXml(image.xml, { ...geometryOf(image), offsetXEmu: 1000000 })
    expect(moved).toContain(
      '<wp:positionH relativeFrom="page"><wp:posOffset>1000000</wp:posOffset></wp:positionH>',
    )
    // the untouched axis keeps its bytes
    expect(moved).toContain(
      '<wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>',
    )
  })

  it('turns an aligned axis into an offset once the picture is dragged', async () => {
    const image = await runPicture(
      anchoredPicture({ posH: '<wp:align>right</wp:align>', posV: '<wp:align>top</wp:align>' }),
    )
    expect(image.offsetXEmu).toBeUndefined()
    const moved = patchRunPictureXml(image.xml, {
      ...geometryOf(image),
      offsetXEmu: 2000000,
      offsetYEmu: 500000,
    })
    expect(moved).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>2000000</wp:posOffset></wp:positionH>',
    )
    expect(moved).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>500000</wp:posOffset></wp:positionV>',
    )
  })

  it('resizes a floating picture: wp:extent and the picture a:ext follow', async () => {
    const image = await runPicture(anchoredPicture({ posH: posOffset(0), posV: posOffset(0) }))
    const resized = patchRunPictureXml(image.xml, {
      ...geometryOf(image),
      widthPx: 300,
      heightPx: 200,
    })
    expect(resized).toContain('<wp:extent cx="2857500" cy="1905000"/>')
    expect(resized).toContain('<a:ext cx="2857500" cy="1905000"/>')
    const again = await reparse(resized)
    expect([again.widthPx, again.heightPx]).toEqual([300, 200])
  })

  it('switches the wrap kind in place, keeping the position', async () => {
    const image = await runPicture(anchoredPicture({ posH: posOffset(457200), posV: posOffset(0) }))
    const behind = patchRunPictureXml(image.xml, { ...geometryOf(image), wrap: 'behind' })
    expect(behind).toContain('behindDoc="1"')
    expect(behind).toContain('<wp:wrapNone/>')
    expect(behind).not.toContain('<wp:wrapSquare')
    expect(behind).toContain('<wp:posOffset>457200</wp:posOffset>')
    expect((await reparse(behind)).wrap).toBe('behind')
  })

  it('flips the side of a tight float without dropping its polygon', async () => {
    const image = await runPicture(
      anchoredPicture({ posH: '<wp:align>left</wp:align>', posV: posOffset(0), wrap: TIGHT }),
    )
    expect(image.wrap).toBe('tight-left')
    const flipped = patchRunPictureXml(image.xml, { ...geometryOf(image), wrap: 'tight-right' })
    expect(flipped).toContain('<wp:wrapPolygon')
    expect(flipped).toContain('<wp:align>right</wp:align>')
    expect((await reparse(flipped)).wrap).toBe('tight-right')
  })

  it('converts between in-line and floating', async () => {
    const floating = await runPicture(anchoredPicture({ posH: posOffset(0), posV: posOffset(0) }))
    const inline = patchRunPictureXml(floating.xml, { ...geometryOf(floating), wrap: null })
    expect(inline).toContain('<wp:inline')
    expect(inline).not.toContain('<wp:anchor')
    expect((await reparse(inline)).wrap).toBeUndefined()

    const drawing = /<w:drawing>[\s\S]*<\/w:drawing>/.exec(IMAGE_PARAGRAPH_XML)![0]
    const square = patchRunPictureXml(drawing, { wrap: 'square-right' })
    expect(square).toContain('<wp:anchor')
    expect(square).toContain('<wp:align>right</wp:align>')
  })

  it('writes a turn and a flip on the picture a:xfrm, with the room the turn takes', async () => {
    const floating = await runPicture(anchoredPicture({ posH: posOffset(0), posV: posOffset(0) }))
    const same = { ...geometryOf(floating), rotDeg: 0, flipH: false, flipV: false }
    expect(patchRunPictureXml(floating.xml, same)).toBe(floating.xml)

    const turned = patchRunPictureXml(floating.xml, { ...same, rotDeg: 90, flipV: true })
    expect(turned).toContain('<a:xfrm rot="5400000" flipV="1">')
    // 200 x 133.3 px turned a quarter: the bounding box is taller by the difference
    expect(turned).toContain('<wp:effectExtent l="0" t="317500" r="0" b="317500"/>')
    expect(turned).toContain(`<wp:positionH relativeFrom="column">${posOffset(0)}</wp:positionH>`)
    expect(await reparse(turned)).toMatchObject({ rotDeg: 90, flipV: true, wrap: floating.wrap })
    // what the file already says is not written again
    expect(patchRunPictureXml(turned, { ...same, rotDeg: 90, flipV: true })).toBe(turned)

    const back = patchRunPictureXml(turned, same)
    expect(back).toContain('<a:xfrm>')
    expect(back).not.toMatch(/\brot=|flipV=/)
    const again = await reparse(back)
    expect([again.rotDeg, again.flipV]).toEqual([undefined, undefined])
  })

  it('keeps the room of a turn in step with a resize, in the same save or a later one', async () => {
    const floating = await runPicture(anchoredPicture({ posH: posOffset(0), posV: posOffset(0) }))
    const turnedAt = (xml: string, widthPx: number, heightPx: number) =>
      patchRunPictureXml(xml, { ...geometryOf(floating), widthPx, heightPx, rotDeg: 90 })
    // 100 x 40 px: 30 px more above and below
    const both = turnedAt(floating.xml, 100, 40)
    expect(both).toContain('<wp:extent cx="952500" cy="381000"/>')
    expect(both).toContain('<wp:effectExtent l="0" t="285750" r="0" b="285750"/>')
    const later = turnedAt(turnedAt(floating.xml, 200, 133), 100, 40)
    expect(later).toContain('<wp:effectExtent l="0" t="285750" r="0" b="285750"/>')
    expect(later.match(/<wp:effectExtent\b/g)).toHaveLength(1)
    expect(turnedAt(later, 100, 40)).toBe(later)
  })

  it('gives a picture without an a:xfrm one to hold the turn', () => {
    const drawing = /<w:drawing>[\s\S]*<\/w:drawing>/.exec(IMAGE_PARAGRAPH_XML)![0]
    const bare = drawing.replace(/<pic:spPr\b[\s\S]*?<\/pic:spPr>/, '')
    expect(bare).not.toContain('<a:xfrm')
    const extent = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(bare)!
    const xfrm = `<a:xfrm rot="10800000"><a:off x="0" y="0"/><a:ext cx="${extent[1]}" cy="${extent[2]}"/></a:xfrm>`
    expect(patchRunPictureXml(bare, { rotDeg: 180 })).toContain(
      `</pic:blipFill><pic:spPr>${xfrm}</pic:spPr>`,
    )
    for (const empty of ['<pic:spPr/>', '<pic:spPr bwMode="auto"></pic:spPr>']) {
      const held = bare.replace('</pic:blipFill>', `</pic:blipFill>${empty}`)
      expect(patchRunPictureXml(held, { rotDeg: 180 })).toContain(
        `${empty.replace(/\/?>.*$/, '>')}${xfrm}</pic:spPr>`,
      )
    }
    // nothing to turn: the picture stays as bare as it was
    expect(patchRunPictureXml(bare, { rotDeg: 0, flipH: false })).toBe(bare)
  })

  it('leaves VML pictures alone', () => {
    const vml =
      '<w:pict><v:shape style="width:10pt;height:10pt"><v:imagedata r:id="rId1"/></v:shape></w:pict>'
    expect(patchRunPictureXml(vml, { wrap: 'front', offsetXEmu: 5 })).toBe(vml)
  })
})

describe('run picture alternative text', () => {
  const drawing = /<w:drawing>[\s\S]*<\/w:drawing>/
    .exec(IMAGE_PARAGRAPH_XML)![0]
    .replace(/(<wp:extent\b[^>]*\/>)/, '$1<wp:docPr id="1" name="Picture 1"/>')

  it('reads, sets and removes wp:docPr descr, leaving the rest of the drawing alone', () => {
    const bare = setRunPictureAlt(drawing, '')
    expect(runPictureAltOf(bare)).toBe('')
    const described = setRunPictureAlt(bare, 'Sales "Q3" & <costs>')
    expect(described).toContain('descr="Sales &quot;Q3&quot; &amp; &lt;costs&gt;"')
    expect(runPictureAltOf(described)).toBe('Sales "Q3" & <costs>')
    expect(setRunPictureAlt(described, 'Chart').match(/descr=/g)).toHaveLength(1)
    expect(setRunPictureAlt(described, '')).toBe(bare)
  })

  it('writes the picture\u2019s own docPr, not one further in', () => {
    const nested = drawing.replace('</wp:inline>', '<wp:docPr id="9" name="Inner"/></wp:inline>')
    const out = setRunPictureAlt(nested, 'Outer')
    expect(out.match(/descr="Outer"/g)).toHaveLength(1)
    expect(out.indexOf('descr="Outer"')).toBeLessThan(out.indexOf('name="Inner"'))
  })

  it('has no alternative text for a VML picture', () => {
    const vml =
      '<w:pict><v:shape style="width:10pt;height:10pt"><v:imagedata r:id="rId1"/></v:shape></w:pict>'
    expect(runPictureAltOf(vml)).toBeNull()
    expect(setRunPictureAlt(vml, 'Logo')).toBe(vml)
  })
})
