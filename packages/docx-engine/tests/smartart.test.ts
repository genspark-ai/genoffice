import { describe, expect, it } from 'vitest'
import { buildSmartArtParts, type SmartArtNode } from '../src/smartart'
import {
  SMARTART_CONTENT_TYPES,
  SMARTART_RADIAL_LAYOUT_ID,
  SMARTART_REL_TYPES,
} from '../src/smartart'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

const tree: SmartArtNode = {
  label: 'Mind map',
  children: [
    { label: 'Product', children: [{ label: 'Docs' }, { label: 'Sheets' }] },
    { label: 'Architecture', children: [{ label: 'Main process' }] },
  ],
}

const REL_IDS = {
  data: 'rId31',
  layout: 'rId32',
  quickStyle: 'rId33',
  colors: 'rId34',
  drawing: 'rId35',
}

const parts = buildSmartArtParts(tree, {}, REL_IDS)

const relsXml = [
  `<Relationship Id="${REL_IDS.data}" Type="${SMARTART_REL_TYPES.data}" Target="diagrams/data1.xml"/>`,
  `<Relationship Id="${REL_IDS.layout}" Type="${SMARTART_REL_TYPES.layout}" Target="diagrams/layout1.xml"/>`,
  `<Relationship Id="${REL_IDS.quickStyle}" Type="${SMARTART_REL_TYPES.quickStyle}" Target="diagrams/quickStyle1.xml"/>`,
  `<Relationship Id="${REL_IDS.colors}" Type="${SMARTART_REL_TYPES.colors}" Target="diagrams/colors1.xml"/>`,
  `<Relationship Id="${REL_IDS.drawing}" Type="${SMARTART_REL_TYPES.drawing}" Target="diagrams/drawing1.xml"/>`,
].join('')

async function buildSmartArtDocx(): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: parts.paragraphXml,
    extraRels: relsXml,
    extraParts: [
      {
        path: 'word/diagrams/data1.xml',
        xml: parts.dataXml,
        contentType: SMARTART_CONTENT_TYPES.data,
      },
      {
        path: 'word/diagrams/layout1.xml',
        xml: parts.layoutXml,
        contentType: SMARTART_CONTENT_TYPES.layout,
      },
      {
        path: 'word/diagrams/quickStyle1.xml',
        xml: parts.quickStyleXml,
        contentType: SMARTART_CONTENT_TYPES.quickStyle,
      },
      {
        path: 'word/diagrams/colors1.xml',
        xml: parts.colorsXml,
        contentType: SMARTART_CONTENT_TYPES.colors,
      },
      {
        path: 'word/diagrams/drawing1.xml',
        xml: parts.drawingXml,
        contentType: SMARTART_CONTENT_TYPES.drawing,
      },
    ],
  })
}

describe('buildSmartArtParts: the data model', () => {
  it('writes a data point, a transition pair and a presentation point per node', () => {
    const data = parts.dataXml
    // 6 nodes: root, Product, Docs, Sheets, Architecture, Main process
    const points = [...data.matchAll(/<dgm:pt modelId="([^"]+)"/g)]
    const dataPoints = [
      ...data.matchAll(/<dgm:pt modelId="[^"]+"(?![^>]*type="(?:pres|parTrans|sibTrans)")/g),
    ]
    const presPoints = [...data.matchAll(/<dgm:pt [^>]*type="pres"/g)]
    const parTrans = [...data.matchAll(/type="parTrans"/g)]
    const sibTrans = [...data.matchAll(/type="sibTrans"/g)]

    expect(dataPoints.length).toBe(6)
    expect(presPoints.length).toBe(6 + 5) // one per node, plus one per link
    expect(parTrans.length).toBe(5) // one per connection
    expect(sibTrans.length).toBe(5)
    expect(points.length).toBe(6 + 10 + 11)
  })

  it('points every presentation point at a point that exists', () => {
    const ids = new Set([...parts.dataXml.matchAll(/<dgm:pt modelId="([^"]+)"/g)].map((m) => m[1]))
    const assoc = [...parts.dataXml.matchAll(/presAssocID="([^"]+)"/g)].map((m) => m[1])
    expect(assoc.length).toBeGreaterThan(0)
    for (const id of assoc) expect(ids.has(id)).toBe(true)
  })

  it('links every node to its parent and maps it to a presentation point', () => {
    const data = parts.dataXml
    const plainCxns = [...data.matchAll(/<dgm:cxn (?![^>]*type=)[^>]*\/>/g)]
    const presOf = [...data.matchAll(/type="presOf"/g)]
    const presParOf = [...data.matchAll(/type="presParOf"/g)]

    expect(plainCxns.length).toBe(5) // 6 nodes form 5 parent-child links
    expect(presOf.length).toBe(6 + 5)
    expect(presParOf.length).toBe(5 + 5)
  })

  it('carries the labels of every node', () => {
    for (const label of ['Mind map', 'Product', 'Docs', 'Sheets', 'Architecture', 'Main process']) {
      expect(parts.dataXml).toContain(`<a:t>${label}</a:t>`)
    }
  })

  it('escapes a label that would break out of the XML', () => {
    const risky = buildSmartArtParts({ label: 'a & b <c>' }, {}, REL_IDS)
    expect(risky.dataXml).toContain('<a:t>a &amp; b &lt;c&gt;</a:t>')
    expect(risky.dataXml).not.toContain('<c>')
  })
})

describe('buildSmartArtParts: the drawing', () => {
  it('writes one shape per node plus the connectors', () => {
    const boxes = [...parts.drawingXml.matchAll(/<a:prstGeom prst="roundRect"/g)]
    const lines = [...parts.drawingXml.matchAll(/<a:prstGeom prst="line"/g)]
    expect(boxes.length).toBe(6)
    // the org-chart layout draws each link as several segments
    expect(lines.length).toBeGreaterThanOrEqual(5)
  })

  it('places the shapes inside the frame it declares', () => {
    const offsets = [...parts.drawingXml.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/g)].map(
      (m) => ({
        x: Number(m[1]),
        y: Number(m[2]),
      }),
    )
    expect(offsets.length).toBeGreaterThan(0)
    const widthEmu = Math.round(parts.widthPx * 9525)
    const heightEmu = Math.round(parts.heightPx * 9525)
    for (const off of offsets) {
      expect(off.x).toBeGreaterThanOrEqual(0)
      expect(off.y).toBeGreaterThanOrEqual(0)
      expect(off.x).toBeLessThanOrEqual(widthEmu)
      expect(off.y).toBeLessThanOrEqual(heightEmu)
    }
  })

  it('places the shapes inside the frame it declares', () => {
    const offsets = [...parts.drawingXml.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/g)].map(
      (m) => ({
        x: Number(m[1]),
        y: Number(m[2]),
      }),
    )
    expect(offsets.length).toBeGreaterThan(0)
    const widthEmu = Math.round(parts.widthPx * 9525)
    const heightEmu = Math.round(parts.heightPx * 9525)
    for (const off of offsets) {
      expect(off.x).toBeGreaterThanOrEqual(0)
      expect(off.y).toBeGreaterThanOrEqual(0)
      expect(off.x).toBeLessThanOrEqual(widthEmu)
      expect(off.y).toBeLessThanOrEqual(heightEmu)
    }
  })

  it('is deterministic', () => {
    const again = buildSmartArtParts(tree, {}, REL_IDS)
    expect(again.dataXml).toBe(parts.dataXml)
    expect(again.drawingXml).toBe(parts.drawingXml)
    expect(again.layoutXml).toBe(parts.layoutXml)
  })
})

describe('buildSmartArtParts: the document paragraph', () => {
  it('points at the four diagram parts', () => {
    expect(parts.paragraphXml).toContain(`r:dm="${REL_IDS.data}"`)
    expect(parts.paragraphXml).toContain(`r:lo="${REL_IDS.layout}"`)
    expect(parts.paragraphXml).toContain(`r:qs="${REL_IDS.quickStyle}"`)
    expect(parts.paragraphXml).toContain(`r:cs="${REL_IDS.colors}"`)
    expect(parts.paragraphXml).toContain('drawingml/2006/diagram')
  })
})

describe('buildSmartArtParts: the radial layout', () => {
  const radial = buildSmartArtParts(tree, { orientation: 'radial' }, REL_IDS)

  it('puts the topic in the middle and the branches around it', () => {
    // the first shape is the topic
    const first = /<dsp:spPr>([\s\S]*?)<\/dsp:spPr>/.exec(radial.drawingXml)![1]
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(first)!
    const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(first)!
    const topicX = Number(off[1]) / 9525 + Number(ext[1]) / 9525 / 2
    const topicY = Number(off[2]) / 9525 + Number(ext[2]) / 9525 / 2
    expect(topicX).toBeCloseTo(radial.widthPx / 2, 0)
    expect(topicY).toBeCloseTo(radial.heightPx / 2, 0)

    // the branches sit on a ring around it
    const centre = radial.widthPx / 2
    const distance = [...radial.drawingXml.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/g)].map(
      (m) => Math.hypot(Number(m[1]) / 9525 - centre, Number(m[2]) / 9525 - centre),
    )
    expect(Math.max(...distance)).toBeGreaterThan(60)
  })

  it('meets every connector on the borders of the boxes it joins', () => {
    const EMU = 9525
    interface Shape {
      x: number
      y: number
      w: number
      h: number
      line: boolean
      flip: boolean
    }
    const shapes: Shape[] = []
    for (const m of radial.drawingXml.matchAll(/<dsp:sp [^>]*>([\s\S]*?)<\/dsp:sp>/g)) {
      const spPr = /<dsp:spPr>([\s\S]*?)<\/dsp:spPr>/.exec(m[1])?.[1]
      if (!spPr) continue
      const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(spPr)
      const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(spPr)
      if (!off || !ext) continue
      shapes.push({
        x: Number(off[1]) / EMU,
        y: Number(off[2]) / EMU,
        w: Number(ext[1]) / EMU,
        h: Number(ext[2]) / EMU,
        line: /<a:prstGeom prst="line"/.test(spPr),
        flip: /<a:xfrm[^>]*flipV="1"/.test(spPr),
      })
    }
    const boxes = shapes.filter((s) => !s.line)
    const lines = shapes.filter((s) => s.line)
    expect(boxes.length).toBe(6)
    expect(lines.length).toBeGreaterThanOrEqual(5)

    /** the two ends of the diagonal a line shape draws */
    const endsOf = (l: Shape): Array<{ x: number; y: number }> =>
      l.flip
        ? [
            { x: l.x, y: l.y + l.h },
            { x: l.x + l.w, y: l.y },
          ]
        : [
            { x: l.x, y: l.y },
            { x: l.x + l.w, y: l.y + l.h },
          ]

    const onBorder = (b: Shape, p: { x: number; y: number }): boolean => {
      const tol = 2
      const inside =
        p.x >= b.x - tol && p.x <= b.x + b.w + tol && p.y >= b.y - tol && p.y <= b.y + b.h + tol
      const edge =
        Math.abs(p.x - b.x) <= tol ||
        Math.abs(p.x - (b.x + b.w)) <= tol ||
        Math.abs(p.y - b.y) <= tol ||
        Math.abs(p.y - (b.y + b.h)) <= tol
      return inside && edge
    }

    for (const line of lines) {
      const [start, end] = endsOf(line)
      expect(boxes.some((b) => onBorder(b, start))).toBe(true)
      expect(boxes.some((b) => onBorder(b, end))).toBe(true)
    }
  })

  it('lays the branches out with the cycle algorithm', () => {
    expect(radial.layoutXml).toContain('<dgm:alg type="cycle">')
    expect(radial.layoutXml).toContain(SMARTART_RADIAL_LAYOUT_ID)
  })

  it('points the data model at the radial layout', () => {
    expect(radial.dataXml).toContain(`loTypeId="${SMARTART_RADIAL_LAYOUT_ID}"`)
    expect(radial.dataXml).toContain(`presId="${SMARTART_RADIAL_LAYOUT_ID}"`)
    expect(radial.dataXml).toContain('presName="radialNode"')
  })

  it('gives the frame the shape of the drawing', () => {
    expect(radial.widthPx).toBe(radial.heightPx)
    expect(radial.widthPx).toBeGreaterThan(0)
  })
})

describe('a document holding the diagram', () => {
  it('reads back: the engine finds the shapes and their labels', async () => {
    const bytes = await buildSmartArtDocx()
    const parsed = await parseDocx(bytes)
    const diagram = parsed.blocks.find((b) => b.diagramDisplay)?.diagramDisplay
    expect(diagram).toBeTruthy()
    if (!diagram) return

    // one box per node, plus the connector segments
    const boxes = diagram.shapes.filter((s) => s.prst === 'roundRect')
    const lines = diagram.shapes.filter((s) => s.prst === 'line')
    expect(boxes.length).toBe(6)
    expect(lines.length).toBeGreaterThanOrEqual(5)
    expect(diagram.widthPx).toBeGreaterThan(0)
    expect(diagram.heightPx).toBeGreaterThan(0)

    const texts = diagram.shapes.flatMap((s) => s.texts ?? []).join(' ')
    for (const label of ['Mind map', 'Product', 'Docs', 'Sheets', 'Architecture', 'Main process']) {
      expect(texts).toContain(label)
    }
  })
})
