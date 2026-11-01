import { describe, expect, it } from 'vitest'
import { buildSmartArtParts, type SmartArtNode } from '../src/smartart'
import { SMARTART_CONTENT_TYPES, SMARTART_REL_TYPES } from '../src/smartart'
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
      { path: 'word/diagrams/data1.xml', xml: parts.dataXml, contentType: SMARTART_CONTENT_TYPES.data },
      { path: 'word/diagrams/layout1.xml', xml: parts.layoutXml, contentType: SMARTART_CONTENT_TYPES.layout },
      {
        path: 'word/diagrams/quickStyle1.xml',
        xml: parts.quickStyleXml,
        contentType: SMARTART_CONTENT_TYPES.quickStyle,
      },
      { path: 'word/diagrams/colors1.xml', xml: parts.colorsXml, contentType: SMARTART_CONTENT_TYPES.colors },
      { path: 'word/diagrams/drawing1.xml', xml: parts.drawingXml, contentType: SMARTART_CONTENT_TYPES.drawing },
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
    const offsets = [...parts.drawingXml.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/g)].map((m) => ({
      x: Number(m[1]),
      y: Number(m[2]),
    }))
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
