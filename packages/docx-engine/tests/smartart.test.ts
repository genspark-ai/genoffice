import { describe, expect, it } from 'vitest'
import {
  buildSmartArtParts,
  SMARTART_CONTENT_TYPES,
  SMARTART_LAYOUT_ID,
  SMARTART_REL_TYPES,
  type SmartArtNode,
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

interface Point {
  id: string
  attrs: string
  body: string
}

function pointsOf(source = parts.dataXml): Point[] {
  return [...source.matchAll(/<dgm:pt ([^>]*)>([\s\S]*?)<\/dgm:pt>/g)].map((m) => ({
    id: attr(m[1]!, 'modelId')!,
    attrs: m[1]!,
    body: m[2]!,
  }))
}

function attr(fragment: string, name: string): string | undefined {
  return new RegExp(`${name}="([^"]*)"`).exec(fragment)?.[1]
}

function cxnsOf(source = parts.dataXml): string[] {
  return [...source.matchAll(/<dgm:cxn ([^>]*)\/>/g)].map((m) => m[1]!)
}

describe('buildSmartArtParts: the layout', () => {
  it("emits Word's built-in hierarchy1 definition verbatim", () => {
    const layout = parts.layoutXml
    expect(layout.includes(`uniqueId="${SMARTART_LAYOUT_ID}"`)).toBe(true)
    // the recursion Word needs to lay a tree out: hierChild -> forEach ->
    // hierRoot -> composite -> text, then hierChild again
    const structure = [
      '<dgm:alg type="hierChild">',
      '<dgm:forEach',
      'name="hierRoot1"',
      '<dgm:alg type="composite"/>',
      '<dgm:alg type="tx"/>',
      '<dgm:layoutNode name="hierChild5">',
      '<dgm:forEach name="Name31" ref="repeat"/>',
    ]
    for (const marker of structure) expect(layout.includes(marker), marker).toBe(true)
    // the hand-written definition Word used to replace is gone
    expect(layout.includes('rootComposite1')).toBe(false)
    expect(layout.includes('urn:genoffice/officeart/layout')).toBe(false)
  })

  it('is byte-identical on every build', () => {
    const again = buildSmartArtParts(tree, {}, REL_IDS)
    expect(again.dataXml).toBe(parts.dataXml)
    expect(again.drawingXml).toBe(parts.drawingXml)
    expect(again.layoutXml).toBe(parts.layoutXml)
    expect(again.quickStyleXml).toBe(parts.quickStyleXml)
    expect(again.colorsXml).toBe(parts.colorsXml)
  })
})

describe('buildSmartArtParts: the data model', () => {
  it('keeps the doc point empty and puts the title in a node point below it', () => {
    const points = pointsOf(parts.dataXml)
    const doc = points.find((p) => p.attrs.includes('type="doc"'))!
    expect(doc).toBeDefined()
    // a layout never presents the doc point's text, so the title must not be there
    expect(doc.body).not.toContain('<a:t>')
    expect(doc.body).not.toContain('Mind map')

    const title = points.find(
      (p) => !p.attrs.includes('type=') && p.body.includes('<a:t>Mind map</a:t>'),
    )
    expect(title, 'the root label needs a regular node point').toBeDefined()

    // and the document point is connected to that node, so the tree has a root
    const docLink = cxnsOf().find((c) => !c.includes('type=') && attr(c, 'srcId') === doc.id)
    expect(docLink).toBeDefined()
    expect(attr(docLink!, 'destId')).toBe(title!.id)
  })

  it('gives every node a data point and a parent/sibling transition pair', () => {
    const points = pointsOf(parts.dataXml)
    const data = points.filter((p) => !p.attrs.includes('type='))
    expect(data).toHaveLength(6) // one per node, the root included
    expect(points.filter((p) => p.attrs.includes('type="parTrans"'))).toHaveLength(6)
    expect(points.filter((p) => p.attrs.includes('type="sibTrans"'))).toHaveLength(6)

    // every transition pair belongs to the connection that brings a node in
    const connectionIds = new Set(
      cxnsOf()
        .filter((c) => !c.includes('type='))
        .map((c) => attr(c, 'modelId')),
    )
    for (const p of points.filter((q) => q.attrs.includes('Trans"'))) {
      expect(connectionIds.has(attr(p.attrs, 'cxnId'))).toBe(true)
    }
  })

  it('links every node to its parent and its transition pair', () => {
    const points = pointsOf(parts.dataXml)
    const data = points.filter((p) => !p.attrs.includes('type='))
    const links = cxnsOf().filter((c) => !c.includes('type='))
    expect(links).toHaveLength(6) // the root from the doc + five parent-child links
    const doc = points.find((p) => p.attrs.includes('type="doc"'))!
    for (const link of links) {
      const src = attr(link, 'srcId')
      expect(src === doc.id || data.some((p) => p.id === src)).toBe(true)
      expect(data.some((p) => p.id === attr(link, 'destId'))).toBe(true)
      // the link carries the transition pair of its destination
      const parTrans = points.find((p) => p.id === attr(link, 'parTransId'))
      const sibTrans = points.find((p) => p.id === attr(link, 'sibTransId'))
      expect(parTrans?.attrs).toContain('type="parTrans"')
      expect(sibTrans?.attrs).toContain('type="sibTrans"')
    }
  })

  it('presents the tree through the layout nodes Word expects', () => {
    const points = pointsOf(parts.dataXml)
    const pres = points.filter((p) => p.attrs.includes('type="pres"'))
    // one doc point, five presentation points per node, one connector per edge
    expect(pres).toHaveLength(1 + 6 * 5 + 5)

    const layoutNames = new Set([
      ...[...parts.layoutXml.matchAll(/<dgm:layoutNode name="([^"]+)"/g)].map((m) => m[1]!),
      ...[...parts.layoutXml.matchAll(/<dgm:forEach name="([^"]+)"/g)].map((m) => m[1]!),
    ])
    for (const p of pres) {
      const name = attr(p.body, 'presName') ?? attr(p.attrs, 'presName')
      expect(name, 'every presentation point names a layout node').toBeTruthy()
      expect(layoutNames.has(name!), `${name} is not a layout node`).toBe(true)
    }
  })

  it('presents a connector from the parent transition, not the sibling one', () => {
    const points = pointsOf(parts.dataXml)
    const connectorPres = points.filter((p) => {
      const name = attr(p.body, 'presName')
      return name === 'Name10' || name === 'Name17' || name === 'Name23'
    })
    expect(connectorPres).toHaveLength(5)
    const parTransIds = new Set(
      points.filter((p) => p.attrs.includes('type="parTrans"')).map((p) => p.id),
    )
    const sibTransIds = new Set(
      points.filter((p) => p.attrs.includes('type="sibTrans"')).map((p) => p.id),
    )
    for (const p of connectorPres) {
      const assoc = attr(p.body, 'presAssocID')!
      expect(parTransIds.has(assoc), 'connector must ride on the parTrans point').toBe(true)
      expect(sibTransIds.has(assoc)).toBe(false)
    }
  })

  it('presents a node through its text point and a connector through its parTrans', () => {
    const points = pointsOf(parts.dataXml)
    const byId = new Map(points.map((p) => [p.id, p]))
    const presOf = cxnsOf().filter((c) => c.includes('type="presOf"'))
    // one per node, one per connector (not the doc-to-root one), one for the doc
    expect(presOf).toHaveLength(6 + 5 + 1)
    for (const cxn of presOf) {
      const src = byId.get(attr(cxn, 'srcId')!)!
      const dest = byId.get(attr(cxn, 'destId')!)!
      const destName = attr(dest.body, 'presName')!
      expect(destName).toBeTruthy()
      if (src.attrs.includes('type="doc"')) {
        expect(destName).toBe('hierChild1')
      } else if (src.attrs.includes('type="parTrans"')) {
        expect(destName).toMatch(/^Name\d+$/)
      } else {
        // a node's data point is presented by the text layout node
        expect(destName).toMatch(/^text\d*$/)
      }
    }
  })

  it('wires every presentation point into one tree under the doc point', () => {
    const points = pointsOf(parts.dataXml)
    const pres = points.filter((p) => p.attrs.includes('type="pres"'))
    const presIds = new Set(pres.map((p) => p.id))
    const links = cxnsOf().filter((c) => c.includes('type="presParOf"'))
    expect(links).toHaveLength(pres.length - 1) // a tree over every presentation point

    const children = links.map((c) => attr(c, 'destId')!)
    for (const c of children) expect(presIds.has(c)).toBe(true)
    const docPres = pres.find((p) => attr(p.body, 'presName') === 'hierChild1')
    expect(docPres, 'the doc point needs a presentation point').toBeDefined()
    // every point but the doc one has exactly one presentation parent
    for (const p of pres) {
      const incoming = children.filter((c) => c === p.id).length
      expect(incoming, `${p.id} needs one presentation parent`).toBe(p.id === docPres!.id ? 0 : 1)
    }
    // and that makes the doc point's presentation the single root
    const top = pres.filter((p) => !children.includes(p.id))
    expect(top).toHaveLength(1)
    expect(top[0]!.id).toBe(docPres!.id)
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

describe('buildSmartArtParts: the cached drawing', () => {
  it('writes one shape per node plus the connectors', () => {
    const boxes = [...parts.drawingXml.matchAll(/<a:prstGeom prst="roundRect"/g)]
    const lines = [...parts.drawingXml.matchAll(/<a:prstGeom prst="line"/g)]
    expect(boxes.length).toBe(6)
    expect(lines.length).toBeGreaterThanOrEqual(5)
  })

  it('places the shapes inside the frame it declares', () => {
    const offsets = [...parts.drawingXml.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/g)].map(
      (m) => ({ x: Number(m[1]), y: Number(m[2]) }),
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

  it('derives the picture from the same constraints the layout declares', () => {
    // hierarchy1's box fact is 0.667: the interpreter's boxes share that aspect
    const extents = [...parts.drawingXml.matchAll(/<a:ext cx="(\d+)" cy="(\d+)"\/>/g)].map((m) => ({
      cx: Number(m[1]),
      cy: Number(m[2]),
    }))
    const boxExtents = extents.slice(0, 6)
    for (const e of boxExtents) {
      expect(e.cy / e.cx).toBeCloseTo(0.667, 2)
    }
  })
})

describe('buildSmartArtParts: the page limit', () => {
  it('keeps the default frame inside the page column', () => {
    expect(parts.widthPx).toBeLessThanOrEqual(700)
    expect(parts.heightPx).toBeLessThanOrEqual(700)
    expect(parts.widthPx).toBeGreaterThan(0)
    expect(parts.heightPx).toBeGreaterThan(0)
  })

  it('scales a wider tree down to the limit', () => {
    const wide: SmartArtNode = {
      label: 'Root',
      children: Array.from({ length: 8 }, (_, i) => ({ label: `Branch ${i}` })),
    }
    for (const options of [{}, { widthPx: 900 }, { maxSizePx: 400 }]) {
      const laid = buildSmartArtParts(wide, options, REL_IDS)
      const limit = options.maxSizePx ?? 700
      expect(laid.widthPx).toBeLessThanOrEqual(limit)
      expect(laid.heightPx).toBeLessThanOrEqual(limit)
      if (options.widthPx) expect(laid.widthPx).toBeLessThanOrEqual(700)
    }
  })

  it('honours an explicit frame the user set within the limit', () => {
    const laid = buildSmartArtParts(tree, { widthPx: 480, maxSizePx: 700 }, REL_IDS)
    expect(laid.widthPx).toBe(480)
    expect(laid.heightPx).toBeLessThanOrEqual(700)
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
