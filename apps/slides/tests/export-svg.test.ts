import { describe, expect, it, vi } from 'vitest'
import type {
  ChartRenderNode,
  GroupRenderNode,
  PlacedBox,
  RenderNode,
  RenderSlide,
  RenderTextLayout,
  ShapeRenderNode,
  TableRenderNode,
} from '@genoffice/pptx-render'

import { arcPath, polylinePath, renderSlideToSvg, rotatedBounds } from '../src/renderer/export-svg'

function box(x: number, y: number, w: number, h: number, rotationDeg = 0): PlacedBox {
  return {
    x,
    y,
    w,
    h,
    rotationDeg,
    flipH: false,
    flipV: false,
    centerX: x + w / 2,
    centerY: y + h / 2,
  }
}

function text(
  runs: Array<Partial<RenderTextLayout['lines'][0]['runs'][0]> & { text: string }>,
): RenderTextLayout {
  return {
    lines: [
      {
        top: 0,
        height: 24,
        paraStart: true,
        runs: runs.map((r, i) => ({
          x: 0,
          baselineY: 19,
          fontFamily: 'Calibri',
          fontSizePx: 20,
          color: '#112233',
          bold: false,
          italic: false,
          underline: false,
          widthPx: 80,
          srcRunIdx: i,
          ...r,
        })),
      },
    ],
    insets: { l: 7.2, t: 3.6, r: 7.2, b: 3.6 },
    anchor: 'top',
    fontScale: 1,
    contentHeight: 24,
    wrap: true,
  }
}

function shape(over: Partial<ShapeRenderNode> = {}): ShapeRenderNode {
  return {
    id: 'n1',
    type: 'shape',
    sourceId: 's1',
    box: box(10, 20, 200, 100),
    fill: { kind: 'solid', color: '#FF0000' },
    ...over,
  }
}

function slide(nodes: RenderNode[]): RenderSlide {
  return {
    widthPx: 960,
    heightPx: 540,
    scale: 1,
    background: { kind: 'solid', color: '#FFFFFF' },
    nodes,
  }
}

describe('vector export: page frame', () => {
  it('emits an svg sized to the slide with the background first', async () => {
    const { svg } = await renderSlideToSvg(slide([]), new Map())
    expect(svg.startsWith('<svg xmlns="http://www.w3.org/2000/svg"')).toBe(true)
    expect(svg).toContain('viewBox="0 0 960 540"')
    expect(svg).toContain('<rect width="960" height="540" fill="#FFFFFF"/>')
    expect(svg.endsWith('</svg>')).toBe(true)
  })

  it('falls back to white when the slide has no background fill', async () => {
    const s = slide([])
    s.background = { kind: 'none' }
    const { svg } = await renderSlideToSvg(s, new Map())
    expect(svg).toContain('<rect width="960" height="540" fill="#ffffff"/>')
  })
})

describe('vector export: shapes', () => {
  it('draws a solid rectangle at its box with fill and stroke', async () => {
    const { svg } = await renderSlideToSvg(
      slide([
        shape({
          stroke: { color: '#00FF00', widthPx: 2, widthPt: 1.5, dash: [4, 2], cap: 'round' },
        }),
      ]),
      new Map(),
    )
    expect(svg).toContain('<g transform="translate(10 20)">')
    expect(svg).toContain(
      '<rect width="200" height="100" fill="#FF0000" stroke="#00FF00" stroke-width="2" stroke-dasharray="4 2" stroke-linecap="round"/>',
    )
  })

  it('pivots rotation and flips on the box center like the canvas backend', async () => {
    const n = shape({ box: { ...box(10, 20, 200, 100, 30), flipH: true } })
    const { svg } = await renderSlideToSvg(slide([n]), new Map())
    expect(svg).toContain(
      'transform="translate(110 70) rotate(30) scale(-1 1) translate(-100 -50)"',
    )
  })

  it('rounds rect corners, clamped to half the short side', async () => {
    const { svg } = await renderSlideToSvg(
      slide([shape({ presetGeometry: 'roundRect', cornerRadiusPx: 80 })]),
      new Map(),
    )
    expect(svg).toContain('<rect width="200" height="100" rx="50" ry="50" fill="#FF0000"/>')
  })

  it('draws ellipses, polygons and custom paths', async () => {
    const { svg } = await renderSlideToSvg(
      slide([
        shape({ id: 'e', presetGeometry: 'ellipse' }),
        shape({ id: 'p', polygonPoints: [0, 0, 200, 0, 100, 100] }),
        shape({ id: 'c', pathData: 'M0 0L200 100Z', fill: { kind: 'none' } }),
      ]),
      new Map(),
    )
    expect(svg).toContain('<ellipse cx="100" cy="50" rx="100" ry="50" fill="#FF0000"/>')
    expect(svg).toContain(
      '<polygon points="0 0 200 0 100 100" stroke-linejoin="round" fill="#FF0000"/>',
    )
    expect(svg).toContain('<path d="M0 0L200 100Z" fill="none" stroke-linejoin="round"/>')
  })

  it('keeps the round join of polygons under a stroke without an explicit join', async () => {
    const { svg } = await renderSlideToSvg(
      slide([
        shape({
          polygonPoints: [0, 0, 200, 0, 100, 100],
          stroke: { color: '#000000', widthPx: 3, widthPt: 2.25 },
        }),
      ]),
      new Map(),
    )
    expect(svg).toContain(
      '<polygon points="0 0 200 0 100 100" stroke-linejoin="round" fill="#FF0000" stroke="#000000" stroke-width="3"/>',
    )
  })

  it('expresses linear gradients as userSpaceOnUse defs spanning the box', async () => {
    const { svg } = await renderSlideToSvg(
      slide([
        shape({
          fill: {
            kind: 'gradient',
            angleDeg: 90,
            stops: [
              { pos: 0, color: '#000000' },
              { pos: 1, color: '#FFFFFF' },
            ],
          },
        }),
      ]),
      new Map(),
    )
    expect(svg).toMatch(
      /<defs><linearGradient id="lg1" gradientUnits="userSpaceOnUse" x1="100" y1="0" x2="100" y2="100">/,
    )
    expect(svg).toContain('<stop offset="0" stop-color="#000000"/>')
    expect(svg).toContain('<stop offset="1" stop-color="#FFFFFF"/>')
    expect(svg).toContain('fill="url(#lg1)"')
  })

  it('prefixes def ids per page so pages sharing one print document never cross-reference', async () => {
    const grad = shape({
      fill: {
        kind: 'gradient',
        angleDeg: 0,
        stops: [
          { pos: 0, color: '#000000' },
          { pos: 1, color: '#FFFFFF' },
        ],
      },
    })
    const { svg } = await renderSlideToSvg(slide([grad]), new Map(), { idPrefix: 'p7-' })
    expect(svg).toContain('<linearGradient id="p7-lg1"')
    expect(svg).toContain('fill="url(#p7-lg1)"')
  })

  it('splits alpha colors into stop-opacity', async () => {
    const { svg } = await renderSlideToSvg(
      slide([
        shape({
          fill: {
            kind: 'gradient',
            angleDeg: 0,
            stops: [
              { pos: 0, color: '#FF000080' },
              { pos: 1, color: '#FF0000' },
            ],
          },
        }),
      ]),
      new Map(),
    )
    expect(svg).toContain('<stop offset="0" stop-color="rgb(255,0,0)" stop-opacity="0.502"/>')
  })

  it('casts outer shadows through a drop-shadow filter with an explicit region', async () => {
    const { svg } = await renderSlideToSvg(
      slide([shape({ shadow: { color: '#00000080', blurPx: 6, offsetX: 3, offsetY: 4 } })]),
      new Map(),
    )
    expect(svg).toContain(
      '<filter id="sh1" filterUnits="userSpaceOnUse" x="-26" y="-26" width="252" height="152">',
    )
    expect(svg).toContain(
      '<feDropShadow dx="3" dy="4" stdDeviation="3" flood-color="rgb(0,0,0)" flood-opacity="0.502"/>',
    )
    expect(svg).toContain('filter="url(#sh1)"')
  })

  it('draws connectors as polylines with the arrow pointer at the tail', async () => {
    const { svg } = await renderSlideToSvg(
      slide([
        shape({
          fill: { kind: 'none' },
          stroke: { color: '#0000FF', widthPx: 2, widthPt: 1.5 },
          line: {
            points: [0, 0, 200, 100],
            tailEnd: { type: 'triangle', widthPx: 8, lengthPx: 10 },
          },
        }),
      ]),
      new Map(),
    )
    expect(svg).toContain(
      '<polyline points="0 0 200 100" fill="none" stroke="#0000FF" stroke-width="2"',
    )
    expect(svg).toContain(
      '<polygon points="0 0 -10 4 -10 -4" fill="#0000FF" stroke="#0000FF" stroke-width="2" stroke-linejoin="round" transform="translate(200 100) rotate(26.565)"/>',
    )
  })
})

describe('vector export: text', () => {
  it('places each glyph run on its baseline inside the insets as real text', async () => {
    const { svg, fontFamilies } = await renderSlideToSvg(
      slide([
        shape({ text: text([{ text: 'Hello & <world>', x: 5, bold: true, underline: true }]) }),
      ]),
      new Map(),
    )
    // x = inset 7.2 + run.x 5; y = inset 3.6 + baseline 19 (fallback face: no anchor drop)
    expect(svg).toMatch(
      /<text x="12.2" y="22.6" font-family="Calibri, Carlito, Arial, sans-serif" font-size="20" font-weight="700" text-decoration="underline" textLength="80" lengthAdjust="spacingAndGlyphs" xml:space="preserve" fill="#112233">Hello &amp; &lt;world&gt;<\/text>/,
    )
    expect([...fontFamilies]).toEqual(['Calibri, Carlito, Arial, sans-serif'])
  })

  it('skips textLength for single glyphs and whitespace, keeps letter spacing', async () => {
    const { svg } = await renderSlideToSvg(
      slide([
        shape({
          text: text([
            { text: 'A', widthPx: 12 },
            { text: '   ', widthPx: 15 },
            { text: 'ab', widthPx: 30, letterSpacingPx: 2 },
          ]),
        }),
      ]),
      new Map(),
    )
    expect(svg).toMatch(/<text[^>]*>A<\/text>/)
    expect(svg).not.toMatch(/<text[^>]*textLength[^>]*>A<\/text>/)
    expect(svg).not.toMatch(/<text[^>]*textLength[^>]*> {3}<\/text>/)
    // the trailing per-char spacing is not part of the visible advance
    expect(svg).toMatch(/<text[^>]*letter-spacing="2"[^>]*textLength="28"[^>]*>ab<\/text>/)
  })

  it('draws highlights behind runs and strips control characters', async () => {
    const { svg } = await renderSlideToSvg(
      slide([shape({ text: text([{ text: 'a\u0001b', highlight: '#FFFF00' }]) })]),
      new Map(),
    )
    expect(svg).toContain('<rect x="7.2" y="3.6" width="80" height="24" fill="#FFFF00"/>')
    expect(svg).toContain('>ab</text>')
  })

  it('anchors rtl runs at their right edge', async () => {
    const { svg } = await renderSlideToSvg(
      slide([shape({ text: text([{ text: 'שלום', rtl: true, x: 10, widthPx: 50 }]) })]),
      new Map(),
    )
    expect(svg).toMatch(/<text x="67.2"[^>]*direction="rtl" unicode-bidi="bidi-override"/)
  })

  it('counter-flips the text layer of a flipped shape', async () => {
    const { svg } = await renderSlideToSvg(
      slide([
        shape({ box: { ...box(10, 20, 200, 100), flipH: true }, text: text([{ text: 'ab' }]) }),
      ]),
      new Map(),
    )
    expect(svg).toContain('<g transform="translate(200 0) scale(-1 1)"><text')
  })

  it('hands WordArt warp to the raster fallback and places the raster in page space', async () => {
    const rasterize = vi.fn(async () => ({
      href: 'data:image/png;base64,AAAA',
      x: 1,
      y: 2,
      w: 300,
      h: 200,
    }))
    const warped = shape({
      text: { ...text([{ text: 'Warp' }]), txWarp: { prst: 'textArchUp' } },
    })
    const { svg } = await renderSlideToSvg(slide([warped]), new Map(), { rasterizeNode: rasterize })
    expect(rasterize).toHaveBeenCalledWith(warped)
    expect(svg).toContain(
      '<image href="data:image/png;base64,AAAA" x="1" y="2" width="300" height="200" preserveAspectRatio="none"/>',
    )
    expect(svg).not.toContain('<text')
  })

  it('rasterizes a node whose vector emission fails, keeping the other nodes', async () => {
    const rasterize = vi.fn(async () => ({
      href: 'data:image/png;base64,BBBB',
      x: 0,
      y: 0,
      w: 1,
      h: 1,
    }))
    // jsdom has no 2d canvas: the pattern fill cannot be painted and throws inside the builder
    const patterned = shape({
      id: 'pat',
      fill: { kind: 'pattern', preset: 'pct50', fg: '#000', bg: '#fff', cellPx: 8 },
    })
    const { svg } = await renderSlideToSvg(slide([patterned, shape({ id: 'ok' })]), new Map(), {
      rasterizeNode: rasterize,
    })
    expect(rasterize).toHaveBeenCalledTimes(1)
    expect(svg).toContain('data:image/png;base64,BBBB')
    expect(svg).toContain('<rect width="200" height="100" fill="#FF0000"/>')
  })
})

describe('vector export: composites', () => {
  it('nests group children in the group transform', async () => {
    const group: GroupRenderNode = {
      id: 'g',
      type: 'group',
      sourceId: 'g',
      box: box(100, 100, 400, 200),
      children: [shape({ box: box(10, 10, 50, 50) })],
    }
    const { svg } = await renderSlideToSvg(slide([group]), new Map())
    expect(svg).toContain(
      '<g transform="translate(100 100)"><g transform="translate(10 10)"><rect width="50" height="50" fill="#FF0000"/></g></g>',
    )
  })

  it('draws table cells with fills, borders and text offset into the cell', async () => {
    const table: TableRenderNode = {
      id: 't',
      type: 'table',
      sourceId: 't',
      box: box(0, 0, 200, 40),
      gridX: [0, 200],
      gridY: [0, 40],
      cells: [
        {
          x: 0,
          y: 0,
          w: 200,
          h: 40,
          row: 0,
          col: 0,
          fill: { kind: 'solid', color: '#EEEEEE' },
          borders: { b: { color: '#000000', widthPx: 1, widthPt: 0.75 } },
          text: text([{ text: 'cell' }]),
        },
      ],
    }
    const { svg } = await renderSlideToSvg(slide([table]), new Map())
    expect(svg).toContain('<rect x="0" y="0" width="200" height="40" fill="#EEEEEE"/>')
    expect(svg).toContain(
      '<line x1="0" y1="40" x2="200" y2="40" stroke="#000000" stroke-width="1"/>',
    )
    expect(svg).toMatch(/<text x="7.2" y="22.6"[^>]*>cell<\/text>/)
  })

  it('draws chart primitives: wedge paths, gridlines, bars, labels', async () => {
    const chart: ChartRenderNode = {
      id: 'c',
      type: 'chart',
      sourceId: 'c',
      box: box(0, 0, 400, 300),
      gridLines: [{ x1: 0, y1: 10, x2: 400, y2: 10, color: '#CCCCCC' }],
      axisLines: [],
      labels: [{ text: 'Q1', x: 10, y: 280, fontSizePx: 12, color: '#333333' }],
      bars: [{ x: 20, y: 100, w: 30, h: 150, color: '#4472C4' }],
      polylines: [{ points: [0, 0, 10, 10, 20, 0], color: '#ED7D31', widthPx: 2, smooth: true }],
      markers: [{ x: 10, y: 10, r: 3, color: '#ED7D31' }],
      swatches: [],
      wedges: [
        { cx: 200, cy: 150, outerR: 100, innerR: 0, startDeg: -90, sweepDeg: 90, color: '#4472C4' },
      ],
    }
    const { svg } = await renderSlideToSvg(slide([chart]), new Map())
    expect(svg).toContain(
      '<path d="M200 50A100 100 0 0 1 300 150L200 150Z" fill="#4472C4" stroke="#ffffff" stroke-width="1"/>',
    )
    expect(svg).toContain(
      '<line x1="0" y1="10" x2="400" y2="10" stroke="#CCCCCC" stroke-width="1"/>',
    )
    expect(svg).toContain('<rect x="20" y="100" width="30" height="150" fill="#4472C4"/>')
    expect(svg).toContain('<circle cx="10" cy="10" r="3" fill="#ED7D31"/>')
    expect(svg).toMatch(
      /<text x="10" y="286" dominant-baseline="central" font-family="Calibri, Carlito, Arial, sans-serif" font-size="12" fill="#333333" xml:space="preserve">Q1<\/text>/,
    )
    expect(svg).toMatch(/<path d="M0 0Q6 10 10 10Q14 10 20 0" fill="none" stroke="#ED7D31"/)
  })
})

describe('vector export: geometry helpers', () => {
  it('polylinePath: straight segments without tension, closed adds Z', () => {
    expect(polylinePath([0, 0, 10, 0, 10, 10], 0, false)).toBe('M0 0L10 0L10 10')
    expect(polylinePath([0, 0, 10, 0, 10, 10], 0, true)).toBe('M0 0L10 0L10 10Z')
    expect(polylinePath([0, 0], 0, false)).toBe('')
  })

  it('arcPath: full sweeps split into two arcs, rings walk back along the inner radius', () => {
    const full = arcPath(0, 0, 0, 10, -90, 360)
    expect(full.match(/A/g)).toHaveLength(2)
    expect(full.endsWith('L0 0Z')).toBe(true)
    const ring = arcPath(0, 0, 5, 10, 0, 180)
    expect(ring).toBe('M10 0A10 10 0 0 1 -10 0L-5 0A5 5 0 0 0 5 0Z')
  })

  it('rotatedBounds: axis-aligned envelope of a rotated box', () => {
    expect(rotatedBounds(box(0, 0, 100, 50))).toEqual({ x: 0, y: 0, w: 100, h: 50 })
    const b = rotatedBounds(box(0, 0, 100, 50, 90))
    expect(b.w).toBeCloseTo(50)
    expect(b.h).toBeCloseTo(100)
    expect(b.x).toBeCloseTo(25)
    expect(b.y).toBeCloseTo(-25)
  })
})
