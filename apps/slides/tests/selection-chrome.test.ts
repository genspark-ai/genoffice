import { describe, it, expect, vi } from 'vitest'

// react-konva's node entry requires the native 'canvas' package; only the pure helper is under test
vi.mock('react-konva', () => {
  const stub = () => null
  return {
    Stage: stub,
    Layer: stub,
    Rect: stub,
    Group: stub,
    Transformer: stub,
    Line: stub,
    Arrow: stub,
    Text: stub,
    Ellipse: stub,
    Image: stub,
    Path: stub,
    Circle: stub,
    Arc: stub,
  }
})

import { selectionChromeColor } from '../src/renderer/SlideCanvas'
import type { RenderSlide, RenderNode, RenderFill } from '@genoffice/pptx-render'

const DARK = '#ffffff' // chrome color on a dark background
const LIGHT = '#232425' // chrome color on a light background

const box = (x = 0, y = 0, w = 960, h = 540) => ({
  x,
  y,
  w,
  h,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})

function shapeNode(fill: RenderFill, flags: Partial<RenderNode> = {}): RenderNode {
  return {
    id: 's1',
    sourceId: 's1',
    type: 'shape',
    box: box(),
    fill,
    ...flags,
  } as unknown as RenderNode
}

function slide(background: RenderFill, nodes: RenderNode[] = []): RenderSlide {
  return {
    widthPx: 960,
    heightPx: 540,
    scale: 1,
    background,
    nodes,
  } as unknown as RenderSlide
}

const noImages = new Map<string, HTMLImageElement>()
const solid = (color: string): RenderFill => ({ kind: 'solid', color })

describe('selectionChromeColor', () => {
  it('near-black chrome on a white slide, white chrome on a dark slide', () => {
    expect(selectionChromeColor(slide(solid('#ffffff')), noImages)).toBe(LIGHT)
    expect(selectionChromeColor(slide(solid('#1a1a2e')), noImages)).toBe(DARK)
    expect(selectionChromeColor(slide({ kind: 'none' }), noImages)).toBe(LIGHT) // unknown → assume light
  })

  it('a fully transparent full-page overlay does not read as its base color', () => {
    // #00000000 = transparent black — visually invisible on the white slide,
    // but the alpha-blind average called it pure black → white (invisible) chrome
    const overlay = shapeNode(solid('#00000000'), { background: true })
    expect(selectionChromeColor(slide(solid('#ffffff'), [overlay]), noImages)).toBe(LIGHT)
    const rgbaOverlay = shapeNode(solid('rgba(0, 0, 0, 0)'), { background: true })
    expect(selectionChromeColor(slide(solid('#ffffff'), [rgbaOverlay]), noImages)).toBe(LIGHT)
  })

  it('a light tint overlay keeps the light verdict; a heavy dark scrim flips it', () => {
    // 10%-alpha black over white: effective luminance 0.9 → still a light background
    const tint = shapeNode(solid('#0000001a'), { background: true })
    expect(selectionChromeColor(slide(solid('#ffffff'), [tint]), noImages)).toBe(LIGHT)
    // 80%-alpha black over white: effective luminance 0.2 → dark background
    const scrim = shapeNode(solid('#000000cc'), { background: true })
    expect(selectionChromeColor(slide(solid('#ffffff'), [scrim]), noImages)).toBe(DARK)
  })

  it('an opaque background element still wins over the slide background', () => {
    const white = shapeNode(solid('#ffffff'), { background: true })
    expect(selectionChromeColor(slide(solid('#000000'), [white]), noImages)).toBe(LIGHT)
    const dark = shapeNode(solid('#111111'), { background: true })
    expect(selectionChromeColor(slide(solid('#ffffff'), [dark]), noImages)).toBe(DARK)
  })

  it('background layers composite in z-order (topmost opaque paint decides)', () => {
    const darkBase = shapeNode(solid('#000000'), { background: true })
    const whiteTop = shapeNode(solid('#ffffff'), { background: true })
    expect(selectionChromeColor(slide(solid('#808080'), [darkBase, whiteTop]), noImages)).toBe(
      LIGHT,
    )
    // transparent layer above an opaque dark one keeps the dark verdict
    const glass = shapeNode(solid('#ffffff00'), { background: true })
    expect(selectionChromeColor(slide(solid('#ffffff'), [darkBase, glass]), noImages)).toBe(DARK)
  })

  it('a gradient overlay alpha-weights its stops', () => {
    // both stops nearly transparent → the white slide still reads light
    const softGrad = shapeNode(
      {
        kind: 'gradient',
        stops: [
          { pos: 0, color: '#00000000' },
          { pos: 1, color: '#00000033' },
        ],
        angleDeg: 0,
      } as RenderFill,
      { background: true },
    )
    expect(selectionChromeColor(slide(solid('#ffffff'), [softGrad]), noImages)).toBe(LIGHT)
    // opaque dark gradient → dark
    const darkGrad = shapeNode(
      {
        kind: 'gradient',
        stops: [
          { pos: 0, color: '#111111' },
          { pos: 1, color: '#333333' },
        ],
        angleDeg: 0,
      } as RenderFill,
      { background: true },
    )
    expect(selectionChromeColor(slide(solid('#ffffff'), [darkGrad]), noImages)).toBe(DARK)
  })

  it('a full-page master decoration counts as background paint', () => {
    // dark slide background fully covered by a white full-page decoration → light chrome
    const deco = shapeNode(solid('#ffffff'), { decoration: true })
    expect(selectionChromeColor(slide(solid('#101010'), [deco]), noImages)).toBe(LIGHT)
    // a small decoration does not
    const badge = { ...shapeNode(solid('#ffffff'), { decoration: true }), box: box(0, 0, 100, 60) }
    expect(
      selectionChromeColor(slide(solid('#101010'), [badge as unknown as RenderNode]), noImages),
    ).toBe(DARK)
  })

  it('a full-page plate nested in a master decoration group counts as background paint', () => {
    const plate = { ...shapeNode(solid('#ffffff')), id: 'c1', sourceId: 'c1' }
    const group = {
      id: 'g1',
      sourceId: 'g1',
      type: 'group',
      box: box(),
      decoration: true,
      children: [plate],
    } as unknown as RenderNode
    expect(selectionChromeColor(slide(solid('#101010'), [group]), noImages)).toBe(LIGHT)
    // children are in group-local coords: this plate only covers the page once the
    // oversized group's offset is applied
    const shifted = {
      ...group,
      box: box(-100, -100, 1160, 740),
      children: [{ ...plate, box: box(100, 100, 960, 540) }],
    }
    expect(
      selectionChromeColor(slide(solid('#101010'), [shifted as unknown as RenderNode]), noImages),
    ).toBe(LIGHT)
    // a small child inside the covering group does not
    const badgeGroup = { ...group, children: [{ ...plate, box: box(0, 0, 100, 60) }] }
    expect(
      selectionChromeColor(
        slide(solid('#101010'), [badgeGroup as unknown as RenderNode]),
        noImages,
      ),
    ).toBe(DARK)
  })
})

describe('selectionChromeColor with a region (colored page + white plate)', () => {
  // Default new-deck template: green page, white rounded content plate, text on the plate.
  // The page verdict says dark (green) → white chrome → invisible on the white plate.
  const greenPage = solid('#3f9e2f') // lum ≈ 0.44 → dark
  const plate = { ...shapeNode(solid('#ffffff')), sourceId: 'plate', box: box(80, 40, 800, 460) }
  const textOnPlate = { x: 200, y: 150, w: 300, h: 60 }

  it('judges under the selection: dark chrome for a text box on the white plate', () => {
    const s = slide(greenPage, [plate as RenderNode])
    expect(selectionChromeColor(s, noImages)).toBe(DARK) // page verdict stays dark
    expect(selectionChromeColor(s, noImages, textOnPlate)).toBe(LIGHT) // plate under the box wins
  })

  it('keeps white chrome for a selection sitting on the bare colored page', () => {
    const s = slide(greenPage, [plate as RenderNode])
    expect(selectionChromeColor(s, noImages, { x: 0, y: 500, w: 100, h: 40 })).toBe(DARK)
  })

  it('a transparent overlay above the plate still does not flip the regional verdict', () => {
    const overlay = { ...shapeNode(solid('#00000000'), { background: true }), sourceId: 'ov' }
    const s = slide(greenPage, [plate as RenderNode, overlay as RenderNode])
    expect(selectionChromeColor(s, noImages, textOnPlate)).toBe(LIGHT)
  })

  it('table cell fills are sampled: a dark cell on the white plate gets white chrome', () => {
    const cell = (x: number, y: number, w: number, h: number, fill: RenderFill) => ({
      x,
      y,
      w,
      h,
      row: 0,
      col: 0,
      fill,
    })
    const table = {
      id: 't1',
      sourceId: 't1',
      type: 'table',
      box: box(100, 100, 400, 100),
      gridX: [0, 200, 400],
      gridY: [0, 100],
      cells: [cell(0, 0, 200, 100, solid('#1f2937')), cell(200, 0, 200, 100, solid('#ffffff'))],
    } as unknown as RenderNode
    const s = slide(greenPage, [plate as RenderNode, table])
    // the cell-edit overlay passes the cell box: dark cell → white chrome, white cell → dark
    expect(selectionChromeColor(s, noImages, { x: 100, y: 100, w: 200, h: 100 })).toBe(DARK)
    expect(selectionChromeColor(s, noImages, { x: 300, y: 100, w: 200, h: 100 })).toBe(LIGHT)
    // a table-style background under cells with no fill of their own
    const banded = {
      ...table,
      bgFill: solid('#111111'),
      cells: [cell(0, 0, 400, 100, { kind: 'none' })],
    } as unknown as RenderNode
    expect(
      selectionChromeColor(slide(greenPage, [plate as RenderNode, banded]), noImages, {
        x: 100,
        y: 100,
        w: 400,
        h: 100,
      }),
    ).toBe(DARK)
  })

  it('cells of a flipped or rotated table are sampled where the overlay draws them', () => {
    const cell = (x: number, y: number, w: number, h: number, fill: RenderFill) => ({
      x,
      y,
      w,
      h,
      row: 0,
      col: 0,
      fill,
    })
    const cells = [cell(0, 0, 200, 100, solid('#1f2937')), cell(200, 0, 200, 100, solid('#ffffff'))]
    const flipped = {
      id: 't2',
      sourceId: 't2',
      type: 'table',
      box: { ...box(100, 100, 400, 100), flipH: true },
      gridX: [0, 200, 400],
      gridY: [0, 100],
      cells,
    } as unknown as RenderNode
    const s = slide(greenPage, [plate as RenderNode, flipped])
    // flipH mirrors the grid: the dark logical-left cell is drawn on the right
    expect(selectionChromeColor(s, noImages, { x: 300, y: 100, w: 200, h: 100 })).toBe(DARK)
    expect(selectionChromeColor(s, noImages, { x: 100, y: 100, w: 200, h: 100 })).toBe(LIGHT)
    // a 90° table stands upright around its center: the dark first cell lands above it
    // (the sampler uses the overlay's unrotated cell box, exactly what the edit frame reports)
    const rotated = { ...flipped, box: { ...box(100, 100, 400, 100), rotationDeg: 90 } }
    const r = slide(greenPage, [plate as RenderNode, rotated as unknown as RenderNode])
    expect(selectionChromeColor(r, noImages, { x: 250, y: 0, w: 100, h: 100 })).toBe(DARK)
    expect(selectionChromeColor(r, noImages, { x: 250, y: 200, w: 100, h: 100 })).toBe(LIGHT)
  })

  it('chart-space and plot-area fills are sampled', () => {
    const chart = {
      id: 'c1',
      sourceId: 'c1',
      type: 'chart',
      box: box(100, 100, 400, 300),
      bgFill: solid('#0b1220'),
      gridLines: [],
    } as unknown as RenderNode
    const s = slide(greenPage, [plate as RenderNode, chart])
    expect(selectionChromeColor(s, noImages, { x: 100, y: 100, w: 400, h: 300 })).toBe(DARK)
    const plotOnly = {
      ...chart,
      bgFill: undefined,
      plotRect: { x: 50, y: 50, w: 300, h: 200, fill: solid('#0b1220') },
    } as unknown as RenderNode
    // the selection over the plot area reads dark; the chart's white margin on the plate stays light
    expect(
      selectionChromeColor(slide(greenPage, [plate as RenderNode, plotOnly]), noImages, {
        x: 150,
        y: 150,
        w: 300,
        h: 200,
      }),
    ).toBe(DARK)
    expect(
      selectionChromeColor(slide(greenPage, [plate as RenderNode, plotOnly]), noImages, {
        x: 100,
        y: 100,
        w: 40,
        h: 40,
      }),
    ).toBe(LIGHT)
  })

  it('empty or degenerate regions fall back to the page verdict', () => {
    const s = slide(greenPage, [plate as RenderNode])
    expect(selectionChromeColor(s, noImages, { x: 0, y: 0, w: 0, h: 0 })).toBe(DARK)
    expect(selectionChromeColor(s, noImages, null)).toBe(DARK)
  })
})
