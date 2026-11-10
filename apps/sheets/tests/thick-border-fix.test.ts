import { BorderStyleTypes } from '@univerjs/core'
import { Border, getLineWidth } from '@univerjs/engine-render'
import { beforeAll, describe, expect, it, vi } from 'vitest'

import { printBorderWidthPt } from '../src/renderer/print-html'
import {
  THICK_BORDER_DEVICE_PX,
  borderStrokeCssPx,
  borderStrokeDevicePx,
  installThickBorderFix,
} from '../src/renderer/thick-border-fix'
import { mapBorderStyle } from '../src/renderer/univer-sync'

describe('mapBorderStyle', () => {
  it('maps the three OOXML weights onto distinct Univer styles', () => {
    expect(mapBorderStyle('thin')).toBe(BorderStyleTypes.THIN)
    expect(mapBorderStyle('medium')).toBe(BorderStyleTypes.MEDIUM)
    expect(mapBorderStyle('thick')).toBe(BorderStyleTypes.THICK)
    expect(mapBorderStyle('mediumDashed')).toBe(BorderStyleTypes.MEDIUM_DASHED)
    expect(mapBorderStyle('nonsense')).toBe(BorderStyleTypes.THIN)
  })
})

describe('border stroke widths', () => {
  it('orders thick > medium > thin in device pixels', () => {
    const thin = borderStrokeDevicePx(BorderStyleTypes.THIN)
    const medium = borderStrokeDevicePx(BorderStyleTypes.MEDIUM)
    const thick = borderStrokeDevicePx(BorderStyleTypes.THICK)
    expect(thin).toBe(1)
    expect(medium).toBe(2)
    expect(thick).toBe(THICK_BORDER_DEVICE_PX)
    expect(thick).toBeGreaterThan(medium)
    expect(medium).toBeGreaterThan(thin)
  })

  it('keeps the stock widths for every non-thick style and widens only thick', () => {
    for (const style of Object.values(BorderStyleTypes)) {
      if (typeof style !== 'number' || style === BorderStyleTypes.THICK) continue
      expect(borderStrokeDevicePx(style)).toBe(getLineWidth(style))
    }
    // Stock Univer's 3px thick shares the footprint of its smeared 2px medium.
    expect(borderStrokeDevicePx(BorderStyleTypes.THICK)).toBeGreaterThan(
      getLineWidth(BorderStyleTypes.THICK),
    )
  })

  it('scales with zoom exactly like the built-in strokes and stays ordered', () => {
    for (const dpr of [1, 2]) {
      for (const zoom of [0.5, 1, 2]) {
        const scale = dpr * zoom
        const thin = borderStrokeCssPx(BorderStyleTypes.THIN, scale)
        const medium = borderStrokeCssPx(BorderStyleTypes.MEDIUM, scale)
        const thick = borderStrokeCssPx(BorderStyleTypes.THICK, scale)
        // Same convention as Univer's setLineWidthByPrecision: device px / scale.
        expect(thin).toBeCloseTo(getLineWidth(BorderStyleTypes.THIN) / scale)
        expect(medium).toBeCloseTo(getLineWidth(BorderStyleTypes.MEDIUM) / scale)
        expect(thick).toBeCloseTo(THICK_BORDER_DEVICE_PX / scale)
        expect(thick).toBeGreaterThan(medium)
        expect(medium).toBeGreaterThan(thin)
      }
    }
  })
})

describe('printBorderWidthPt', () => {
  it('prints thin / medium / thick at Excel weights', () => {
    expect(printBorderWidthPt(BorderStyleTypes.THIN)).toBe(0.75)
    expect(printBorderWidthPt(BorderStyleTypes.HAIR)).toBe(0.75)
    expect(printBorderWidthPt(BorderStyleTypes.MEDIUM)).toBe(1.5)
    expect(printBorderWidthPt(BorderStyleTypes.MEDIUM_DASH_DOT)).toBe(1.5)
    expect(printBorderWidthPt(BorderStyleTypes.THICK)).toBe(2.25)
    expect(printBorderWidthPt(undefined)).toBe(0.75)
  })
})

interface StrokeCall {
  readonly lineWidth: number
  readonly strokeStyle: string
  readonly from: readonly [number, number]
  readonly to: readonly [number, number]
}

function fakeContext(): {
  ctx: Record<string, unknown>
  strokes: StrokeCall[]
  translates: number[][]
} {
  const strokes: StrokeCall[] = []
  const translates: number[][] = []
  let lineWidth = 0
  let from: readonly [number, number] = [0, 0]
  let to: readonly [number, number] = [0, 0]
  const ctx: Record<string, unknown> = {
    strokeStyle: '',
    save: vi.fn(),
    restore: vi.fn(),
    beginPath: vi.fn(),
    closePath: vi.fn(),
    closePathByEnv: vi.fn(),
    setLineDash: vi.fn(),
    translateWithPrecisionRatio: (x: number, y: number) => {
      translates.push([x, y])
    },
    setLineWidthByPrecision: (value: number) => {
      lineWidth = value
    },
    moveToByPrecision: (x: number, y: number) => {
      from = [x, y]
    },
    lineToByPrecision: (x: number, y: number) => {
      to = [x, y]
    },
    stroke: () => {
      // Univer's diagonal helper strokes a zero-length path for straight
      // edges; only real segments matter here.
      if (from[0] === to[0] && from[1] === to[1]) return
      strokes.push({ lineWidth, strokeStyle: ctx.strokeStyle as string, from, to })
    },
  }
  return { ctx, strokes, translates }
}

function fakeRenderContext(ctx: Record<string, unknown>) {
  return {
    ctx,
    precisionScale: 1,
    overflowCache: undefined,
    diffRanges: [],
    viewRanges: [],
    spreadsheetSkeleton: {
      worksheet: { getRowVisible: () => true, getColVisible: () => true },
      getCellWithCoordByIndex: () => ({
        startX: 100,
        startY: 40,
        endX: 180,
        endY: 60,
        isMerged: false,
        isMergedMainCell: false,
        mergeInfo: { startRow: 2, endRow: 2, startColumn: 1, endColumn: 1 },
      }),
    },
  }
}

describe('installThickBorderFix', () => {
  beforeAll(() => {
    installThickBorderFix()
    installThickBorderFix() // idempotent
  })

  it('draws thick edges at the wider, pixel-aligned width and leaves the rest to Univer', () => {
    const { ctx, strokes, translates } = fakeContext()
    const border = new Border()
    border.renderBorderByCell(fakeRenderContext(ctx) as never, 2, 1, {
      t: { type: 't', style: BorderStyleTypes.THIN, color: '#111111' },
      b: { type: 'b', style: BorderStyleTypes.THICK, color: '#ff0000' },
      r: { type: 'r', style: BorderStyleTypes.MEDIUM, color: '#222222' },
    } as never)
    const thin = strokes.find((s) => s.strokeStyle === '#111111')
    const medium = strokes.find((s) => s.strokeStyle === '#222222')
    const thick = strokes.find((s) => s.strokeStyle === '#ff0000')
    expect(thin?.lineWidth).toBe(getLineWidth(BorderStyleTypes.THIN))
    expect(medium?.lineWidth).toBe(getLineWidth(BorderStyleTypes.MEDIUM))
    expect(thick?.lineWidth).toBe(THICK_BORDER_DEVICE_PX)
    // The thick bottom edge runs along the cell's bottom, extended by the
    // half-width buffer so corners close, and is drawn after the others so
    // it wins on shared edges.
    expect(thick?.from[1]).toBe(60)
    expect(thick?.to[1]).toBe(60)
    expect(strokes.indexOf(thick!)).toBeGreaterThan(strokes.indexOf(medium!))
    // Even width: the half-pixel blur offset is cancelled for the thick pass.
    expect(translates).toEqual([[-0.5, -0.5]])
  })

  it('does not touch cells without thick edges', () => {
    const { ctx, strokes, translates } = fakeContext()
    const border = new Border()
    border.renderBorderByCell(fakeRenderContext(ctx) as never, 2, 1, {
      l: { type: 'l', style: BorderStyleTypes.MEDIUM, color: '#333333' },
    } as never)
    expect(strokes).toHaveLength(1)
    expect(strokes[0]?.lineWidth).toBe(2)
    expect(translates).toEqual([])
  })

  it('skips hidden rows like the stock extension', () => {
    const { ctx, strokes } = fakeContext()
    const border = new Border()
    const context = fakeRenderContext(ctx)
    context.spreadsheetSkeleton.worksheet.getRowVisible = () => false
    const result = border.renderBorderByCell(context as never, 2, 1, {
      t: { type: 't', style: BorderStyleTypes.THICK, color: '#000000' },
    } as never)
    expect(result).toBe(true)
    expect(strokes).toHaveLength(0)
  })
})
