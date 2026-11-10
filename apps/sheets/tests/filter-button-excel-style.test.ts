import { describe, expect, it } from 'vitest'

import {
  drawExcelFilterButton,
  filterButtonColors,
  patchFilterButtonShapes,
} from '../src/renderer/filter-button-excel-style'

function recordingContext(): { ctx: CanvasRenderingContext2D; calls: string[] } {
  const calls: string[] = []
  const handler: ProxyHandler<object> = {
    get: (_target, name) => {
      if (typeof name !== 'string') return undefined
      return (...args: unknown[]) => {
        calls.push(`${name}(${args.map(String).join(',')})`)
      }
    },
    set: (_target, name, value) => {
      calls.push(`${String(name)}=${String(value)}`)
      return true
    },
  }
  return { ctx: new Proxy({}, handler) as CanvasRenderingContext2D, calls }
}

const theme = { getColorFromTheme: (key: string) => `theme(${key})` }

describe('drawExcelFilterButton', () => {
  it('draws a bordered box with one arrow when the column has no criteria', () => {
    const { ctx, calls } = recordingContext()
    drawExcelFilterButton(
      ctx,
      16,
      { hasCriteria: false, hovered: false },
      filterButtonColors(theme, false),
    )
    expect(calls.filter((c) => c.startsWith('roundRect'))).toHaveLength(1)
    expect(calls).toContain('strokeStyle=theme(gray.300)')
    expect(calls.filter((c) => c === 'closePath()')).toHaveLength(1)
  })

  it('adds the funnel beside a smaller arrow once criteria apply', () => {
    const { ctx, calls } = recordingContext()
    drawExcelFilterButton(
      ctx,
      16,
      { hasCriteria: true, hovered: true },
      filterButtonColors(theme, true),
    )
    expect(calls.filter((c) => c === 'closePath()')).toHaveLength(2)
    expect(calls).toContain('strokeStyle=theme(primary.600)')
    expect(calls).toContain('fillStyle=theme(gray.50)')
  })
})

describe('patchFilterButtonShapes', () => {
  it('replaces _draw on the shared prototype and ignores other scene objects', () => {
    class FakeButton {
      _cellWidth = 80
      _cellHeight = 20
      _hovered = false
      _filterParams = { hasCriteria: false }
      _themeService = theme
      _draw(_ctx: CanvasRenderingContext2D): void {
        throw new Error('stock draw must not run')
      }
    }
    const button = new FakeButton()
    expect(patchFilterButtonShapes([{ _draw: () => undefined }, button])).toBe(true)
    const { ctx, calls } = recordingContext()
    globalThis.Path2D ??= class {
      rect(): void {}
    } as unknown as typeof Path2D
    new FakeButton()._draw(ctx)
    expect(calls.some((c) => c.startsWith('clip('))).toBe(true)
    expect(calls.filter((c) => c.startsWith('roundRect'))).toHaveLength(1)
    expect(patchFilterButtonShapes([])).toBe(false)
  })
})
