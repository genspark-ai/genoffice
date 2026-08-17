import { describe, expect, it } from 'vitest'
import {
  shapeBackgroundImage,
  shapePaths,
  shapeTextInsetsPx,
} from '../src/renderer/editor/shape-svg'
import { textboxBoxStyle } from '../src/renderer/editor/protected-render'

function decodedSvg(url: string | null): string {
  expect(url).toBeTruthy()
  return decodeURIComponent(url!.slice('url("data:image/svg+xml,'.length, -2))
}

describe('preset geometry coverage for docs shapes', () => {
  it.each(['triangle', 'star5', 'rightArrow', 'ellipse'])('%s has a path', (prst) => {
    const paths = shapePaths(prst, 100, 80)
    expect(paths).toBeTruthy()
    expect([paths!.main, paths!.fillOnly, paths!.strokeOnly].some(Boolean)).toBe(true)
  })
})

describe('straight connector rendering', () => {
  it('draws level at the vertical center without diag', () => {
    const svg = decodedSvg(shapeBackgroundImage('lineArrow', 200, 40, undefined, '112233'))
    expect(svg).toContain('M 0 19 L 198 19')
  })

  it('flipH reverses a level line so the arrowhead lands on the left tip', () => {
    const svg = decodedSvg(
      shapeBackgroundImage('lineArrow', 200, 40, undefined, '112233', { flipH: true }),
    )
    expect(svg).toContain('M 198 19 L 0 19')
  })

  it('diag runs corner to corner; flipV picks the rising diagonal', () => {
    const svg = decodedSvg(
      shapeBackgroundImage('lineArrow', 200, 40, undefined, '112233', {
        diag: true,
        flipV: true,
      }),
    )
    expect(svg).toContain('M 0 38 L 198 0')
  })
})

describe('shapeTextInsetsPx', () => {
  it('ellipse uses the inscribed rectangle', () => {
    const ins = shapeTextInsetsPx('ellipse', 200, 100)!
    expect(ins.l).toBeCloseTo(200 * ((1 - Math.SQRT1_2) / 2), 5)
    expect(ins.t).toBeCloseTo(100 * ((1 - Math.SQRT1_2) / 2), 5)
  })

  it('rightArrow keeps text out of the head and above/below the shaft', () => {
    expect(shapeTextInsetsPx('rightArrow', 240, 80)).toEqual({ l: 0, t: 20, r: 40, b: 20 })
  })

  it('plain rect has no geometry insets', () => {
    expect(shapeTextInsetsPx('rect', 200, 100)).toBeNull()
  })
})

describe('textboxBoxStyle geometry padding', () => {
  it('adds the ellipse text rect to the declared insets', () => {
    const style = textboxBoxStyle({
      paras: [],
      prst: 'ellipse',
      widthPx: 200,
      heightPx: 100,
      fill: '2E6E9E',
    })
    const k = (1 - Math.SQRT1_2) / 2
    const top = Math.round((4.8 + 100 * k) * 100) / 100
    const left = Math.round((9.6 + 200 * k) * 100) / 100
    expect(style).toContain(`padding:${top}px ${left}px ${top}px ${left}px`)
  })

  it('plain textboxes keep the default insets', () => {
    const style = textboxBoxStyle({ paras: [] })
    expect(style).toContain('padding:4.8px 9.6px 4.8px 9.6px')
  })
})
