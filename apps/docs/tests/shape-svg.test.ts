import { describe, expect, it } from 'vitest'
import {
  custGeomBackgroundCss,
  shapeBackgroundImage,
  shapePaths,
  shapeTextInsetsPx,
} from '../src/renderer/editor/shape-svg'
import { textboxBoxStyle } from '../src/renderer/editor/protected-render'

function decodedSvg(url: string | null): string {
  expect(url).toBeTruthy()
  return decodeURIComponent(url!.slice('url("data:image/svg+xml,'.length, -2))
}

function decodedCss(css: string): string {
  const match = /url\("data:image\/svg\+xml,([^"]*)"\)/.exec(css)
  expect(match).toBeTruthy()
  return decodeURIComponent(match![1]!)
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

describe('authored outline width', () => {
  it('keeps the historical encoded SVG byte-identical when no width is authored', () => {
    expect(decodedSvg(shapeBackgroundImage('rect', 10, 10, '112233', '445566'))).toBe(
      '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="-1 -1 10 10">' +
        '<path d="M 0 0 L 8 0 L 8 8 L 0 8 Z" fill="#112233" stroke="#445566" stroke-width="2"/>' +
        '</svg>',
    )
  })

  it('uses the authored width for the main path instead of the 2px default', () => {
    const svg = decodedSvg(shapeBackgroundImage('rect', 100, 80, '4472C4', '2F5496', undefined, 3))
    expect(svg).toContain('stroke-width="3"')
    expect(svg).not.toContain('stroke-width="2"')
  })

  it('keeps the historical 2 / 2.5 defaults when no width is authored', () => {
    expect(decodedSvg(shapeBackgroundImage('rect', 100, 80, '4472C4', '2F5496'))).toContain(
      'stroke-width="2"',
    )
    expect(decodedSvg(shapeBackgroundImage('lineArrow', 200, 40, undefined, '112233'))).toContain(
      'stroke-width="2.5"',
    )
  })

  it('applies the authored width to stroke-only line paths', () => {
    const svg = decodedSvg(
      shapeBackgroundImage('lineArrow', 200, 40, undefined, '112233', undefined, 3),
    )
    expect(svg).toContain('stroke-width="3"')
  })

  it('clamps a hairline to at least 0.5px', () => {
    const svg = decodedSvg(
      shapeBackgroundImage('rect', 100, 80, '4472C4', '2F5496', undefined, 0.1),
    )
    expect(svg).toContain('stroke-width="0.5"')
  })

  it('grows the viewBox around a thick stroke so it is not clipped', () => {
    const thin = decodedSvg(shapeBackgroundImage('rect', 100, 80, '4472C4', '2F5496'))
    expect(thin).toContain('viewBox="-1 -1 100 80"')
    // stroke 8 -> inset 4.5 and the viewport grows by 2*(4.5-1) each way, so the
    // path (0..98, stroke reaching 102 / -4) stays inside -4.5..102.5
    const thick = decodedSvg(
      shapeBackgroundImage('rect', 100, 80, '4472C4', '2F5496', undefined, 8),
    )
    expect(thick).toContain('viewBox="-4.5 -4.5 107 87"')
  })

  it('threads the width through the custGeom background helper', () => {
    const css = custGeomBackgroundCss(
      { path: 'M 0 0 L 1 0 L 1 1 L 0 1 Z' },
      100,
      80,
      '4472C4',
      '2F5496',
      4,
    )
    expect(decodedCss(css)).toContain('stroke-width="4"')
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

  it('draws a preset shape outline at its authored a:ln width', () => {
    const style = textboxBoxStyle({
      paras: [],
      prst: 'ellipse',
      widthPx: 200,
      heightPx: 100,
      fill: '2E6E9E',
      borderColor: '2F5496',
      borderWidthPx: 3,
    })
    expect(decodedCss(style)).toContain('stroke-width="3"')
  })

  it('draws a custGeom shape outline at its authored a:ln width', () => {
    const style = textboxBoxStyle({
      paras: [],
      pathData: { path: 'M 0 0 L 1 0 L 1 1 L 0 1 Z' },
      widthPx: 100,
      heightPx: 80,
      fill: '2E6E9E',
      borderColor: '2F5496',
      borderWidthPx: 2,
    })
    expect(decodedCss(style)).toContain('stroke-width="2"')
  })
})
