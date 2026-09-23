import { describe, expect, it } from 'vitest'
import type { ExtractedPage } from '../src/extract'
import { authoredInkBoxes } from '../src/pipeline'

function inkPage(paths: unknown[]): ExtractedPage {
  return {
    index: 0,
    widthPt: 1000,
    heightPt: 1000,
    rotation: 0,
    chars: [],
    images: [],
    paths,
    degraded: false,
    scanned: false,
    hasStructTree: false,
    vectorRegions: [],
    badUnicodeRatio: 0,
  } as unknown as ExtractedPage
}

describe('authoredInkBoxes path bounds', () => {
  it('survives a 130k-point subpath (Math.min(...points) throws past ~125k arguments)', () => {
    const page = inkPage([
      {
        filled: true,
        fillAlpha: 255,
        fillColor: '000000',
        subpaths: [{ points: Array.from({ length: 130_000 }, (_, i) => ({ x: i, y: i % 97 })) }],
      },
    ])
    const boxes = authoredInkBoxes(page)
    expect(boxes).toHaveLength(1)
    expect(boxes[0]).toEqual({ x0: 0, y0: 0, x1: 129_999, y1: 96 })
  })

  it('still skips subpaths under three points and unfilled paths', () => {
    const page = inkPage([
      {
        filled: true,
        fillAlpha: 255,
        fillColor: '000000',
        subpaths: [{ points: [{ x: 0, y: 0 }, { x: 1, y: 1 }] }],
      },
      {
        filled: false,
        fillAlpha: 255,
        fillColor: '000000',
        subpaths: [{ points: [{ x: 0, y: 0 }, { x: 1, y: 1 }, { x: 2, y: 2 }] }],
      },
    ])
    expect(authoredInkBoxes(page)).toEqual([])
  })
})
