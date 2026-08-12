import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  MARKUP_COLORS,
  geomDispSize,
  pdfRectToCss,
  pdfToView,
  quadSetsMatch,
  quadToRect,
  selectionQuadsByPage,
  viewToPdf,
  type PageGeom,
} from '../src/renderer/annotations'

const geom = (rot: number, pw = 100, ph = 200): PageGeom => ({ pw, ph, rot })

describe('geomDispSize', () => {
  it('keeps width/height at 0 and 180 degrees', () => {
    expect(geomDispSize(geom(0))).toEqual({ width: 100, height: 200 })
    expect(geomDispSize(geom(180))).toEqual({ width: 100, height: 200 })
  })

  it('swaps width/height at 90 and 270 degrees', () => {
    expect(geomDispSize(geom(90))).toEqual({ width: 200, height: 100 })
    expect(geomDispSize(geom(270))).toEqual({ width: 200, height: 100 })
  })

  it('normalizes negative and >360 rotations', () => {
    expect(geomDispSize(geom(-90))).toEqual({ width: 200, height: 100 })
    expect(geomDispSize(geom(450))).toEqual({ width: 200, height: 100 })
    expect(geomDispSize(geom(360))).toEqual({ width: 100, height: 200 })
  })
})

describe('viewToPdf / pdfToView', () => {
  it('maps top-left display origin to PDF space at rot 0', () => {
    expect(viewToPdf(geom(0), 0, 0)).toEqual([0, 200])
    expect(viewToPdf(geom(0), 10, 30)).toEqual([10, 170])
  })

  it('handles each rotation quadrant', () => {
    expect(viewToPdf(geom(90), 10, 30)).toEqual([30, 10])
    expect(viewToPdf(geom(180), 10, 30)).toEqual([90, 30])
    expect(viewToPdf(geom(270), 10, 30)).toEqual([70, 190])
  })

  it('pdfToView is the inverse of viewToPdf for every rotation', () => {
    for (const rot of [0, 90, 180, 270]) {
      const g = geom(rot)
      const [px, py] = viewToPdf(g, 12, 34)
      expect(pdfToView(g, px, py)).toEqual([12, 34])
    }
  })
})

describe('pdfRectToCss', () => {
  it('converts a PDF rect to a scaled CSS box at rot 0', () => {
    // PDF rect [10, 20, 40, 60] on a 100x200 page: top = 200 - 60 = 140
    expect(pdfRectToCss(geom(0), [10, 20, 40, 60], 2)).toEqual({
      left: 20,
      top: 280,
      width: 60,
      height: 80,
    })
  })

  it('produces a non-negative box regardless of corner order', () => {
    const box = pdfRectToCss(geom(90), [40, 60, 10, 20], 1)
    expect(box.width).toBeGreaterThan(0)
    expect(box.height).toBeGreaterThan(0)
  })
})

describe('quadToRect', () => {
  it('returns the bounding rect of corners in any order', () => {
    expect(quadToRect([5, 9, 1, 2, 8, 3, 4, 7])).toEqual([1, 2, 8, 9])
  })
})

describe('quadSetsMatch', () => {
  const q1 = [100, 716, 200, 716, 100, 698, 200, 698]
  const q2 = [100, 666, 200, 666, 100, 648, 200, 648]

  it('matches identical quad sets', () => {
    expect(quadSetsMatch([q1, q2], [q1, q2])).toBe(true)
  })

  it('is order-insensitive', () => {
    expect(quadSetsMatch([q1, q2], [q2, q1])).toBe(true)
  })

  it('absorbs small coordinate drift (zoom rounding / float32 round-trip)', () => {
    const drifted = q1.map((v) => v + 1.5)
    expect(quadSetsMatch([q1], [drifted])).toBe(true)
  })

  it('rejects offsets beyond the tolerance', () => {
    const shifted = q1.map((v, i) => (i % 2 === 1 ? v - 14 : v)) // one text line lower
    expect(quadSetsMatch([q1], [shifted])).toBe(false)
  })

  it('rejects differing quad counts (subset selections add, not toggle)', () => {
    expect(quadSetsMatch([q1], [q1, q2])).toBe(false)
  })

  it('does not reuse a quad for two matches', () => {
    expect(quadSetsMatch([q1, q1], [q1, q2])).toBe(false)
  })
})

describe('MARKUP_COLORS', () => {
  it('defines a normalized rgb triple for every markup type', () => {
    for (const type of ['highlight', 'underline', 'strikeout'] as const) {
      const c = MARKUP_COLORS[type]
      expect(c).toHaveLength(3)
      for (const ch of c) {
        expect(ch).toBeGreaterThanOrEqual(0)
        expect(ch).toBeLessThanOrEqual(1)
      }
    }
  })
})

describe('selectionQuadsByPage', () => {
  afterEach(() => {
    vi.restoreAllMocks()
    document.body.innerHTML = ''
  })

  const domRect = (left: number, top: number, right: number, bottom: number) =>
    ({
      left,
      top,
      right,
      bottom,
      width: right - left,
      height: bottom - top,
    }) as DOMRect

  function setupPage(): { scrollEl: HTMLElement; pageEl: HTMLElement } {
    const scrollEl = document.createElement('div')
    const pageEl = document.createElement('div')
    pageEl.className = 'pdf-page'
    scrollEl.appendChild(pageEl)
    document.body.appendChild(scrollEl)
    vi.spyOn(pageEl, 'getBoundingClientRect').mockReturnValue(domRect(0, 0, 100, 200))
    return { scrollEl, pageEl }
  }

  function mockSelection(scrollEl: HTMLElement, clientRects: DOMRect[], collapsed = false) {
    const anchor = document.createTextNode('x')
    scrollEl.appendChild(anchor)
    const range = {
      commonAncestorContainer: anchor,
      getClientRects: () => clientRects,
    } as unknown as Range
    vi.spyOn(window, 'getSelection').mockReturnValue({
      isCollapsed: collapsed,
      rangeCount: 1,
      getRangeAt: () => range,
    } as unknown as Selection)
  }

  it('returns null when there is no selection or it is collapsed', () => {
    const { scrollEl } = setupPage()
    vi.spyOn(window, 'getSelection').mockReturnValue(null)
    expect(selectionQuadsByPage(scrollEl, [geom(0)], 1)).toBeNull()

    mockSelection(scrollEl, [domRect(10, 20, 50, 30)], true)
    expect(selectionQuadsByPage(scrollEl, [geom(0)], 1)).toBeNull()
  })

  it('returns null when the selection is outside the scroll container', () => {
    const { scrollEl } = setupPage()
    const outside = document.createElement('div')
    document.body.appendChild(outside)
    const anchor = document.createTextNode('x')
    outside.appendChild(anchor)
    vi.spyOn(window, 'getSelection').mockReturnValue({
      isCollapsed: false,
      rangeCount: 1,
      getRangeAt: () => ({ commonAncestorContainer: anchor, getClientRects: () => [] }),
    } as unknown as Selection)
    expect(selectionQuadsByPage(scrollEl, [geom(0)], 1)).toBeNull()
  })

  it('maps a client rect to a PDF-space quad on the containing page', () => {
    const { scrollEl } = setupPage()
    mockSelection(scrollEl, [domRect(10, 20, 50, 30)])
    const result = selectionQuadsByPage(scrollEl, [geom(0)], 1)
    expect(result).not.toBeNull()
    // rot 0, page 100x200: (10,20)->(10,180), (50,30)->(50,170)
    expect(result!.get(0)).toEqual([[10, 180, 50, 180, 10, 170, 50, 170]])
  })

  it('drops rects that fully contain other rects and dedupes identical boxes', () => {
    const { scrollEl } = setupPage()
    mockSelection(scrollEl, [
      domRect(5, 15, 60, 40), // big box containing the two below -> dropped
      domRect(10, 20, 50, 30),
      domRect(10, 20, 50, 30), // duplicate -> deduped
    ])
    const result = selectionQuadsByPage(scrollEl, [geom(0)], 1)
    expect(result!.get(0)).toHaveLength(1)
  })

  it('ignores rects whose center falls outside every page', () => {
    const { scrollEl } = setupPage()
    mockSelection(scrollEl, [domRect(500, 500, 540, 510)])
    expect(selectionQuadsByPage(scrollEl, [geom(0)], 1)).toBeNull()
  })

  it('divides coordinates by the display scale', () => {
    const { scrollEl, pageEl } = setupPage()
    vi.spyOn(pageEl, 'getBoundingClientRect').mockReturnValue(domRect(0, 0, 200, 400))
    mockSelection(scrollEl, [domRect(20, 40, 100, 60)])
    const result = selectionQuadsByPage(scrollEl, [geom(0)], 2)
    expect(result!.get(0)).toEqual([[10, 180, 50, 180, 10, 170, 50, 170]])
  })
})
