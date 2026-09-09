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

  function setupPage(pageRect = domRect(0, 0, 100, 200)): {
    scrollEl: HTMLElement
    pageEl: HTMLElement
  } {
    const scrollEl = document.createElement('div')
    const pageEl = document.createElement('div')
    pageEl.className = 'pdf-page'
    scrollEl.appendChild(pageEl)
    document.body.appendChild(scrollEl)
    vi.spyOn(pageEl, 'getBoundingClientRect').mockReturnValue(pageRect)
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

  it('normalizes mixed-height fragments on one visual line and bridges a word-sized gap', () => {
    const { scrollEl } = setupPage()
    mockSelection(scrollEl, [domRect(10, 20, 30, 32), domRect(35, 16, 80, 30)])

    expect(selectionQuadsByPage(scrollEl, [geom(0)], 1)!.get(0)).toEqual([
      [10, 184, 80, 184, 10, 168, 80, 168],
    ])
  })

  it('normalizes each visual line independently in a multiline selection', () => {
    const { scrollEl } = setupPage()
    mockSelection(scrollEl, [
      domRect(10, 20, 30, 32),
      domRect(35, 16, 80, 30),
      domRect(10, 50, 30, 62),
      domRect(35, 46, 80, 60),
    ])

    expect(selectionQuadsByPage(scrollEl, [geom(0)], 1)!.get(0)).toEqual([
      [10, 184, 80, 184, 10, 168, 80, 168],
      [10, 154, 80, 154, 10, 138, 80, 138],
    ])
  })

  it('keeps wide-gap fragments at their own vertical bounds', () => {
    const { scrollEl } = setupPage()
    mockSelection(scrollEl, [domRect(10, 20, 30, 32), domRect(70, 16, 90, 30)])

    expect(selectionQuadsByPage(scrollEl, [geom(0)], 1)!.get(0)).toEqual([
      [10, 180, 30, 180, 10, 168, 30, 168],
      [70, 184, 90, 184, 70, 170, 90, 170],
    ])
  })

  it('does not bridge a third line through an overlapping short fragment', () => {
    const { scrollEl } = setupPage()
    mockSelection(scrollEl, [
      domRect(10, 10, 30, 20),
      domRect(35, 15, 55, 25),
      domRect(60, 20, 80, 30),
    ])

    expect(selectionQuadsByPage(scrollEl, [geom(0)], 1)!.get(0)).toEqual([
      [10, 190, 55, 190, 10, 175, 55, 175],
      [60, 180, 80, 180, 60, 170, 80, 170],
    ])
  })

  it('keeps a fragment that equally overlaps two line bands independent', () => {
    const { scrollEl } = setupPage()
    mockSelection(scrollEl, [
      domRect(10, 10, 30, 20),
      domRect(60, 30, 80, 40),
      domRect(35, 15, 55, 35),
    ])

    expect(selectionQuadsByPage(scrollEl, [geom(0)], 1)!.get(0)).toEqual([
      [10, 190, 30, 190, 10, 180, 30, 180],
      [60, 170, 80, 170, 60, 160, 80, 160],
      [35, 185, 55, 185, 35, 165, 55, 165],
    ])
  })

  it('does not let progressively eroded overlaps bridge every line into one group', () => {
    const { scrollEl } = setupPage()
    mockSelection(scrollEl, [
      domRect(10, 10, 20, 20),
      domRect(25, 15, 35, 25),
      domRect(40, 17.5, 50, 27.5),
      domRect(55, 18.75, 65, 28.75),
    ])

    expect(selectionQuadsByPage(scrollEl, [geom(0)], 1)!.get(0)).toEqual([
      [10, 190, 35, 190, 10, 175, 35, 175],
      [40, 182.5, 65, 182.5, 40, 171.25, 65, 171.25],
    ])
  })

  it.each([
    [90, [10, 32, 30, 32, 10, 20, 30, 20], [35, 30, 80, 30, 35, 16, 80, 16]],
    [270, [70, 180, 90, 180, 70, 168, 90, 168], [20, 184, 65, 184, 20, 170, 65, 170]],
  ])('keeps sideways page rectangles unnormalized at %i degrees', (rot, first, second) => {
    const { scrollEl } = setupPage(domRect(0, 0, 200, 100))
    mockSelection(scrollEl, [domRect(20, 10, 32, 30), domRect(16, 35, 30, 80)])

    expect(selectionQuadsByPage(scrollEl, [geom(rot)], 1)!.get(0)).toEqual([first, second])
  })
})
