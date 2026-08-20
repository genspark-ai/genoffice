import { describe, expect, it } from 'vitest'
import { detectVectorRegions, pageConfidence, PAGE_CONFIDENCE_MIN } from '../src/analyze'
import type { RawPath, RawSubpath } from '../src/ir'
import { mkText } from './helpers/chars'

const PAGE = { x0: 0, y0: 0, x1: 612, y1: 792 }

const curveSub = (x0: number, y0: number, x1: number, y1: number): RawSubpath => ({
  points: [
    { x: x0, y: y0 },
    { x: (x0 + x1) / 2, y: y1 },
    { x: x1, y: y1 },
  ],
  closed: false,
  hasCurves: true,
})

const strokedPath = (subpaths: RawSubpath[]): RawPath => ({
  subpaths,
  filled: false,
  stroked: true,
  fillColor: '000000',
  strokeColor: '000000',
  strokeWidth: 1,
})

/** a scatter of bezier arcs inside [x0,x1]×[y0,y1] */
function artCluster(x0: number, y0: number, x1: number, y1: number, n = 10): RawPath[] {
  const paths: RawPath[] = []
  for (let i = 0; i < n; i++) {
    const t = i / n
    paths.push(
      strokedPath([
        curveSub(x0 + t * (x1 - x0) * 0.5, y0 + t * (y1 - y0) * 0.5, x1 - t * 10, y1 - t * 5),
      ]),
    )
  }
  return paths
}

describe('detectVectorRegions', () => {
  it('clusters dense curve paths over a text-sparse area into one region', () => {
    const regions = detectVectorRegions(artCluster(100, 400, 300, 550), [], PAGE)
    expect(regions).toHaveLength(1)
    const r = regions[0]!
    expect(r.x0).toBeLessThanOrEqual(100)
    expect(r.x1).toBeGreaterThanOrEqual(295)
    expect(r.y0).toBeLessThanOrEqual(400)
    expect(r.y1).toBeGreaterThanOrEqual(545)
  })

  it('ignores small stray decorations', () => {
    const regions = detectVectorRegions(artCluster(100, 400, 300, 550, 3), [], PAGE)
    expect(regions).toHaveLength(0)
  })

  it('leaves text-dense areas alone (decorated paragraph, not an illustration)', () => {
    const chars = []
    for (let row = 0; row < 12; row++) {
      chars.push(
        ...mkText('dense running paragraph text filling the whole area here', 100, {
          y: 540 - row * 12,
          fontSize: 10,
        }).chars,
      )
    }
    const regions = detectVectorRegions(artCluster(100, 400, 300, 550), chars, PAGE)
    expect(regions).toHaveLength(0)
  })

  it('vetoes a candidate containing paragraphs of body-size text (sparse card layout)', () => {
    // 3 full-width body lines inside the art area, but with enough padding
    // that char ink stays below the density gate — the line veto must catch it
    const chars = []
    for (let row = 0; row < 3; row++) {
      chars.push(
        ...mkText('running body text here', 110, { y: 520 - row * 40, fontSize: 10 }).chars,
      )
    }
    expect(detectVectorRegions(artCluster(100, 400, 300, 550), chars, PAGE)).toHaveLength(0)
  })

  it('keeps rasterizing charts whose only text is short tick labels', () => {
    const chars = [
      ...mkText('10%', 95, { y: 420, fontSize: 6 }).chars,
      ...mkText('20%', 95, { y: 470, fontSize: 6 }).chars,
      ...mkText('30%', 95, { y: 520, fontSize: 6 }).chars,
    ]
    expect(detectVectorRegions(artCluster(100, 400, 300, 550), chars, PAGE)).toHaveLength(1)
  })

  it('falls back to the bare art cluster when absorbed furniture swallows a text area', () => {
    const art = artCluster(100, 400, 300, 550)
    // a frame whose center sits inside the cluster but which reaches down over
    // a paragraph zone well below the art
    const frame = strokedPath([
      {
        points: [
          { x: 90, y: 560 },
          { x: 90, y: 250 },
          { x: 320, y: 250 },
        ],
        closed: false,
        hasCurves: false,
      },
    ])
    const chars = []
    for (let row = 0; row < 3; row++) {
      chars.push(
        ...mkText('running body text here', 110, { y: 330 - row * 14, fontSize: 10 }).chars,
      )
    }
    const regions = detectVectorRegions([...art, frame], chars, PAGE)
    expect(regions).toHaveLength(1)
    // the admitted region is the bare cluster: it stops above the paragraphs
    expect(regions[0]!.y0).toBeGreaterThan(350)
  })

  it("absorbs the illustration's axis-aligned frame into the region", () => {
    const art = artCluster(100, 400, 300, 550)
    // chart axes: an L-shaped axis-aligned polyline anchored inside the cluster
    const axes = strokedPath([
      {
        points: [
          { x: 90, y: 560 },
          { x: 90, y: 390 },
          { x: 320, y: 390 },
        ],
        closed: false,
        hasCurves: false,
      },
    ])
    const regions = detectVectorRegions([...art, axes], [], PAGE)
    expect(regions).toHaveLength(1)
    expect(regions[0]!.x0).toBeLessThanOrEqual(90)
    expect(regions[0]!.x1).toBeGreaterThanOrEqual(320)
    expect(regions[0]!.y0).toBeLessThanOrEqual(390)
  })
})

describe('pageConfidence', () => {
  it('is 1.0 for a clean page', () => {
    expect(
      pageConfidence({ badUnicodeRatio: 0, streamTableConfidences: [], warningCount: 0 }),
    ).toBe(1)
  })

  it('penalizes bad ToUnicode ratio, weak stream tables and warnings', () => {
    expect(
      pageConfidence({ badUnicodeRatio: 0.1, streamTableConfidences: [], warningCount: 0 }),
    ).toBeCloseTo(0.7)
    expect(
      pageConfidence({ badUnicodeRatio: 0, streamTableConfidences: [0.5], warningCount: 0 }),
    ).toBeCloseTo(0.75)
    expect(
      pageConfidence({ badUnicodeRatio: 0, streamTableConfidences: [], warningCount: 2 }),
    ).toBeCloseTo(0.88)
  })

  it('caps each penalty and floors at 0', () => {
    expect(
      pageConfidence({ badUnicodeRatio: 1, streamTableConfidences: [], warningCount: 100 }),
    ).toBeCloseTo(1 - 0.45 - 0.3)
  })

  it('combined weak signals cross the downgrade threshold', () => {
    const conf = pageConfidence({
      badUnicodeRatio: 0.12,
      streamTableConfidences: [0.55],
      warningCount: 3,
    })
    expect(conf).toBeLessThan(PAGE_CONFIDENCE_MIN)
  })
})
