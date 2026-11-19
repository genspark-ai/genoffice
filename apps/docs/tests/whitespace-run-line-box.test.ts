/**
 * Word never sizes a line by a space / tab run and never by the paragraph mark
 * while ink is on the line; a no-break space run is ink (probe 2026-11-08).
 */
import { describe, expect, it } from 'vitest'
import {
  latinParaFactor,
  layoutParaOfRuns,
  paraRunFontFamily,
  strutHalfPoints,
} from '../src/renderer/editor/para-layout-spec'
import { HeuristicMetrics, simulateLines } from '../src/renderer/line-metrics'
import { inlineRunsOf, inlineSegments } from '../src/renderer/text-layout/inline-segments'

const TNR = 'Times New Roman'
const times = (text: string) => ({ text, fontAscii: TNR, sizeHalfPoints: 24 })
const para = (runs: object[]) =>
  layoutParaOfRuns(runs as never, runs.map((r) => (r as { text: string }).text).join(''), null)
const LATIN = 'var(--doc-line-factor-latin,1.2)'

describe('whitespace runs beside ink', () => {
  it('a space or tab run in another face / size leaves the Times factor, face and strut alone', () => {
    for (const ws of [
      { text: ' ', sizeHalfPoints: 22 },
      { text: ' ', fontAscii: 'Calibri', sizeHalfPoints: 28 },
      { text: '\t', fontAscii: 'Calibri', sizeHalfPoints: 28 },
      { text: ' ' },
    ]) {
      const p = para([times('The night shift'), ws, times('was noisy')])
      expect(latinParaFactor(p, LATIN)).toBe('1.1499')
      expect(paraRunFontFamily(p)).toContain(TNR)
      expect(strutHalfPoints(p)).toEqual({ halfPoints: 24, mixed: false })
    }
  })

  it('a no-break space run is ink and counts at its own face and size', () => {
    const p = para([
      times('a'),
      { text: ' ', fontAscii: 'Calibri', sizeHalfPoints: 28 },
      times('b'),
    ])
    expect(latinParaFactor(p, LATIN)).not.toBe('1.15')
    expect(strutHalfPoints(p)).toEqual({ halfPoints: 24, mixed: false })
  })

  it('text-layout tags space / tab runs beside ink, never U+00A0 or a space-only paragraph', () => {
    const segs = (runs: Array<{ text: string }>) =>
      inlineSegments({
        runs: inlineRunsOf(runs).runs,
        autospace: null,
        inheritedAscii: null,
        inheritedEa: null,
        wordWrap: true,
        balancedSpaces: false,
        justified: false,
        simsunGapLineHeight: null,
        lineMultiple: null,
      }).segments
    expect(segs([{ text: 'ab' }, { text: '  ' }, { text: '\t' }, { text: 'c' }])).toEqual([
      { from: 2, to: 4, cls: 'doc-ws-run' },
      { from: 4, to: 5, cls: 'doc-ws-run' },
    ])
    expect(segs([{ text: 'ab' }, { text: ' ' }])).toEqual([])
    expect(segs([{ text: ' ' }, { text: '  ' }])).toEqual([])
  })

  it('a space run the balanced-space rule already wraps keeps that span only', () => {
    const segs = inlineSegments({
      runs: inlineRunsOf([{ text: '\u4e2d\u6587' }, { text: ' ' }, { text: 'ab' }]).runs,
      autospace: null,
      inheritedAscii: null,
      inheritedEa: null,
      wordWrap: true,
      balancedSpaces: true,
      justified: false,
      simsunGapLineHeight: null,
      lineMultiple: null,
    }).segments
    expect(segs.map((s) => s.cls)).toEqual(['doc-balanced-space'])
  })

  it('the pagination simulation does not raise a line by a space or tab run', () => {
    const metrics = new HeuristicMetrics()
    const base = simulateLines([{ text: 'ab cd', sizeHalfPoints: 24 }], 500, metrics, 11, 'Calibri')
    for (const ws of ['  ', '\t']) {
      const runs = [
        { text: 'ab', sizeHalfPoints: 24 },
        { text: ws, sizeHalfPoints: 56 },
        { text: 'cd', sizeHalfPoints: 24 },
      ]
      const lines = simulateLines(runs, 500, metrics, 11, 'Calibri')
      expect(lines).toHaveLength(1)
      expect(lines[0].naturalLineH).toBeCloseTo(base[0].naturalLineH, 3)
    }
    // a tab-only paragraph is content: its own run sizes the line (strutHalfPoints agrees)
    const tabOnly = simulateLines([{ text: '\t', sizeHalfPoints: 56 }], 500, metrics, 11, 'Calibri')
    expect(tabOnly[0].naturalLineH).toBeGreaterThan(base[0].naturalLineH * 2)
  })
})
