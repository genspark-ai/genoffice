import { describe, it, expect } from 'vitest'
import { HeuristicMetrics, type RunStyle } from '../src/metrics'
import { layoutText } from '../src/text-layout'
import { makeViewport } from '../src/coords'
import type { TextBody } from '@genoffice/pptx-engine'

const vp = makeViewport({ cx: 9525 * 1000, cy: 9525 * 1000 }, 1000)

// A metrics provider that reports the HG family as missing (same-script substitute in use)
class SubstitutingMetrics extends HeuristicMetrics {
  substituted(style: RunStyle): boolean {
    return /^HG/.test(style.fontFamily)
  }
}

const body = (fontFamily: string, bold?: boolean): TextBody => ({
  anchor: 'top',
  insets: { l: 0, t: 0, r: 0, b: 0 },
  autofit: 'none',
  wrap: true,
  paragraphs: [{ runs: [{ text: 'Sample', fontSize: 18, fontFamily, ...(bold ? { bold } : {}) }] }],
})
const firstRun = (fontFamily: string, bold?: boolean) =>
  layoutText({
    body: body(fontFamily, bold),
    boxWidthPx: 400,
    boxHeightPx: 100,
    metrics: new SubstitutingMetrics(),
    vp,
  }).lines[0]!.runs[0]!

describe('weight implied by the family name', () => {
  it('a missing ultra-bold HG family draws bold', () => {
    expect(firstRun('HGS\u5275\u82f1\u89d2\uff7a\uff9e\uff7c\uff6f\uff78UB').bold).toBe(true)
  })
  it('a present family is left alone even when the name says heavy', () => {
    expect(firstRun('Arial Black').bold).toBe(false)
  })
  it('a light family with b=1 renders at regular weight (PowerPoint: synthetic bold on the light face)', () => {
    expect(firstRun('Calibri Light', true).bold).toBe(false)
    expect(firstRun('Calibri', true).bold).toBe(true)
  })
})
