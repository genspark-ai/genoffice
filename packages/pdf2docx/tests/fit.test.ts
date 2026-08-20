/** Line-width preflight unit tests (P21 A): injected measurer, no system fonts. */
import { describe, expect, it } from 'vitest'
import { preflightFitBlock } from '../src/rebuild/fit'
import type { SpanMeasurer } from '../src/rebuild/fit'
import type { Line, Span, TextBlock } from '../src/ir'

/** fixed advance per character (pt at fontSize 10), scaled linearly with size */
const perChar =
  (widthPt: number): SpanMeasurer =>
  (text, _family, sizePt) =>
    [...text].length * widthPt * (sizePt / 10)

const span = (text: string, over: Partial<Span> = {}): Span => ({
  text,
  box: { x0: 0, y0: 0, x1: 100, y1: 10 },
  fontSize: 10,
  fontFamily: 'Helvetica',
  bold: false,
  italic: false,
  color: '000000',
  dir: 'ltr',
  script: 'latin',
  ...over,
})

const lineOf = (spans: Span[]): Line => ({
  spans,
  box: { x0: 0, y0: 0, x1: 100, y1: 10 },
  baseline: 8,
  endsWithHyphen: false,
})

const block = (lines: Line[], over: Partial<TextBlock> = {}): TextBlock => ({
  kind: 'text',
  lines,
  box: { x0: 0, y0: 0, x1: 100, y1: 10 * lines.length },
  align: 'left',
  firstLineIndentPt: 0,
  dir: 'ltr',
  ...over,
})

describe('preflightFitBlock', () => {
  it('leaves fitting blocks untouched', () => {
    const s = span('aaaaaaaaaa') // 10 chars × 5pt = 50pt, avail 100
    preflightFitBlock(block([lineOf([s])]), 100, perChar(5))
    expect(s.charSpacingPt).toBeUndefined()
    expect(s.fontSize).toBe(10)
  })

  it('tightens overflow with negative character spacing', () => {
    const s = span('aaaaaaaaaa') // 10 × 10.4 = 104pt vs avail 100 (past the wrap gate)
    preflightFitBlock(block([lineOf([s])]), 100, perChar(10.4))
    expect(s.charSpacingPt).toBeLessThan(0)
    // (99.8 − 104) / 10 = −0.42, snapped DOWN to the twip grid → −0.45
    expect(s.charSpacingPt!).toBeCloseTo(-0.45, 5)
    expect(s.fontSize).toBe(10) // spacing alone covers it — no shrink
  })

  it('steps the font size down when spacing alone cannot cover the overflow', () => {
    const s = span('aaaaaaaaaa') // 10 × 12 = 120pt vs avail 100: needs -2pt/char ≫ 0.05em cap
    preflightFitBlock(block([lineOf([s])]), 100, perChar(12))
    expect(s.fontSize).toBeLessThan(10)
    expect(s.fontSize).toBeGreaterThanOrEqual(8.5) // ≤3 half-point steps
    expect(s.charSpacingPt ?? 0).toBeGreaterThanOrEqual(-0.5 - 1e-9) // capped at 0.05em
  })

  it('takes the worst line and charges the whole block uniformly', () => {
    const a = span('aaaaa') // 5 × 10.4 = 52 fits
    const b = span('bbbbbbbbbb') // 104 overflows
    preflightFitBlock(block([lineOf([a]), lineOf([b])]), 100, perChar(10.4))
    expect(a.charSpacingPt).toBe(b.charSpacingPt)
    expect(b.charSpacingPt).toBeLessThan(0)
  })

  it('ignores trailing spaces on a line (they hang, they never wrap)', () => {
    const s = span('aaaaaaaaaa          ')
    preflightFitBlock(block([lineOf([s])]), 100, perChar(10))
    expect(s.charSpacingPt).toBeUndefined()
  })

  it('skips unmeasurable and out-of-scope blocks', () => {
    const cjk = span('中文', { script: 'cjk' })
    preflightFitBlock(block([lineOf([cjk])]), 10, perChar(50))
    expect(cjk.charSpacingPt).toBeUndefined()

    const rtl = span('שלום', { script: 'hebrew' })
    preflightFitBlock(block([lineOf([rtl])], { dir: 'rtl' }), 10, perChar(50))
    expect(rtl.charSpacingPt).toBeUndefined()

    const unresolved = span('aaaaaaaaaa')
    preflightFitBlock(block([lineOf([unresolved])]), 10, () => null)
    expect(unresolved.charSpacingPt).toBeUndefined()
  })

  it('respects the first-line indent', () => {
    const first = span('aaaaaaaaaa') // 10 × 8.9 = 89 vs avail 100-15=85 → overflow
    const rest = span('aaaaaaaaaa') // 89 vs 100 → fits
    preflightFitBlock(
      block([lineOf([first]), lineOf([rest])], { firstLineIndentPt: 15 }),
      100,
      perChar(8.9),
    )
    expect(first.charSpacingPt).toBeLessThan(0)
    expect(rest.charSpacingPt).toBe(first.charSpacingPt)
  })
})
