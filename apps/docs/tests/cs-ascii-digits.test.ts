import { DOMSerializer } from '@tiptap/pm/model'
import { renderTableSpec } from '../src/renderer/editor/protected-render'
import { describe, expect, it } from 'vitest'
import {
  csAsciiDigitRanges,
  inlineSegments,
  inlineRunsOf,
} from '../src/renderer/text-layout/inline-segments'

const AR = 'التعليم العالي'

describe('csAsciiDigitRanges (Word for Mac probe 2026-11-08, docDefaults cs Arial + rtl)', () => {
  it('digit groups beside Arabic letters take the ascii face, the period inside a number too', () => {
    const text = `1.2 ${AR}: 2024`
    expect(csAsciiDigitRanges(text)).toEqual([
      { from: 0, to: 3 },
      { from: text.length - 4, to: text.length },
    ])
  })

  it('digits after a Latin letter join the Latin stretch and stay in the cs face', () => {
    expect(csAsciiDigitRanges('abc 1.2 def 2024')).toEqual([])
    expect(csAsciiDigitRanges(' 2024', 'def')).toEqual([])
    expect(csAsciiDigitRanges(' 2024', `def ${AR}`)).toEqual([{ from: 1, to: 5 }])
  })

  it('punctuation without digits and Arabic-Indic digits are left to the cs face', () => {
    expect(csAsciiDigitRanges(`${AR}. (x): 50%`)).toEqual([])
    expect(csAsciiDigitRanges('١.٢ ' + AR)).toEqual([])
  })
  it('multiplication and division signs are neutrals, not Latin letters (W7)', () => {
    expect(csAsciiDigitRanges(`${AR} 3 \u00d7 4 \u00f7 2`)).toEqual([
      { from: AR.length + 1, to: AR.length + 2 },
      { from: AR.length + 5, to: AR.length + 6 },
      { from: AR.length + 9, to: AR.length + 10 },
    ])
    expect(csAsciiDigitRanges('x \u00d7 4', AR)).toEqual([])
  })
})

describe('inlineSegments doc-cs-ascii', () => {
  const para = (runs: Parameters<typeof inlineRunsOf>[0], inheritedAscii = 'Calibri') => ({
    runs: inlineRunsOf(runs).runs,
    autospace: null,
    inheritedAscii,
    inheritedEa: null,
    wordWrap: true,
    balancedSpaces: false,
    justified: false,
    simsunGapLineHeight: null,
    lineMultiple: 'inherited' as const,
  })

  it('an rtl run under cs Arial moves its digits to the inherited Calibri chain and lifts the line to 1.22', () => {
    const segs = inlineSegments(para([{ text: `1.6 ${AR}`, csFont: 'Arial', cs: true }])).segments
    expect(segs).toHaveLength(1)
    expect(segs[0]).toMatchObject({ from: 0, to: 3, cls: 'doc-cs-ascii' })
    expect(segs[0].style).toMatch(/^font-family:'Calibri'/)
    expect(segs[0].style).toContain('line-height:')
    expect(segs[0].style).toMatch(/line-height:calc\(/)
  })

  it('a digit group whose first digit an autospace pad wraps keeps the rest as its own span', () => {
    const segs = inlineSegments({
      ...para([{ text: `${AR} \u4e2d365 \u4e2d`, csFont: 'Arial', cs: true }]),
      autospace: true,
    }).segments
    const digits = segs.filter((s) => s.cls === 'doc-cs-ascii')
    expect(digits.map((s) => `${s.from}-${s.to}`)).toEqual([`${AR.length + 3}-${AR.length + 5}`])
    for (let i = 0; i < segs.length; i++)
      for (let k = i + 1; k < segs.length; k++)
        expect(segs[i].from < segs[k].to && segs[k].from < segs[i].to).toBe(false)
  })

  it('no lift when the ascii face is not taller than the cs face', () => {
    const segs = inlineSegments(
      para([{ text: `${AR} 2024`, csFont: 'Calibri', cs: true, fontAscii: 'Arial' }]),
    ).segments
    expect(segs).toHaveLength(1)
    expect(segs[0].style).toMatch(/^font-family:'Arial'/)
    expect(segs[0].style).not.toContain('line-height')
  })

  it('same face in both slots, a non-cs run, or pinned lines yield no face switch', () => {
    expect(
      inlineSegments(para([{ text: `1.6 ${AR}`, csFont: 'Calibri', cs: true }])).segments,
    ).toEqual([])
    expect(inlineSegments(para([{ text: `1.6 ${AR}`, cs: true }])).segments).toEqual([])
    const pinned = {
      ...para([{ text: `1.6 ${AR}`, csFont: 'Arial', cs: true }]),
      lineMultiple: null,
    }
    expect(inlineSegments(pinned).segments[0].style).not.toContain('line-height')
  })
})

describe('static DOM (protected tables) lifts digit lines like the editor decorations', () => {
  const render = (spec: unknown): HTMLElement =>
    DOMSerializer.renderSpec(document, spec as never).dom as HTMLElement
  const run = { text: `1.6 ${AR}`, csFont: 'Arial', cs: true, fontAscii: 'Calibri' }

  it('an auto-spaced rich cell paragraph carries the per-line lift', () => {
    const dom = render(renderTableSpec({ rows: [[{ paras: [''], richParas: [{ runs: [run] }] }]] }))
    const span = dom.querySelector<HTMLElement>('.doc-cs-ascii')!
    expect(span.textContent).toBe('1.6')
    expect(span.style.lineHeight).toMatch(/^calc\(/)
  })

  it('an exact-spaced cell paragraph switches the face but keeps the pinned line', () => {
    const dom = render(
      renderTableSpec({
        rows: [
          [{ paras: [''], richParas: [{ runs: [run], lineRule: 'exact', lineRawTwips: 240 }] }],
        ],
      }),
    )
    const span = dom.querySelector<HTMLElement>('.doc-cs-ascii')!
    expect(span.style.fontFamily).toMatch(/^"Calibri"/)
    expect(span.style.lineHeight).toBe('')
  })
})
