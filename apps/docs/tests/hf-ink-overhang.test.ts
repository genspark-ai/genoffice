/**
 * A strip paragraph under an auto line multiple below 1 (w:line 14 = 0.058)
 * shrinks its line box under the glyphs; Word paints them whole over the
 * strip edge (probe 2026-11-02), so the strip's paint clip has to give room
 * for the overhang without changing the flow height.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { hfInkOverhangPx, makeGapHfEl } from '../src/renderer/editor/hf-dom'

const squeezed = (
  lineRawTwips: number,
  runs: object[] = [{ text: 'header', sizeHalfPoints: 16 }],
) => ({ runs, lineRule: 'auto', lineRawTwips }) as never

describe('strip glyph overhang under a line multiple below 1', () => {
  it('is half the shrunken part of the natural line, nothing at or above a single line', () => {
    const over = hfInkOverhangPx([squeezed(14)])
    // 8pt Calibri line 9.76pt = 13.0px, (1 - 0.058) / 2 of it = 6.1px, ceiled + 1
    expect(over).toBeGreaterThanOrEqual(7)
    expect(over).toBeLessThanOrEqual(8)
    expect(hfInkOverhangPx([squeezed(120)])).toBeLessThan(over)
    expect(hfInkOverhangPx([squeezed(240)])).toBe(0)
    expect(hfInkOverhangPx([squeezed(360)])).toBe(0)
    expect(
      hfInkOverhangPx([{ runs: [{ text: 'x' }], lineRule: 'exact', lineRawTwips: 14 } as never]),
    ).toBe(0)
  })

  it('the gap strip carries the allowance as --hf-ink-over; a plain strip does not', () => {
    const el = makeGapHfEl({
      kind: 'header',
      value: { text: '', paras: [squeezed(14)] },
      pageNo: 1,
      pageTotal: 1,
    })
    expect(el.style.getPropertyValue('--hf-ink-over')).toMatch(/^\d+px$/)
    const plain = makeGapHfEl({
      kind: 'header',
      value: { text: '', paras: [squeezed(240)] },
      pageNo: 1,
      pageTotal: 1,
    })
    expect(plain.style.getPropertyValue('--hf-ink-over')).toBe('')
  })

  it('the strip clips with overflow: clip widened by the allowance (hidden cut the glyphs)', () => {
    const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')
    const strip = /\n\.page-hf \{([^}]*)\}/.exec(css)?.[1] ?? ''
    expect(strip).toContain('overflow: clip;')
    expect(strip).toContain('overflow-clip-margin: var(--hf-ink-over, 0px);')
    expect(strip).not.toContain('overflow: hidden')
  })
})
