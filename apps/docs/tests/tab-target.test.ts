import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { resolveTabTarget, type TabTargetInput } from '../src/renderer/editor/decoration-extensions'

// Letter page, 1" margins: 6.5" column = 624px; a TOC-style right stop at the edge
// (pinned to paraW - 1 so the 0.5px round-up cannot re-trigger a wrap)
const PARA_W = 624
const base: TabTargetInput = {
  x: 100,
  minAdv: 5,
  segWidth: 12,
  restWidth: 12,
  paraW: PARA_W,
  stops: [{ x: 624, val: 'right', leader: 'dot' }],
  gridPx: 48,
}

describe('resolveTabTarget', () => {
  it('right-aligns a short segment at the stop and glues it to the tab', () => {
    const r = resolveTabTarget(base)
    expect(r).toMatchObject({
      target: 611,
      val: 'right',
      leader: 'dot',
      collapsed: false,
      glue: true,
    })
  })

  // the text before the tab has run past (stop - segment): Word collapses the
  // tab to zero width and the segment flows on inline, so the tab must not be
  // glued (a glued tab would wrap the preceding word down with the segment).
  // The advance is half a space plus slack: Chromium skips a tab closer than
  // half the paragraph font's space width to its stop on to the next stop,
  // which would push the whole segment onto the next line
  it('collapses without glue when the segment no longer fits before the stop', () => {
    const r = resolveTabTarget({ ...base, x: 615 })
    expect(r).toMatchObject({ collapsed: true, glue: false, val: 'right' })
    expect(r.target).toBeCloseTo(618)
    expect(r.target - 615).toBeGreaterThan(base.minAdv / 2)
    expect(r.target - 615).toBeLessThan(base.minAdv)
    // a job-heading row: long left text, a right stop near the edge, and a
    // date segment wider than the room left before the stop
    const heading = resolveTabTarget({
      ...base,
      x: 528,
      segWidth: 210,
      restWidth: 210,
      paraW: 659,
      stops: [{ x: 653, val: 'right' }],
    })
    expect(heading).toMatchObject({ collapsed: true, glue: false })
    expect(heading.target).toBeCloseTo(531)
    // once the same tab starts a fresh line the segment fits and is glued again
    expect(resolveTabTarget({ ...base, x: 60 })).toMatchObject({
      target: 611,
      collapsed: false,
      glue: true,
    })
  })

  it('advances to the default grid, unglued, when the text already passes the stop', () => {
    const r = resolveTabTarget({ ...base, x: 640, paraW: 700, stops: [{ x: 600, val: 'right' }] })
    expect(r).toMatchObject({ target: 672, val: 'left', collapsed: false, glue: false })
  })

  // TOC "<tab><tab>Box 1: ...<tab>9" with stops 397/900/8820 and left=360: the
  // first tab sits 2.5px before the 397 stop; Word still lands on it, GenOffice
  // used to skip to 900 and send the next tab to the far leader stop
  it('takes a custom stop closer than the minimum advance and collapses to it', () => {
    const stops: TabTargetInput['stops'] = [
      { x: 26.5, val: 'left' },
      { x: 60, val: 'left' },
      { x: 588, val: 'right', leader: 'dot' },
    ]
    const first = resolveTabTarget({ ...base, x: 24, segWidth: 0, restWidth: 300, stops })
    expect(first).toMatchObject({ val: 'left', leader: undefined, collapsed: true })
    expect(first.target).toBeGreaterThan(24)
    expect(first.target).toBeLessThan(26.5 + base.minAdv)
    expect(
      resolveTabTarget({ ...base, x: 26.5, segWidth: 300, restWidth: 300, stops }).target,
    ).toBe(60)
  })

  it('never glues left tabs or wide / wrapped / empty segments', () => {
    expect(resolveTabTarget({ ...base, stops: [{ x: 300, val: 'left' }] }).glue).toBe(false)
    expect(resolveTabTarget({ ...base, stops: [] }).glue).toBe(false)
    expect(resolveTabTarget({ ...base, segWidth: 400, restWidth: 400 }).glue).toBe(false)
    expect(resolveTabTarget({ ...base, segWidth: PARA_W, restWidth: PARA_W }).glue).toBe(false)
    expect(resolveTabTarget({ ...base, segWidth: 0, restWidth: 0 }).glue).toBe(false)
    expect(resolveTabTarget({ ...base, segWidth: 12, restWidth: 400 }).glue).toBe(false)
  })

  it('centers a center-stop segment and glues it', () => {
    const r = resolveTabTarget({
      ...base,
      stops: [{ x: 300, val: 'center' }],
      segWidth: 40,
      restWidth: 40,
    })
    expect(r).toMatchObject({ target: 280, val: 'center', glue: true })
  })

  it('falls back to the default grid past the last custom stop', () => {
    expect(resolveTabTarget({ ...base, x: 630, stops: [] }).collapsed).toBe(true)
    expect(resolveTabTarget({ ...base, stops: [] }).target).toBe(144)
    expect(resolveTabTarget({ ...base, stops: [], gridPx: 0 }).target).toBe(105)
  })

  it('keeps a left stop under a breakable segment that overflows by more than a space', () => {
    // Word wraps the segment at its spaces (probe 2026-09-17); pinning would
    // pull the stop left and let text that Word wraps ride the line
    const input = {
      ...base,
      x: 22,
      stops: [{ x: 65, val: 'left' as const }],
      segWidth: 570,
      restWidth: 570,
    }
    expect(resolveTabTarget({ ...input, segBreakable: true }).target).toBe(65)
    // an unbreakable segment (single word) still pins: Chromium would
    // otherwise drop the whole word to the next line
    expect(resolveTabTarget({ ...input, segBreakable: false }).target).toBe(PARA_W - 1 - 570)
    // a rounding-class overflow (under a space) pins either way
    expect(
      resolveTabTarget({ ...input, segWidth: 560, restWidth: 560, segBreakable: true }).target,
    ).toBe(PARA_W - 1 - 560)
  })

  it('pins a stop past the edge flush right, reserving the trailing segments', () => {
    const r = resolveTabTarget({
      ...base,
      x: 400,
      stops: [{ x: 700, val: 'left' }],
      segWidth: 40,
      restWidth: 100,
    })
    expect(r.target).toBe(PARA_W - 1 - 100)
    expect(r.glue).toBe(false)
  })
})

// Word: on the first line of a hanging indent the left indent is an implicit
// tab stop (w:ind left=900 hanging=900 -> 60px; column 624px, default grid 48px)
describe('resolveTabTarget with a hanging indent', () => {
  const hang: TabTargetInput = {
    ...base,
    x: 20,
    hangingX: 60,
    stops: [{ x: 624, val: 'right', leader: 'dot' }],
  }

  it('sends a first-line tab before the indent to the indent, not the leader stop beyond it', () => {
    expect(resolveTabTarget(hang)).toMatchObject({
      target: 60,
      val: 'left',
      leader: undefined,
      collapsed: false,
      glue: false,
    })
  })

  it('beats an explicit stop beyond the indent (TOC "III.<tab>" with a 397-twip stop past a 360 hang)', () => {
    const r = resolveTabTarget({
      ...hang,
      hangingX: 24,
      x: 15,
      stops: [
        { x: 26.5, val: 'left' },
        { x: 60, val: 'left' },
      ],
    })
    expect(r.target).toBe(24)
  })

  it('still honours an explicit stop before the indent', () => {
    const r = resolveTabTarget({
      ...hang,
      x: 0,
      stops: [
        { x: 24, val: 'left' },
        { x: 624, val: 'right', leader: 'dot' },
      ],
    })
    expect(r).toMatchObject({ target: 24, val: 'left' })
  })

  it('uses a stop exactly at the indent, keeping its leader', () => {
    const r = resolveTabTarget({ ...hang, stops: [{ x: 60, val: 'left', leader: 'dot' }] })
    expect(r).toMatchObject({ target: 60, leader: 'dot' })
  })

  it('ignores the indent once the tab is past it (wrapped line or wide label)', () => {
    // a tab on a wrapped line never sits before the indent: the caller omits hangingX
    expect(resolveTabTarget({ ...hang, hangingX: undefined }).target).toBe(611)
    // a label wider than the hang: next explicit stop, else the default grid
    expect(resolveTabTarget({ ...hang, x: 70 }).target).toBe(611)
    expect(resolveTabTarget({ ...hang, x: 70, stops: [] }).target).toBe(96)
  })

  // TOC "<tab>Box 2: ...<tab>15" (left=1620 hanging=1620, stop 900): the text
  // after the label overflows the line and wraps to the indent in Word
  it('does not pin an overflowing first-line segment flush right', () => {
    const r = resolveTabTarget({
      ...hang,
      x: 0,
      hangingX: 108,
      segWidth: 450,
      restWidth: 462,
      paraW: 492,
      stops: [
        { x: 60, val: 'left' },
        { x: 588, val: 'right', leader: 'dot' },
      ],
    })
    expect(r).toMatchObject({ target: 60, val: 'left', collapsed: false })
  })

  it('collapses when the indent is closer than the minimum advance', () => {
    const r = resolveTabTarget({ ...hang, x: 57 })
    expect(r.collapsed).toBe(true)
    expect(r.target).toBeCloseTo(60)
  })
})

describe('tab leader / glue CSS', () => {
  const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

  it('draws dot leaders as period glyphs anchored at the stop, not as a border', () => {
    const rule = /\.doc-tab-leader-dot::before\s*\{\s*content:\s*'(\.+)';/.exec(css)
    expect(rule, 'dot leader pseudo-element').not.toBeNull()
    expect(rule![1].length).toBeGreaterThan(200)
    expect(css).not.toMatch(/\.doc-tab-leader-dot[^{]*\{[^}]*border-bottom/)
    expect(css).toMatch(/\.doc-tab-leader-hyphen::before\s*\{[^}]*direction: rtl/)
  })

  it('renders hanging-area tabs as an inline-block box that resets the paragraph indent', () => {
    const rule = css.match(/\.doc-tab-fixed \{[^}]*\}/)?.[0] ?? ''
    expect(rule).toContain('display: inline-block')
    expect(rule).toContain('text-indent: 0')
    expect(rule).toContain('tab-size: 0')
  })

  it('glues right-ish tabs with white-space: pre', () => {
    expect(css).toMatch(/\.doc-tab-glue\s*\{\s*white-space:\s*pre;\s*\}/)
  })
})
