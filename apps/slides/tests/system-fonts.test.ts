import { existsSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import type { RunStyle } from '@genoffice/pptx-render'

// shaped-metrics pulls in the harfbuzz wasm (?asset only resolves in electron builds); disconnected in unit tests
vi.mock('../src/main/shaped-metrics', () => ({
  initShapedMetrics: () => {},
  shapedMeasure: () => null,
  shapedFamily: () => null,
}))

import { createSystemFontMetrics } from '../src/main/fonts'

const style = (fontFamily: string, over: Partial<RunStyle> = {}): RunStyle => ({
  fontFamily,
  fontSizePx: 100,
  bold: false,
  italic: false,
  ...over,
})

const mac = process.platform === 'darwin'
// With Office for Mac installed, its private DFonts (real Yu Gothic/Malgun/MingLiU…) win over
// same-script substitution — the assertions accept either the real family or the substitute.
const office = existsSync('/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts')

describe.runIf(mac)(
  'Japanese/Korean/Traditional-Chinese font substitution (macOS system fonts)',
  () => {
    const m = createSystemFontMetrics()

    it('missing Japanese fonts substitute to Hiragino (same family for metrics and rendering)', () => {
      const yuSans = office ? /^Yu Gothic/ : /^Hiragino Sans$/
      expect(m.displayFamily!(style('游ゴシック'))).toMatch(yuSans)
      expect(m.displayFamily!(style('Yu Gothic'))).toMatch(yuSans)
      expect(m.displayFamily!(style('メイリオ'))).toMatch(office ? /Meiryo/ : /^Hiragino Sans$/)
      expect(m.displayFamily!(style('ＭＳ Ｐゴシック'))).toMatch(
        office ? /MS P?Gothic|Hiragino Sans/ : /^Hiragino Sans$/,
      )
      expect(m.displayFamily!(style('游明朝'))).toMatch(
        office ? /Yu Mincho|Hiragino Mincho/ : /^Hiragino Mincho ProN$/,
      )
      expect(m.displayFamily!(style('MS Mincho'))).toMatch(
        office ? /MS Mincho|Hiragino Mincho/ : /^Hiragino Mincho ProN$/,
      )
    })

    it('missing Korean/Traditional-Chinese fonts substitute to same-script fonts, not Arial/Simplified Chinese', () => {
      const ko = office ? /Malgun Gothic|Apple SD Gothic Neo/ : /^Apple SD Gothic Neo$/
      expect(m.displayFamily!(style('맑은 고딕'))).toMatch(ko)
      expect(m.displayFamily!(style('Malgun Gothic'))).toMatch(ko)
      expect(m.displayFamily!(style('Microsoft JhengHei'))).toMatch(
        office ? /JhengHei|Heiti|PingFang|Songti/ : /Heiti|PingFang|Songti/,
      )
      expect(m.displayFamily!(style('PMingLiU'))).toMatch(office ? /MingLiU|Songti/ : /Songti/)
    })

    it('ttc parsing yields real metrics: CJK full-width 1em, Latin non-heuristic', () => {
      expect(m.measure('あア亜', style('游ゴシック'))).toBeCloseTo(300, 0)
      expect(m.measure('한글', style('맑은 고딕'))).toBeGreaterThan(150)
      const met = m.metrics(style('游ゴシック'))
      expect(met.ascent).toBeGreaterThan(50)
    })

    it('bold picks the matching face inside the ttc', () => {
      expect(m.displayFamily!(style('맑은 고딕', { bold: true }))).toMatch(
        office ? /Malgun Gothic|Apple SD Gothic Neo/ : /^Apple SD Gothic Neo$/,
      )
      expect(m.measure('한', style('맑은 고딕', { bold: true }))).toBeGreaterThan(50)
    })

    it('non-CJK path is unaffected', () => {
      expect(m.displayFamily!(style('Arial'))).toBe('Arial')
      expect(m.measure('Hello', style('Arial'))).toBeGreaterThan(100)
    })
  },
)
