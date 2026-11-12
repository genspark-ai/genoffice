import { copyFileSync, existsSync, mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest'
import type { RunStyle } from '@genoffice/pptx-render'

// Vitest turns ?asset imports into /@fs URLs, while font metrics reads the files from disk.
vi.mock('@genoffice/ui/fonts/Carlito-Regular.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Regular.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-Bold.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Bold.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-Italic.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Italic.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-BoldItalic.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-BoldItalic.ttf'),
}))

// shaped-metrics pulls in the harfbuzz wasm (?asset only resolves in electron builds); disconnected in unit tests
vi.mock('../src/main/shaped-metrics', () => ({
  initShapedMetrics: () => {},
  shapedMeasure: () => null,
  shapedFamily: () => null,
  gtMeasure: () => null,
  complexScriptOf: () => null,
}))

import {
  createSystemFontMetrics,
  familyAvailable,
  resetFontRegistry,
  setUserFontDir,
} from '../src/main/fonts'

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

    it('Japanese font requests resolve to matching system faces', () => {
      const yuSans = office ? /^Yu Gothic/ : /^Hiragino Sans$/
      expect(m.displayFamily!(style('游ゴシック'))).toMatch(yuSans)
      expect(m.displayFamily!(style('Yu Gothic'))).toMatch(yuSans)
      expect(m.displayFamily!(style('メイリオ'))).toMatch(office ? /Meiryo/ : /^Hiragino Sans$/)
      expect(m.displayFamily!(style('ＭＳ Ｐゴシック'))).toMatch(
        office ? /MS P?Gothic|Hiragino Sans/ : /^Hiragino Sans$/,
      )
      expect(m.displayFamily!(style('游明朝'))).toMatch(/Yu ?Mincho|Hiragino Mincho/) // public-hygiene: fixture
      expect(m.displayFamily!(style('MS Mincho'))).toMatch(
        office ? /MS Mincho|Hiragino Mincho/ : /^Hiragino Mincho ProN$/,
      )
    })

    it('a missing Japanese sans substitutes to the plain Yu Gothic face, not the narrow-kana UI face', () => {
      const fam = m.displayFamily!(style('BIZ UDPGothic'))
      expect(fam).toMatch(office ? /^Yu Gothic( Medium)?$|^Hiragino Sans$/ : /^Hiragino Sans$/)
      expect(fam).not.toMatch(/UI/)
      expect(m.displayFamily!(style('Yu Gothic UI'))).toMatch(office ? /Yu Gothic UI/ : /Hiragino/)
    })

    it('SinoType (ST*) families resolve to the matching macOS faces, not Calibri', () => {
      expect(m.displayFamily!(style('\u534e\u6587\u6977\u4f53'))).toMatch(/Kai/i)
      expect(m.displayFamily!(style('STKaiti'))).toMatch(/Kai/i)
      expect(m.displayFamily!(style('\u534e\u6587\u5b8b\u4f53'))).toMatch(/Song/i)
      expect(m.displayFamily!(style('\u534e\u6587\u884c\u6977'))).toMatch(/Xingkai|Kai/i)
      expect(m.displayFamily!(style('KaiTi_GB2312'))).toMatch(/Kai/i)
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

    it.runIf(office)(
      'proportional MS faces inside msgothic/msmincho.ttc survive a bold request',
      () => {
        const kana = '\u30a2\u30a4\u30a6\u30a8\u30aa\u304b'
        for (const fam of ['MS PGothic', 'MS UI Gothic', 'MS PMincho']) {
          for (const bold of [false, true]) {
            const s = style(fam, { bold, substScript: 'ja' })
            expect(m.displayFamily!(s, kana)).toBe(fam)
            expect(m.measure(kana, s)).toBeLessThan(560)
          }
        }
      },
    )

    it.runIf(office)(
      'Office families under 8.3 filenames resolve through the name table, not to Calibri',
      () => {
        // Constan.ttf, gara.ttf, pala.ttf, consola.ttf: no filename spells the family
        expect(m.displayFamily!(style('Constantia'))).toBe('Constantia')
        expect(m.displayFamily!(style('Garamond'))).toBe('Garamond')
        expect(m.displayFamily!(style('Palatino Linotype'))).toBe('Palatino Linotype')
        expect(m.displayFamily!(style('Consolas'))).toBe('Consolas')
        expect(m.displayFamily!(style('Constantia', { bold: true }))).toBe('Constantia')
        // and they measure as themselves: Consolas is monospaced, Calibri is not
        const i = m.measure('iiii', style('Consolas'))
        const w = m.measure('wwww', style('Consolas'))
        expect(Math.abs(i - w)).toBeLessThan(1)
      },
    )

    it('Liberation Sans substitutes to its metric twin Arial, as PowerPoint does', () => {
      expect(m.displayFamily!(style('Liberation Sans'))).toBe('Arial')
      expect(m.displayFamily!(style('Liberation Serif'))).toBe('Times New Roman')
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

    it('unresolvable families substitute to Calibri regardless of apparent class (PPT probe truth)', () => {
      const cal = office ? /^Calibri$/ : /^Carlito(?: GO)?$/
      expect(m.displayFamily!(style('Zxqvwt Nonexistent'))).toMatch(cal)
      expect(m.displayFamily!(style('Qqzgaramond'))).toMatch(cal) // serif-looking name
      expect(m.displayFamily!(style('Zxqvwt Mono Courier'))).toMatch(cal) // mono-looking name
      expect(m.displayFamily!(style('Apercu Light'))).toMatch(cal) // weight suffix dropped too
    })

    it("DejaVu Sans takes PowerPoint's table substitute Verdana, unkerned like any substitution", () => {
      expect(m.displayFamily!(style('DejaVu Sans'))).toBe('Verdana')
      expect(m.substituted!(style('DejaVu Sans'))).toBe(true)
      expect(m.displayFamily!(style('Verdana'))).toBe('Verdana')
      expect(m.substituted!(style('Verdana'))).toBe(false)
    })

    it.runIf(office)(
      'Ethiopic text in a face without the script draws in Nyala; Latin words keep the face',
      () => {
        const ethiopic = '\u12a0\u1275\u120c\u1271\u1295'
        expect(m.displayFamily!(style('Calibri'), ethiopic)).toBe('Nyala')
        expect(m.displayFamily!(style('Calibri', { bold: true }), ethiopic)).toBe('Nyala')
        expect(m.displayFamily!(style('Calibri'), 'Athlete')).toBe('Calibri')
        expect(m.displayFamily!(style('Nyala'), ethiopic)).toBe('Nyala')
      },
    )

    it('a "Family Weight" request with an installed base family picks the sub-family face', () => {
      expect(m.displayFamily!(style('Avenir Light'))).toBe('Avenir Light')
      expect(m.measure('Hamburgefontsiv', style('Avenir Light'))).toBeGreaterThan(100)
    })

    it('two-word weight suffixes strip; a base family without the face keeps its own name', () => {
      // Avenir Next Ultra Light is a real macOS face; Georgia has no Light face — the
      // request must fall back to plain Georgia, not register Regular bytes under a lie
      expect(m.displayFamily!(style('Avenir Next Ultra Light'))).toBe('Avenir Next Ultra Light')
      expect(m.displayFamily!(style('Georgia Light'))).toBe('Georgia')
    })

    it.runIf(office)('a weight-suffix request picks the matching cloud face, not Regular', () => {
      const w = (fam: string) =>
        m.measure('contract', { fontFamily: fam, fontSizePx: 100, bold: false, italic: false })
      // Office CloudFonts carry Montserrat in every weight as numeric files
      expect(w('Montserrat SemiBold')).toBeGreaterThan(w('Montserrat') + 5)
    })

    it('non-CJK path is unaffected', () => {
      expect(m.displayFamily!(style('Arial'))).toBe('Arial')
      expect(m.measure('Hello', style('Arial'))).toBeGreaterThan(100)
    })

    it('geometric bullets a face lacks draw and measure in Arial (PowerPoint symbol fallback)', () => {
      // Arial Rounded MT Bold ships with macOS and has • but no ● / ■
      const rounded = style('Arial Rounded MT Bold')
      expect(m.displayFamily!(rounded, 'a')).toMatch(/Arial Rounded/)
      expect(m.displayFamily!(rounded, '•')).toMatch(/Arial Rounded/)
      expect(m.displayFamily!(rounded, '●')).toBe('Arial')
      expect(m.displayFamily!(rounded, '■')).toBe('Arial')
      expect(m.measure('●', rounded)).toBeCloseTo(m.measure('●', style('Arial')), 5)
      // Text that mixes letters in keeps the OS fallback like before
      expect(m.displayFamily!(rounded, '●a')).toMatch(/Arial Rounded/)
    })

    it.runIf(office)('Meiryo UI picks the UI face out of meiryo.ttc, not plain Meiryo', () => {
      expect(m.displayFamily!(style('Meiryo UI'))).toBe('Meiryo UI')
      expect(m.displayFamily!(style('Meiryo UI', { bold: true }))).toBe('Meiryo UI')
      // The UI face's narrow kana are the whole point — matching PowerPoint's wrap positions
      const ui = m.measure('アイウエオかきくけこ', style('Meiryo UI'))
      const plain = m.measure('アイウエオかきくけこ', style('Meiryo'))
      expect(ui).toBeLessThan(plain * 0.85)
    })

    it.runIf(office)('Yu Gothic UI picks the UI face out of YuGothM/YuGothB.ttc', () => {
      expect(m.displayFamily!(style('Yu Gothic UI'))).toBe('Yu Gothic UI')
      expect(m.displayFamily!(style('Yu Gothic UI', { bold: true }))).toMatch(/Yu Gothic UI/)
    })

    it.runIf(office)('legacy Korean fonts (Gulim/Batang) parse via layout-table strip', () => {
      expect(m.displayFamily!(style('굴림'))).toMatch(/Gulim/)
      expect(m.displayFamily!(style('바탕'))).toMatch(/Batang/)
      expect(m.measure('한글', style('굴림'))).toBeCloseTo(200, 0)
    })
  },
)

describe('variable-font family names alias to the static family and back', () => {
  // Static "Zqvstat-*" weights and a lone "ZqvvarVariable" file, all Carlito copies: the
  // filename index is what resolution keys on, so the face names stay Carlito
  const carlito = resolve(process.cwd(), '../../packages/ui/src/fonts')
  let dir: string
  beforeAll(() => {
    dir = mkdtempSync(resolve(tmpdir(), 'slides-vf-'))
    copyFileSync(resolve(carlito, 'Carlito-Regular.ttf'), resolve(dir, 'Zqvstat-Regular.ttf'))
    copyFileSync(resolve(carlito, 'Carlito-Bold.ttf'), resolve(dir, 'Zqvstat-Bold.ttf'))
    copyFileSync(resolve(carlito, 'Carlito-Regular.ttf'), resolve(dir, 'ZqvvarVariable.ttf'))
    setUserFontDir(dir)
    resetFontRegistry()
  })
  afterAll(() => {
    setUserFontDir(null as unknown as string)
    resetFontRegistry()
    rmSync(dir, { recursive: true, force: true })
  })

  it('"X Variable" with only static X installed takes the static face of the right weight', () => {
    const m = createSystemFontMetrics()
    const regular = style('Zqvstat Variable')
    const bold = style('Zqvstat Variable', { bold: true })
    expect(m.displayFamily!(regular)).toMatch(/^Carlito/)
    expect(m.displayFamily!(bold)).toMatch(/^Carlito/)
    expect(m.measure('Hello', regular)).toBeCloseTo(m.measure('Hello', style('Carlito')), 5)
    expect(m.measure('Hello', bold)).toBeCloseTo(
      m.measure('Hello', style('Carlito', { bold: true })),
      5,
    )
    expect(m.measure('Hello', bold)).not.toBeCloseTo(m.measure('Hello', regular), 0)
  })

  it('static X with only "X Variable" installed takes the variable file', () => {
    const m = createSystemFontMetrics()
    expect(m.displayFamily!(style('Zqvvar'))).toMatch(/^Carlito/)
    expect(m.measure('Hello', style('Zqvvar'))).toBeCloseTo(m.measure('Hello', style('Carlito')), 5)
  })

  it('an unknown family stays a substitution either way round', () => {
    expect(familyAvailable('Zqvnothing Variable')).toBe(false)
    expect(familyAvailable('Zqvnothing')).toBe(false)
    expect(familyAvailable('Zqvstat Variable')).toBe(true)
    expect(familyAvailable('Zqvvar')).toBe(true)
  })
})
