import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { ParsedDocFull, StyleInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { docStyleCss } from '../src/renderer/doc-style-css'
import {
  latinViaEastAsiaSlotFor,
  lineHeightFactor,
  setLatinViaEastAsiaSlot,
  wordSlotFaces,
} from '../src/renderer/line-metrics'
import { latinParaFactor, layoutParaOfRuns } from '../src/renderer/editor/para-layout-spec'
import { prefetchFaces, setFontMetricsProvider } from '../src/renderer/font-resolve'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((e) => e.destroy())
  setLatinViaEastAsiaSlot(false)
  setFontMetricsProvider(undefined)
})

function runStyle(attrs: Record<string, unknown>): string {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { styleId: 'Normal' },
          content: [
            { type: 'text', text: 'K-Team +----+', marks: [{ type: 'docTextStyle', attrs }] },
          ],
        },
      ],
    },
  })
  editors.push(editor)
  return editor.view.dom.querySelector<HTMLElement>('[data-doc-style]')!.style.fontFamily
}

const BASE = /\.doc-page, \.pv-page \{([^}]*font-family[^}]*)\}/
const parsedWith = (docDefaults: Record<string, unknown>): ParsedDocFull =>
  ({ styles: new Map(), docDefaults, blocks: [] }) as unknown as ParsedDocFull

// Word probe 2026-11-02: docDefaults rFonts Arial in every slot + w:lang w:val="ko"
// lays every Latin run out in SimSun (0.5em, 14.4pt pitch at 11pt) whatever the
// run's own Latin face, run-level lang included; eastAsia Batang -> Batang Latin;
// Arial Unicode MS keeps itself; no docDefaults lang -> Arial / Courier New.
describe('docDefaults East Asian w:lang w:val lays Latin out with the eastAsia slot', () => {
  it('fires on ja / ko / zh document languages only', () => {
    for (const lang of ['ko', 'ko-KR', 'ja', 'ja-JP', 'zh-CN', 'zh-Hant-TW'])
      expect(latinViaEastAsiaSlotFor({ docDefaults: { lang } }), lang).toBe(true)
    for (const lang of ['en-US', 'de-DE', 'kok', undefined])
      expect(latinViaEastAsiaSlotFor({ docDefaults: { lang } }), String(lang)).toBe(false)
    expect(latinViaEastAsiaSlotFor({})).toBe(false)
  })

  it('maps the slots: a Latin eastAsia face becomes SimSun, an East Asian one stays', () => {
    expect(wordSlotFaces('Courier New', 'Courier New', false)).toEqual({
      ascii: 'Courier New',
      eastAsia: 'Courier New',
    })
    expect(wordSlotFaces('Courier New', 'Courier New', true)).toEqual({
      ascii: 'SimSun',
      eastAsia: 'SimSun',
    })
    expect(wordSlotFaces('Arial', 'Arial', true)).toEqual({ ascii: 'SimSun', eastAsia: 'SimSun' })
    expect(wordSlotFaces('Courier New', 'Batang', true)).toEqual({
      ascii: 'Batang',
      eastAsia: 'Batang',
    })
    expect(wordSlotFaces('Arial Unicode MS', 'Arial Unicode MS', true)).toEqual({
      ascii: 'Arial Unicode MS',
      eastAsia: 'Arial Unicode MS',
    })
    expect(wordSlotFaces('Courier New', null, true)).toEqual({
      ascii: 'SimSun',
      eastAsia: 'SimSun',
    })
    for (const face of ['PingFang SC', 'Heiti SC', 'Apple SD Gothic Neo', 'Malgun Gothic'])
      expect(wordSlotFaces('Arial', face, true), face).toEqual({ ascii: face, eastAsia: face })
    expect(wordSlotFaces(null, null, true)).toEqual({ ascii: null, eastAsia: null })
  })

  it('an installed Latin face keeps mapping to SimSun despite its East Asian code-page bit', async () => {
    setFontMetricsProvider(async () => ({
      unitsPerEm: 2048,
      hheaAscender: 1854,
      hheaDescender: -434,
      hheaLineGap: 67,
      typoAscender: null,
      typoDescender: null,
      typoLineGap: null,
      winAscent: null,
      winDescent: null,
      useTypoMetrics: false,
      xAvgCharWidth: 904,
      codePageRange1: 0x3e0001,
    }))
    await prefetchFaces(['Arial', 'Courier New'])
    expect(wordSlotFaces('Arial', 'Arial', true)).toEqual({ ascii: 'SimSun', eastAsia: 'SimSun' })
    expect(wordSlotFaces('Courier New', null, true)).toEqual({
      ascii: 'SimSun',
      eastAsia: 'SimSun',
    })
  })

  it('an ascii-only run takes the inherited East Asian face, not its own Latin name', () => {
    setLatinViaEastAsiaSlot(true, 'Batang')
    expect(runStyle({ font: 'Courier New', fontAscii: 'Courier New' })).toContain('Batang')
    expect(
      runStyle({ font: 'Courier New', fontAscii: 'Courier New', eastAsiaFont: 'Courier New' }),
    ).toContain('GenOffice SimSun Latin')
    setLatinViaEastAsiaSlot(true, 'Arial')
    expect(runStyle({ font: 'Courier New', fontAscii: 'Courier New' })).toContain(
      'GenOffice SimSun Latin',
    )
  })

  it('a style without fonts keeps inheriting instead of pinning the remapped face', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('Strong', {
      styleId: 'Strong',
      name: 'Strong',
      type: 'character',
      display: { bold: true },
    } as StyleInfo)
    styles.set('Code', {
      styleId: 'Code',
      name: 'Code',
      type: 'character',
      display: { font: 'Courier New', fontAscii: 'Courier New' },
    } as StyleInfo)
    const css = docStyleCss({
      styles,
      docDefaults: { asciiFont: 'Arial', eastAsiaFont: 'Arial', lang: 'ko' },
      blocks: [],
    } as unknown as ParsedDocFull)
    const strong = /\[data-style="Strong"\] \{([^}]*)\}/.exec(css)![1]
    expect(strong).not.toContain('font-family')
    const code = /\[data-style="Code"\] \{([^}]*)\}/.exec(css)![1]
    expect(code).toContain('GenOffice SimSun Latin')
    expect(code).not.toContain('Courier New')
  })

  it('a Latin text paragraph without rFonts takes the SimSun factor, an empty one does not', () => {
    const para = (text: string) => layoutParaOfRuns(text ? [{ text } as never] : null, text, 'auto')
    expect(latinParaFactor(para('Shaker Tarima'), 'var(--doc-line-factor-latin,1.2)')).toBe(
      'var(--doc-line-factor-latin,1.2)',
    )
    setLatinViaEastAsiaSlot(true, 'Arial')
    expect(latinParaFactor(para('Shaker Tarima'), 'var(--doc-line-factor-latin,1.2)')).toBe('1.3')
    expect(latinParaFactor(para(''), 'var(--doc-line-factor-latin,1.2)')).toBe(
      'var(--doc-line-factor-latin,1.2)',
    )
    setLatinViaEastAsiaSlot(true, 'Batang')
    expect(latinParaFactor(para('Shaker Tarima'), 'var(--doc-line-factor-latin,1.2)')).toBe(
      String(lineHeightFactor('Batang')),
    )
  })

  it('runs render through the SimSun chain when the flag is on', () => {
    expect(runStyle({ font: 'Courier New', fontAscii: 'Courier New' })).toContain('Courier New')
    setLatinViaEastAsiaSlot(true)
    const ff = runStyle({ font: 'Courier New', fontAscii: 'Courier New' })
    expect(ff).toContain('GenOffice SimSun Latin')
    expect(ff).not.toContain('Courier New')
    expect(runStyle({ font: 'Arial Unicode MS', fontAscii: 'Arial Unicode MS' })).toContain(
      'Arial Unicode MS',
    )
  })

  it('the document base family follows SimSun under lang=ko, the empty-line factor stays Arial', () => {
    const ko = docStyleCss(parsedWith({ asciiFont: 'Arial', eastAsiaFont: 'Arial', lang: 'ko' }))
    const base = BASE.exec(ko)![1]
    expect(base).toContain('GenOffice SimSun Latin')
    expect(base).not.toContain('Arial')
    // empty paragraphs keep the ascii face's height (Word probe 2026-11-02)
    expect(base).toContain('--doc-line-factor-latin:1.15')
    expect(ko).toMatch(/--doc-latin-chain:[^;}]*GenOffice SimSun Latin/)
    const en = docStyleCss(parsedWith({ asciiFont: 'Arial', eastAsiaFont: 'Arial', lang: 'en-US' }))
    expect(BASE.exec(en)![1]).toContain('Arial')
    expect(BASE.exec(en)![1]).not.toContain('SimSun')
    const batang = docStyleCss(
      parsedWith({ asciiFont: 'Arial', eastAsiaFont: 'Batang', lang: 'ko' }),
    )
    expect(BASE.exec(batang)![1]).toContain('Batang')
    expect(BASE.exec(batang)![1]).not.toContain('Arial')
  })
})
