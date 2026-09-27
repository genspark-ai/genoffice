/**
 * Typed-grid strut alignment (SAS prod_043): Chromium unions the strut's and
 * each inline box's half-leading geometry, so a Latin-primary strut under
 * EA-primary run spans pads every line ~1px past its grid cell. Mixed
 * declared/inherited-font CJK paragraphs get .doc-grid-strut, and docStyleCss
 * emits a glyphless metrics face (the PUA-blank source under the EA face's
 * measured ascent/descent) that leads the paragraph chain.
 */
import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

afterEach(() => {
  vi.restoreAllMocks()
  document.head.innerHTML = ''
})

const para = (content: object[]) =>
  ({ type: 'doc', content: [{ type: 'docParagraph', content }] }) as never

const text = (t: string, attrs?: Record<string, unknown>) => ({
  type: 'text',
  text: t,
  ...(attrs ? { marks: [{ type: 'docTextStyle', attrs }] } : {}),
})

const pOf = (editor: Editor) => editor.view.dom.querySelector('p') as HTMLElement

describe('blockAttrs .doc-grid-strut class', () => {
  it('marks paragraphs mixing declared-CJK-face runs with inheriting runs', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([text('本研究は', { font: 'ＭＳ 明朝', fontAscii: 'ＭＳ 明朝' }), text('UT')]),
    })
    expect(pOf(editor).classList.contains('doc-grid-strut')).toBe(true)
    expect(pOf(editor).style.fontFamily).toBe('')
    editor.destroy()
  })

  it('all-declared paragraphs keep the direct family swap instead', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([text('本研究は', { font: 'ＭＳ 明朝', fontAscii: 'ＭＳ 明朝' })]),
    })
    expect(pOf(editor).classList.contains('doc-grid-strut')).toBe(false)
    expect(pOf(editor).style.fontFamily).not.toBe('')
    editor.destroy()
  })

  it('inheriting paragraphs with CJK text get .doc-ea-strut instead', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([text('\u53cd\u6b3a\u8bc8 AntiFraud')]),
    })
    expect(pOf(editor).classList.contains('doc-ea-strut')).toBe(true)
    expect(pOf(editor).classList.contains('doc-grid-strut')).toBe(false)
    editor.destroy()
  })

  it('inheriting Latin-only paragraphs stay unmarked', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([text('plain Latin')]),
    })
    expect(pOf(editor).classList.contains('doc-grid-strut')).toBe(false)
    expect(pOf(editor).classList.contains('doc-ea-strut')).toBe(false)
    editor.destroy()
  })

  it('Latin-declared mixes stay unmarked (an EA-geometry strut would add slop)', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([text('Latin run', { fontAscii: 'Arial' }), text('plain')]),
    })
    expect(pOf(editor).classList.contains('doc-grid-strut')).toBe(false)
    editor.destroy()
  })
})

const GRID_SECT =
  '<w:sectPr><w:pgSz w:w="11900" w:h="16840"/>' +
  '<w:docGrid w:type="linesAndChars" w:linePitch="329" w:charSpace="-820"/></w:sectPr>'

function parsedWith(
  sectPrXml: string | null,
  eastAsiaFont?: string,
  slot?: { eaSlotEmpty?: boolean; eaFromLang?: boolean },
  normal?: Record<string, unknown>,
): ParsedDocFull {
  const styles = new Map()
  if (normal)
    styles.set('Normal', { id: 'Normal', type: 'paragraph', isDefault: true, display: normal })
  return {
    styles,
    docDefaults: eastAsiaFont ? { asciiFont: 'Century', eastAsiaFont, ...slot } : {},
    blocks: sectPrXml ? [{ docxIndex: 0, originalXml: `<w:p>${sectPrXml}</w:p>` }] : [],
  } as unknown as ParsedDocFull
}

function stubMetricsCanvas(ascent: number, descent: number) {
  const fake = {
    font: '',
    measureText: () => ({ fontBoundingBoxAscent: ascent, fontBoundingBoxDescent: descent }),
  }
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(
    fake as unknown as CanvasRenderingContext2D,
  )
}

function mountBlankFace() {
  const style = document.createElement('style')
  style.textContent =
    "@font-face { font-family: 'GenOffice PUA Blank'; src: url('./GenOfficePUABlank.woff2') format('woff2'); }"
  document.head.appendChild(style)
}

describe('docStyleCss grid strut face', () => {
  it('emits the metrics @font-face and the strut rule for typed-grid EA docs', () => {
    stubMetricsCanvas(1143, 286)
    mountBlankFace()
    const css = docStyleCss(parsedWith(GRID_SECT, 'ＭＳ 明朝'))
    expect(css).toContain("@font-face { font-family:'GenOffice Grid Strut'")
    expect(css).toContain('ascent-override:114.3%')
    expect(css).toContain('descent-override:28.6%')
    expect(css).toContain('GenOfficePUABlank.woff2')
    // the Office-private MS Mincho renders through a stand-in, so the
    // inheriting .doc-ea-strut paragraphs share the strut too
    expect(css).toContain(
      ".doc-page :is(.doc-grid-strut,.doc-ea-strut) { font-family:'GenOffice Grid Strut',var(--doc-grid-strut-tail,serif) }",
    )
    expect(css).toContain('--doc-grid-strut-tail:')
  })

  it('emits nothing without an EA face or with a Latin one', () => {
    stubMetricsCanvas(1143, 286)
    mountBlankFace()
    expect(docStyleCss(parsedWith(GRID_SECT))).not.toContain('Grid Strut')
    expect(docStyleCss(parsedWith(null, 'Century'))).not.toContain('Grid Strut')
  })

  it('a missing EA face rendered by a stand-in gets the strut without a grid', () => {
    // Yu Mincho body on a Latin strut: every CJK line unioned ~0.17em past its
    // 1.44 x 1.08 line-height (18.9pt for Word's 17.0 at 11pt, +4 pages)
    stubMetricsCanvas(880, 120)
    mountBlankFace()
    const css = docStyleCss(parsedWith(null, '\u6e38\u660e\u671d'))
    expect(css).toContain("@font-face { font-family:'GenOffice Grid Strut'")
    expect(css).toContain('ascent-override:88%')
    expect(css).toContain('.doc-page :is(.doc-grid-strut,.doc-ea-strut) {')
    expect(docStyleCss(parsedWith(null, '\uff2d\uff33 \u660e\u671d'))).toContain('.doc-ea-strut')
  })

  it('an empty theme slot resolved by the language still drives the grid strut', () => {
    stubMetricsCanvas(952, 351)
    mountBlankFace()
    const resolved = parsedWith(GRID_SECT, 'PMingLiU', { eaSlotEmpty: true, eaFromLang: true })
    expect(docStyleCss(resolved)).toContain('ascent-override:95.2%')
    const unresolved = parsedWith(GRID_SECT, 'Yu Gothic', { eaSlotEmpty: true })
    expect(docStyleCss(unresolved)).not.toContain('Grid Strut')
  })

  it('Office-private faces under a metric alias get the strut without a grid', () => {
    stubMetricsCanvas(969, 391)
    mountBlankFace()
    const css = docStyleCss(parsedWith(null, 'DengXian'))
    expect(css).toContain("@font-face { font-family:'GenOffice Grid Strut'")
    expect(css).toContain('ascent-override:96.9%')
    expect(css).toContain('descent-override:39.1%')
    expect(css).toContain('.doc-page :is(.doc-grid-strut,.doc-ea-strut) {')
  })

  it('takes the face the CJK spans render with, not the w:lang backfill', () => {
    stubMetricsCanvas(969, 391)
    mountBlankFace()
    // zh-CN empty theme slot: Normal resolves to DengXian, docDefaults backfills SimSun
    const css = docStyleCss(
      parsedWith(
        null,
        'SimSun',
        { eaSlotEmpty: true, eaFromLang: true },
        {
          font: 'DengXian',
          fontAscii: 'Calibri',
          eaSlotEmpty: true,
        },
      ),
    )
    expect(css).toContain("@font-face { font-family:'GenOffice Grid Strut'")
    expect(css).toContain('.doc-ea-strut')
  })

  it('degrades silently when canvas metrics are unavailable', () => {
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null)
    mountBlankFace()
    // a face not probed by earlier tests (fontChainMetricsPct caches per chain)
    expect(docStyleCss(parsedWith(GRID_SECT, 'ＭＳ ゴシック'))).not.toContain('Grid Strut')
  })
})
