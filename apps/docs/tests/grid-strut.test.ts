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

function parsedWith(sectPrXml: string | null, eastAsiaFont?: string): ParsedDocFull {
  return {
    styles: new Map(),
    docDefaults: eastAsiaFont ? { asciiFont: 'Century', eastAsiaFont } : {},
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
    expect(css).toContain(
      ".doc-page .doc-grid-strut { font-family:'GenOffice Grid Strut',var(--doc-grid-strut-tail,serif) }",
    )
    expect(css).toContain('--doc-grid-strut-tail:')
  })

  it('emits nothing without a typed grid or without an EA face', () => {
    stubMetricsCanvas(1143, 286)
    mountBlankFace()
    expect(docStyleCss(parsedWith(null, 'ＭＳ 明朝'))).not.toContain('Grid Strut')
    expect(docStyleCss(parsedWith(GRID_SECT))).not.toContain('Grid Strut')
  })

  it('degrades silently when canvas metrics are unavailable', () => {
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null)
    mountBlankFace()
    // a face not probed by earlier tests (fontChainMetricsPct caches per chain)
    expect(docStyleCss(parsedWith(GRID_SECT, 'ＭＳ ゴシック'))).not.toContain('Grid Strut')
  })
})
