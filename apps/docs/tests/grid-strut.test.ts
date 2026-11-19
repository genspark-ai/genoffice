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
import { declaredStrutFace, resetStrutFaces } from '../src/renderer/text-layout/strut-face'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
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
    const run = pOf(editor).querySelector('span[style*="font-family"]') as HTMLElement
    expect(run.style.getPropertyValue('--doc-run-ff').replace(/'/g, '"').replace(/,/g, ', ')).toBe(
      run.style.fontFamily,
    )
    editor.destroy()
  })

  it('a declared Latin + East Asian face pair marks the paragraph', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([text('\u4e2d\u6587', { font: 'Microsoft YaHei', fontAscii: 'Calibri' })]),
    })
    expect(pOf(editor).classList.contains('doc-grid-strut')).toBe(true)
    editor.destroy()
  })

  // the strut face carries the document base face's box (doc-style-css), so a
  // single declared face has nothing to union and keeps its own face leading
  // the line: Meiryo UI body lines under a Mincho default sat 2-3px above Word
  it('a single declared face on every slot keeps the family swap and stays unmarked', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([text('\u4e2d\u6587', { font: 'Meiryo UI', fontAscii: 'Meiryo UI' })]),
    })
    const p = pOf(editor)
    expect(p.classList.contains('doc-grid-strut')).toBe(false)
    expect(p.style.fontFamily).not.toBe('')
    expect(
      p.style.getPropertyValue('--doc-grid-strut-tail').replace(/'/g, '"').replace(/,/g, ', '),
    ).toBe(p.style.fontFamily)
    editor.destroy()
  })

  it('all-declared Latin paragraphs stay unmarked', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([text('Latin', { font: 'Arial', fontAscii: 'Arial' })]),
    })
    expect(pOf(editor).classList.contains('doc-grid-strut')).toBe(false)
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
  // the strut box comes from the resolver's hhea split; the canvas stub only
  // keeps the fallback path alive for faces the resolver does not know
  it('emits the metrics @font-face and the strut rule for typed-grid EA docs', () => {
    stubMetricsCanvas(1143, 286)
    mountBlankFace()
    const css = docStyleCss(parsedWith(GRID_SECT, 'ＭＳ 明朝'))
    expect(css).toContain("@font-face { font-family:'GenOffice Grid Strut'")
    expect(css).toContain('ascent-override:100.94%')
    expect(css).toContain('descent-override:29.06%')
    expect(css).toContain('GenOfficePUABlank.woff2')
    // the Office-private MS Mincho renders through a stand-in, so the
    // inheriting .doc-ea-strut paragraphs share the strut too
    expect(css).toContain(
      // !important: all-declared CJK paragraphs carry their chain inline and get the class too
      ".doc-page :is(.doc-grid-strut,.doc-ea-strut) { font-family:var(--doc-strut-face,'GenOffice Grid Strut'),var(--doc-grid-strut-tail,serif) !important }",
    )
    expect(css).toContain('--doc-grid-strut-tail:')
  })

  it('emits no body strut face without an EA face or with a Latin one; the rules stay for paragraph faces', () => {
    stubMetricsCanvas(1143, 286)
    mountBlankFace()
    for (const css of [
      docStyleCss(parsedWith(GRID_SECT)),
      docStyleCss(parsedWith(null, 'Century')),
    ]) {
      expect(css).not.toContain("@font-face { font-family:'GenOffice Grid Strut'")
      expect(css).not.toContain('.doc-ea-strut')
      expect(css).toContain(
        ".doc-page :is(.doc-grid-strut) { font-family:var(--doc-strut-face,'GenOffice Grid Strut'),var(--doc-grid-strut-tail,serif) !important }",
      )
      expect(css).toContain('--doc-grid-strut-tail:')
    }
  })

  it('a missing EA face rendered by a stand-in gets the strut without a grid', () => {
    // Yu Mincho body on a Latin strut: every CJK line unioned ~0.17em past its
    // 1.44 x 1.08 line-height (18.9pt for Word's 17.0 at 11pt, +4 pages)
    stubMetricsCanvas(880, 120)
    mountBlankFace()
    const css = docStyleCss(parsedWith(null, '\u6e38\u660e\u671d'))
    expect(css).toContain("@font-face { font-family:'GenOffice Grid Strut'")
    expect(css).toContain('ascent-override:104.52%')
    expect(css).toContain('.doc-page :is(.doc-grid-strut,.doc-ea-strut) {')
    // the CJK spans lead with the strut face too: their loaded range alias is
    // not the face the strut measured, and offset boxes union past the grid
    expect(css).toContain(
      ".doc-page :is(.doc-grid-strut,.doc-ea-strut) .doc-east-asian-font { font-family:var(--doc-strut-face,'GenOffice Grid Strut'),",
    )
    // balanced spaces too: their 0.8/0.2 one-glyph face unions a Gulim 12pt line
    // under a Malgun strut to 17.1pt for Word's 15.6
    expect(css).toContain(
      ".doc-page :is(.doc-grid-strut,.doc-ea-strut) .doc-balanced-space { font-family:var(--doc-strut-face,'GenOffice Grid Strut'),'GenOffice Balanced Space' !important }",
    )
    expect(docStyleCss(parsedWith(null, '\uff2d\uff33 \u660e\u671d'))).toContain('.doc-ea-strut')
  })

  it('an empty theme slot resolved by the language still drives the grid strut', () => {
    stubMetricsCanvas(952, 351)
    mountBlankFace()
    const resolved = parsedWith(GRID_SECT, 'PMingLiU', { eaSlotEmpty: true, eaFromLang: true })
    expect(docStyleCss(resolved)).toContain('ascent-override:95.08%')
    const unresolved = parsedWith(GRID_SECT, 'Yu Gothic', { eaSlotEmpty: true })
    expect(docStyleCss(unresolved)).not.toContain("@font-face { font-family:'GenOffice Grid Strut'")
  })

  it('Office-private faces under a metric alias get the strut without a grid', () => {
    stubMetricsCanvas(969, 391)
    mountBlankFace()
    const css = docStyleCss(parsedWith(null, 'DengXian'))
    expect(css).toContain("@font-face { font-family:'GenOffice Grid Strut'")
    expect(css).toContain('ascent-override:96.64%')
    expect(css).toContain('descent-override:38.82%')
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

  it('the resolver box needs no canvas; an unknown face without canvas degrades silently', () => {
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null)
    mountBlankFace()
    expect(docStyleCss(parsedWith(GRID_SECT, '\uff2d\uff33 \u30b4\u30b7\u30c3\u30af'))).toContain(
      'ascent-override:100.94%',
    )
  })
})

describe('per-paragraph strut face (strut-face.ts)', () => {
  const fontsDesc = Object.getOwnPropertyDescriptor(document, 'fonts')
  afterEach(() => {
    resetStrutFaces()
    if (fontsDesc) Object.defineProperty(document, 'fonts', fontsDesc)
    else delete (document as { fonts?: unknown }).fonts
  })

  function stubFontFaces() {
    const added: Array<{ family: string; src: string; descriptors: Record<string, string> }> = []
    class FakeFontFace {
      family: string
      constructor(family: string, src: string, descriptors: Record<string, string>) {
        this.family = family
        added.push({ family, src, descriptors })
      }
      load() {
        return Promise.resolve(this)
      }
    }
    vi.stubGlobal('FontFace', FakeFontFace)
    Object.defineProperty(document, 'fonts', {
      value: {
        add: vi.fn(),
        check: () => true,
        load: () => Promise.resolve([]),
        addEventListener: () => {},
        removeEventListener: () => {},
      },
      configurable: true,
    })
    return added
  }

  it('registers one glyphless face per metric pair and names it in --doc-strut-face', () => {
    mountBlankFace()
    const added = stubFontFaces()
    // MS Gothic and MS Mincho share Word's hhea (0.8594 / 0.1406) and line factor: one face
    const gothic = declaredStrutFace('ＭＳ ゴシック')
    expect(gothic).toBe('GenOffice Grid Strut 1')
    expect(declaredStrutFace('MS Mincho')).toBe(gothic)
    expect(added).toHaveLength(1)
    expect(added[0].src).toContain('GenOfficePUABlank.woff2')
    expect(added[0].descriptors).toEqual({
      ascentOverride: '100.94%',
      descentOverride: '29.06%',
      lineGapOverride: '0%',
    })
    expect(declaredStrutFace('Yu Mincho')).toBe('GenOffice Grid Strut 2')
    expect(declaredStrutFace('Calibri')).toBe('GenOffice Grid Strut 3')
    expect(document.fonts.add).toHaveBeenCalledTimes(3)

    // a Cambria + MS Gothic paragraph under a Latin-only document: the runs'
    // face leads its strut (Word probe 2026-11-07: every partner pitches 1.3029)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([text('日本語 abc', { font: 'ＭＳ ゴシック', fontAscii: 'Cambria' })]),
    })
    const p = pOf(editor)
    expect(p.classList.contains('doc-grid-strut')).toBe(true)
    expect(p.style.getPropertyValue('--doc-strut-face')).toBe("'GenOffice Grid Strut 1'")
    editor.destroy()
  })

  it('takes the face of the run that pairs two slots, not an earlier single-slot run', () => {
    mountBlankFace()
    stubFontFaces()
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([
        text('\u65e5\u672c', { font: 'Meiryo', fontAscii: 'Meiryo' }),
        text('\u8a9e abc', { font: 'MS Gothic', fontAscii: 'Cambria' }),
      ]),
    })
    const face = pOf(editor).style.getPropertyValue('--doc-strut-face')
    expect(face).toBe(`'${declaredStrutFace('MS Gothic')}'`)
    expect(face).not.toBe(`'${declaredStrutFace('Meiryo')}'`)
    editor.destroy()
  })

  it('stays silent without the FontFace API or the blank source', () => {
    expect(declaredStrutFace('MS Gothic')).toBeNull()
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: para([text('\u65e5\u672c\u8a9e', { font: 'MS Gothic', fontAscii: 'Cambria' })]),
    })
    expect(pOf(editor).style.getPropertyValue('--doc-strut-face')).toBe('')
    editor.destroy()
  })
})
