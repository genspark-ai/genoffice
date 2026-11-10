/**
 * Word pair-kerns text only under w:kern (and only from its half-point size
 * threshold up); Chromium kerns every run by default. Document text therefore
 * renders with font-kerning:none unless the document, style or run asks
 * (SAS batch 2: unkerned Calibri lines 0.7-1.9px over a cell width wrapped in
 * Word but not in GenOffice).
 */
import { describe, expect, it } from 'vitest'
import { DOMSerializer } from '@tiptap/pm/model'
import type { ParsedDocFull, StyleDisplay, StyleInfo, TextboxDisplay } from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { renderTextboxSpec } from '../src/renderer/editor/protected-render'
import { TextStyleMark } from '../src/renderer/editor/marks'
import { fontKerningCss, wordKerns } from '../src/renderer/line-metrics'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedWith(
  docDefaults: ParsedDocFull['docDefaults'],
  styles: Array<[string, StyleDisplay]> = [],
): ParsedDocFull {
  const map = new Map<string, StyleInfo>()
  for (const [styleId, display] of styles) {
    map.set(styleId, { styleId, name: styleId, type: 'paragraph', display } as StyleInfo)
  }
  return { styles: map, docDefaults, blocks: [] } as unknown as ParsedDocFull
}

const DOC_KERN_RULE = '.page-wrap, .doc-page, .pv-page { font-kerning:normal }'

describe('wordKerns', () => {
  it('follows the w:kern threshold against the run size', () => {
    expect(wordKerns(undefined, 22)).toBeUndefined()
    expect(wordKerns(0, 22)).toBe(false)
    expect(wordKerns(2, 22)).toBe(true)
    expect(wordKerns(28, 22)).toBe(false)
    expect(wordKerns(28, 56)).toBe(true)
    expect(wordKerns(28, undefined)).toBe(true)
    expect(fontKerningCss({ kernHalfPoints: 0 })).toBe('none')
    expect(fontKerningCss({ kernHalfPoints: 2, sizeHalfPoints: 20 })).toBe('normal')
    expect(fontKerningCss({})).toBeNull()
  })
})

describe('docStyleCss kerning', () => {
  it('leaves the styles.css font-kerning:none default for documents without w:kern', () => {
    expect(docStyleCss(parsedWith({ sizeHalfPoints: 22 }))).not.toContain('font-kerning')
  })

  it('turns kerning on document-wide for a docDefaults w:kern below the base size', () => {
    expect(docStyleCss(parsedWith({ sizeHalfPoints: 22, kernHalfPoints: 2 }))).toContain(
      DOC_KERN_RULE,
    )
    expect(docStyleCss(parsedWith({ sizeHalfPoints: 22, kernHalfPoints: 28 }))).not.toContain(
      DOC_KERN_RULE,
    )
  })

  it('resolves style-level w:kern against the style size, falling back to the document size', () => {
    const css = docStyleCss(
      parsedWith({ sizeHalfPoints: 22 }, [
        ['Title', { kernHalfPoints: 28, sizeHalfPoints: 56 }],
        ['Small', { kernHalfPoints: 28 }],
        ['Plain', { kernHalfPoints: 0 }],
      ]),
    )
    expect(css).toMatch(/\[data-style="Title"\] \{[^}]*font-kerning:normal/)
    expect(css).toMatch(/\[data-style="Small"\] \{[^}]*font-kerning:none/)
    expect(css).toMatch(/\[data-style="Plain"\] \{[^}]*font-kerning:none/)
  })

  it('re-tests an inherited docDefaults threshold against a style that only sets its size', () => {
    const css = docStyleCss(
      parsedWith({ sizeHalfPoints: 22, kernHalfPoints: 28 }, [
        ['Heading1', { sizeHalfPoints: 32 }],
        ['Body', { sizeHalfPoints: 24 }],
        ['NoSize', { bold: true }],
      ]),
    )
    expect(css).not.toContain(DOC_KERN_RULE)
    expect(css).toMatch(/\[data-style="Heading1"\] \{[^}]*font-kerning:normal/)
    expect(css).toMatch(/\[data-style="Body"\] \{[^}]*font-kerning:none/)
    expect(css).not.toMatch(/\[data-style="NoSize"\] \{[^}]*font-kerning/)
  })
})

describe('docStyleCss table-style kerning', () => {
  it('applies the table style rPr w:kern to cell text', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('Grid', {
      styleId: 'Grid',
      name: 'Grid',
      type: 'table',
      tableDisplay: { wholeTable: { kernHalfPoints: 0 } },
    } as StyleInfo)
    const css = docStyleCss({
      styles,
      docDefaults: { sizeHalfPoints: 22, kernHalfPoints: 2 },
      blocks: [],
    } as unknown as ParsedDocFull)
    expect(css).toContain(DOC_KERN_RULE)
    expect(css).toMatch(/table\[data-tbl-style="Grid"\] td, [^{]*th \{[^}]*font-kerning:none/)
  })
})

describe('run-level kerning', () => {
  it('renders the docTextStyle kern attr as font-kerning', () => {
    const spec = TextStyleMark.config.renderHTML!.call(TextStyleMark as never, {
      mark: { attrs: { kern: true } } as never,
      HTMLAttributes: {},
    }) as [string, Record<string, string>]
    expect(spec[1].style).toContain('font-kerning:normal')
    const off = TextStyleMark.config.renderHTML!.call(TextStyleMark as never, {
      mark: { attrs: { kern: false } } as never,
      HTMLAttributes: {},
    }) as [string, Record<string, string>]
    expect(off[1].style).toContain('font-kerning:none')
  })

  it('emits font-kerning on protected-render runs', () => {
    const box: TextboxDisplay = {
      paras: [
        {
          runs: [{ text: 'A', kernHalfPoints: 2 }, { text: 'B', kernHalfPoints: 0 }, { text: 'C' }],
        },
      ],
    }
    const dom = DOMSerializer.renderSpec(document, renderTextboxSpec(box) as never)
      .dom as HTMLElement
    const spans = [...dom.querySelectorAll('span')].filter((s) => !s.children.length)
    expect(spans[0].style.fontKerning).toBe('normal')
    expect(spans[1].style.fontKerning).toBe('none')
    expect(spans[2].style.fontKerning).toBe('')
  })
})
