/**
 * Word zeroes HTML auto spacing (beforeAutospacing/afterAutospacing) only
 * between adjacent items of the same list definition: same abstractNum, any
 * ilvl, numIds sharing the abstractNum included; a different list keeps the
 * 14pt (Word for Mac probe 2026-09-24). The collapse rules are keyed on the
 * data-num attribute renderHTML puts on every .doc-li.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { NumberingDef, ParsedDocFull } from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedWith(
  docDefaults: Record<string, unknown>,
  lists: Record<string, string>,
  styles = new Map(),
): ParsedDocFull {
  const numbering = new Map<string, NumberingDef>()
  for (const [numId, abstractNumId] of Object.entries(lists)) {
    numbering.set(numId, { numId, abstractNumId, levels: new Map() } as unknown as NumberingDef)
  }
  return { styles, docDefaults, blocks: [], numbering } as unknown as ParsedDocFull
}

describe('docStyleCss list auto-spacing collapse', () => {
  it('collapses default-level auto spacing only inside one list definition', () => {
    const css = docStyleCss(
      parsedWith({ spaceBeforeAuto: true, spaceAfterAuto: true }, { '1': '0', '3': '0', '2': '1' }),
    )
    const same = ':is([data-num="1"],[data-num="3"])'
    expect(css).toContain(
      `.doc-page .doc-li${same}:not([data-style]):has(+ .doc-li${same}) { margin-bottom:0 }`,
    )
    expect(css).toContain(
      `.doc-page .doc-li${same} + .doc-li${same}:not([data-style]) { margin-top:0 }`,
    )
    expect(css).toContain(
      '.doc-page .doc-li[data-num="2"]:not([data-style]):has(+ .doc-li[data-num="2"]) { margin-bottom:0 }',
    )
    // no blanket rule between any two items
    expect(css).not.toContain('.doc-li:not([data-style]):has(+ .doc-li)')
    expect(css).not.toContain('.doc-li + .doc-li:not([data-style])')
    // lists created in the editor (numIds unknown at parse) collapse among themselves
    const fresh = ':not([data-num="1"],[data-num="3"],[data-num="2"])'
    expect(css).toContain(
      `.doc-page .doc-li${fresh}:not([data-style]):has(+ .doc-li${fresh}) { margin-bottom:0 }`,
    )
    expect(css).toContain(
      `.doc-page .doc-li${fresh} + .doc-li${fresh}:not([data-style]) { margin-top:0 }`,
    )
  })

  it('falls back to the blanket collapse when the document has no numbering', () => {
    const css = docStyleCss(parsedWith({ spaceAfterAuto: true }, {}))
    expect(css).toContain('.doc-page .doc-li:not([data-style]):has(+ .doc-li) { margin-bottom:0 }')
    expect(css).toContain(
      '.doc-page .doc-li.sp-auto-a:has(+ .doc-li) { margin-bottom:0 !important }',
    )
  })

  it('emits the direct sp-auto collapse per list and never a blanket one', () => {
    const css = docStyleCss(parsedWith({}, { '5': '2' }))
    expect(css).toContain(
      '.doc-page .doc-li[data-num="5"].sp-auto-a:has(+ .doc-li[data-num="5"]) { margin-bottom:0 !important }',
    )
    expect(css).toContain(
      '.doc-page .doc-li[data-num="5"] + .doc-li[data-num="5"].sp-auto-b { margin-top:0 !important }',
    )
    expect(css).not.toContain('.doc-li.sp-auto-a:has(+ .doc-li)')
  })

  it('keys a style-level auto spacing collapse on the list too', () => {
    const styles = new Map()
    styles.set('ListParagraph', {
      styleId: 'ListParagraph',
      name: 'List Paragraph',
      type: 'paragraph',
      display: { spaceAfterAuto: true, spaceBeforeAuto: true },
    })
    const css = docStyleCss(parsedWith({}, { '7': '4' }, styles))
    expect(css).toContain(
      '.doc-page .doc-li[data-num="7"][data-style="ListParagraph"]:has(+ .doc-li[data-num="7"]) { margin-bottom:0 }',
    )
    expect(css).toContain(
      '.doc-page .doc-li[data-num="7"] + .doc-li[data-num="7"][data-style="ListParagraph"] { margin-top:0 }',
    )
    expect(css).not.toContain('.doc-li[data-style="ListParagraph"]:has(+ .doc-li)')
  })
})

describe('docListItem renderHTML', () => {
  it('carries the numId as data-num', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docListItem',
            attrs: { kind: 'bullet', numId: '12', ilvl: 1 },
            content: [{ type: 'text', text: 'item' }],
          },
        ],
      },
    })
    const html = editor.getHTML()
    editor.destroy()
    expect(html).toMatch(
      /<div[^>]*class="[^"]*doc-li[^"]*"[^>]*data-num="12"|data-num="12"[^>]*class="[^"]*doc-li/,
    )
  })
})
