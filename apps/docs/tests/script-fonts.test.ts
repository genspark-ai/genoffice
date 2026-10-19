import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { scriptFontHtml } from '../src/renderer/editor/script-fonts'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

const para = (text: string) => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: text ? [{ type: 'text', text }] : [],
})

function open(paragraphs: string[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: paragraphs.map(para) },
  })
  editors.push(editor)
  return editor
}

const decorated = (editor: Editor) =>
  [...editor.view.dom.querySelectorAll('.doc-east-asian-font')].map((el) => el.textContent)

describe('ScriptFonts decorations', () => {
  it('wraps only the East Asian runs', () => {
    const editor = open(['Report \u62a5\u544a 2026', 'plain', '\u4e2d\u6587'])
    expect(decorated(editor)).toEqual(['\u62a5\u544a', '\u4e2d\u6587'])
  })

  it('tracks an edit in one block without losing the others', () => {
    const editor = open(['a \u7532 b', 'x', 'c \u4e59 d'])
    // typing into the middle block: earlier ranges map through, later ones too
    editor.commands.insertContentAt(editor.state.doc.child(0).nodeSize + 2, '\u4e19')
    expect(decorated(editor)).toEqual(['\u7532', '\u4e19', '\u4e59'])
    // removing the character drops its decoration again
    const pos = editor.state.doc.child(0).nodeSize + 2
    editor.commands.deleteRange({ from: pos, to: pos + 1 })
    expect(decorated(editor)).toEqual(['\u7532', '\u4e59'])
  })

  it('matches a full rebuild after many-block edits', () => {
    const editor = open(Array.from({ length: 6 }, (_, i) => `p${i} \u6587`))
    editor.commands.selectAll()
    editor.commands.insertContent('\u65b0 new')
    expect(decorated(editor)).toEqual(['\u65b0'])
  })
})

describe('ScriptFonts spans stay East-Asian-primary', () => {
  /**
   * The decoration span's first available font is what Chromium uses to place
   * the inline box, and the typed-grid strut is built on the spans sharing the
   * EA-primary half-leading geometry — so the rendered declaration must stay
   * `var(--doc-east-asian-font, inherit)` in both emitters. The IME pre-edit
   * fix is scoped to compositions in styles.css
   * (`.ProseMirror-composing .doc-east-asian-font`), which leaves every line
   * nobody is typing on byte-identical. jsdom resolves neither var() nor font
   * fallback, so the declaration itself is the invariant under test.
   */
  const spanFamily = (editor: Editor): string =>
    (editor.view.dom.querySelector('.doc-east-asian-font') as HTMLElement).style.fontFamily

  const htmlSpanFamily = (html: string): string => {
    const out = document.createElement('div')
    out.innerHTML = scriptFontHtml(html)
    return (out.querySelector('.doc-east-asian-font') as HTMLElement).style.fontFamily
  }

  it('keeps the decoration on the East Asian chain', () => {
    const editor = open(['\u90fd\u4f1a'])
    expect(spanFamily(editor)).toBe('var(--doc-east-asian-font, inherit)')
  })

  it('keeps the serialized-HTML span on the East Asian chain too (no composition in a read-only pane)', () => {
    expect(htmlSpanFamily('<p><span>\u90fd\u4f1a</span></p>')).toBe(
      'var(--doc-east-asian-font, inherit)',
    )
  })
})

describe('ScriptFonts marks the composition window', () => {
  /**
   * styles.css switches the span to the Latin chain only while
   * `.ProseMirror-composing` sits on the editor root. prosemirror-view never
   * sets that class (it toggles `input.composing` and nothing else), so the
   * editor has to, from the composition events. These dispatch the events
   * rather than touching classList: the plumbing is the thing under test.
   */
  const COMPOSING = 'ProseMirror-composing'

  const composing = (editor: Editor): boolean => editor.view.dom.classList.contains(COMPOSING)

  const fire = (editor: Editor, type: 'compositionstart' | 'compositionend'): void => {
    editor.view.dom.dispatchEvent(new CompositionEvent(type, { bubbles: true, data: '' }))
  }

  it('adds the class on compositionstart and drops it on compositionend', () => {
    const editor = open(['\u90fd\u4f1a'])
    expect(composing(editor)).toBe(false)
    fire(editor, 'compositionstart')
    expect(composing(editor)).toBe(true)
    // our handlers return false, so prosemirror-view still flips its own flag —
    // swallowing the event would break view.composing everywhere else
    expect(editor.view.composing).toBe(true)
    fire(editor, 'compositionend')
    expect(composing(editor)).toBe(false)
    expect(editor.view.composing).toBe(false)
  })

  it('drops the class when the editor is destroyed mid-composition', () => {
    const editor = open(['\u90fd\u4f1a'])
    fire(editor, 'compositionstart')
    expect(composing(editor)).toBe(true)
    const dom = editor.view.dom
    editors.splice(editors.indexOf(editor), 1) // already destroyed; skip the afterEach
    editor.destroy()
    expect(dom.classList.contains(COMPOSING)).toBe(false)
  })
})
