import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { eastAsianSegments, scriptFontHtml } from '../src/renderer/editor/script-fonts'

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

  // Word probe 2026-10-22: under balanceSingleByteDoubleByteWidth halfwidth kana
  // take the single-byte grid delta, so they must not share the full-width span
  it('splits halfwidth kana off the full-width runs', () => {
    const editor = open(['\u30ab\u30bf\uff76\uff85\u6f22 x'])
    const spans = [...editor.view.dom.querySelectorAll('.doc-east-asian-font')]
    expect(
      spans.map((el) => [el.textContent, el.classList.contains('doc-east-asian-halfwidth')]),
    ).toEqual([
      ['\u30ab\u30bf', false],
      ['\uff76\uff85', true],
      ['\u6f22', false],
    ])
    expect(eastAsianSegments('a\uff76\u3002\uffa1b')).toEqual([
      { index: 1, text: '\uff76' },
      { index: 2, text: '\u3002' },
      { index: 3, text: '\uffa1' },
    ])
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

describe('ScriptFonts and IME composition', () => {
  const textNodes = (dom: Node): string[] => {
    const walker = document.createTreeWalker(dom, NodeFilter.SHOW_TEXT)
    const out: string[] = []
    for (let n = walker.nextNode(); n; n = walker.nextNode()) out.push((n as Text).data)
    return out
  }

  const startComposition = (editor: Editor) => {
    const view = editor.view
    view.dom.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }))
    return view
  }

  // The decoration spans wrap every East Asian run, so the pinyin a Chinese
  // IME keeps in the DOM splits the run in two. Rebuilding the decorations
  // while the composition runs rewrites the text node the composition lives
  // in: Chrome drops the selection with it, the composition is cancelled and
  // the caret lands past the next bracket.
  it('keeps the composing text node intact while a composition runs', () => {
    const editor = open(['（看各种文件）'])
    const caret = 1 + '（看各种文件'.length
    editor.commands.setTextSelection(caret)
    const view = startComposition(editor)
    expect(view.composing).toBe(true)

    // the provisional pinyin reaches the document while the composition runs
    view.dispatch(view.state.tr.insertText('gezhong', caret))

    // the decoration set was not rebuilt: the run stayed one piece and the
    // composing text is still in the node the composition lives in
    expect(textNodes(view.dom)).toEqual(['（看各种文件gezhong）'])
    expect(decorated(editor)).toEqual(['（看各种文件gezhong）'])
  })

  it('picks the committed characters up once the composition is over', () => {
    const editor = open(['（看各种文件）'])
    const caret = 1 + '（看各种文件'.length
    editor.commands.setTextSelection(caret)
    const view = startComposition(editor)
    view.dispatch(view.state.tr.insertText('gezhong', caret))

    view.dom.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }))
    expect(view.composing).toBe(false)
    // the commit replaces the pinyin with the chosen characters
    view.dispatch(view.state.tr.insertText('各种', caret, caret + 'gezhong'.length))

    expect(decorated(editor)).toEqual(['（看各种文件各种）'])
  })

  // A Japanese or Korean IME (and a raw-pinyin Enter commit) has the final
  // text in the DOM before compositionend, so ProseMirror dispatches nothing
  // and the mapped set stays: the characters appended at the end of the run
  // would keep the Latin face until the next non-IME edit touches the block.
  it('decorates a committed run that no transaction changed', async () => {
    const editor = open(['日本語'])
    const caret = 1 + '日本語'.length
    editor.commands.setTextSelection(caret)
    const view = startComposition(editor)
    view.dispatch(view.state.tr.insertText('にゅうりょく', caret))
    view.dispatch(view.state.tr.insertText('入力', caret, caret + 'にゅうりょく'.length))

    view.dom.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }))
    await new Promise((resolve) => setTimeout(resolve, 50))

    expect(decorated(editor)).toEqual(['日本語入力'])
  })
})

describe('ScriptFonts spans stay East-Asian-primary', () => {
  /**
   * The decoration span's first available font is what Chromium uses to place
   * the inline box, and the typed-grid strut is built on the spans sharing the
   * EA-primary half-leading geometry — so the rendered declaration must stay
   * `var(--doc-east-asian-font, inherit)` in both emitters, and nothing may
   * swap it while an IME composition is open: that moved every line in the
   * document on each keystroke. jsdom resolves neither var() nor font
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
