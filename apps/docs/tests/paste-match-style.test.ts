/**
 * Paste and Match Style (⌥⇧⌘V, Ctrl+Shift+V elsewhere) pastes the clipboard's
 * text literally (#2856): the menu click arms a one-shot plain-paste flag in
 * the renderer and the editor's paste lanes consume it instead of running the
 * Markdown conversion that a plain Ctrl+V of Markdown-looking text still gets.
 * The editor here carries the App paste lanes this fix touches, trimmed to
 * text payloads — the plain-text reroute, the match-style arm and the
 * Markdown conversion lane (including its empty-paragraph wholesale branch).
 */
import { describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { DOMParser as PmDOMParser } from '@tiptap/pm/model'
import { TextSelection } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  armPlainPaste,
  takePlainPaste,
  takeTextReroute,
} from '../src/renderer/editor/paste-options'
import { pasteTextSlice } from '../src/renderer/editor/paste-text'
import { markdownPasteHtml } from '../src/renderer/editor/markdown-paste'

const MD_TEXT = '# 标题\n\n- 一\n- 二'

/** the App editorProps paste lanes this fix touches (cleanPastedHtml is a
 *  no-op on Markdown output, which has no loose-list <p>s to unwrap). The
 *  carets the lanes see in the app are always inside a paragraph's inline
 *  content, so every test pastes from that shape. */
function handlePaste(
  view: EditorView,
  event: { clipboardData: { getData: (t: string) => string } },
) {
  const data = event.clipboardData
  const html = data.getData('text/html')
  const plainReroute = takeTextReroute() || (!html && takePlainPaste())
  if (plainReroute) {
    const plain = data.getData('text/plain')
    if (plain) {
      const slice = pasteTextSlice(plain, view.state.selection.$from, view)
      view.dispatch(view.state.tr.replaceSelection(slice).scrollIntoView())
      return true
    }
  }
  if (!html && data.getData('text/plain')) {
    const markdownHtml = markdownPasteHtml(data.getData('text/plain'))
    if (markdownHtml !== null) {
      const dom = new window.DOMParser().parseFromString(markdownHtml, 'text/html')
      const parser = PmDOMParser.fromSchema(view.state.schema)
      const $from = view.state.selection.$from
      if (
        view.state.selection.empty &&
        $from.parent.isTextblock &&
        $from.parent.content.size === 0 &&
        $from.depth === 1
      ) {
        // the App's wholesale replace of an empty paragraph by block content
        const parsed = parser.parse(dom.body)
        if (parsed.content.childCount > 0) {
          view.dispatch(view.state.tr.replaceWith($from.before(1), $from.after(1), parsed.content))
          return true
        }
      } else {
        view.dispatch(view.state.tr.replaceSelection(parser.parseSlice(dom.body)).scrollIntoView())
        return true
      }
    }
  }
  return false
}

function pasteEvent(html: string, text: string) {
  return { clipboardData: { getData: (type: string) => (type === 'text/html' ? html : text) } }
}

function makeEditor(paragraph: { text?: string | null } = { text: 'host' }) {
  const content: Array<{ type: 'text'; text: string }> = paragraph.text
    ? [{ type: 'text', text: paragraph.text }]
    : []
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content }],
    },
  })
}

function caretAtEnd(editor: Editor): void {
  // inside the last paragraph's inline content: content.size itself is the
  // block boundary, a state a real caret never sits in (there the open slice
  // start stops merging into the host paragraph)
  const pos = editor.state.doc.content.size - 1
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos)))
}

const blockTypes = (editor: Editor) => editor.state.doc.content.content.map((n) => n.type.name)

const hasNode = (editor: Editor, type: string) => {
  let found = false
  editor.state.doc.descendants((node) => {
    if (node.type.name === type) found = true
    return !found
  })
  return found
}

describe('paste and match style pastes literally (#2856)', () => {
  it('inserts Markdown-looking clipboard text as literal characters', () => {
    const editor = makeEditor()
    caretAtEnd(editor)
    armPlainPaste()
    expect(handlePaste(editor.view, pasteEvent('', MD_TEXT))).toBe(true)
    // the whole payload lands as text: the heading line merges into the host
    // paragraph, the list markers stay characters, nothing parses as structure
    expect(editor.state.doc.textContent).toBe('host# 标题- 一- 二')
    expect(blockTypes(editor)).toEqual(['docParagraph', 'docParagraph', 'docParagraph'])
    editor.destroy()
  })

  it('a plain Ctrl+V of the same text still converts Markdown', () => {
    const editor = makeEditor({ text: null })
    caretAtEnd(editor)
    expect(takePlainPaste()).toBe(false)
    expect(handlePaste(editor.view, pasteEvent('', MD_TEXT))).toBe(true)
    expect(blockTypes(editor)).toContain('docHeading')
    expect(hasNode(editor, 'docListItem')).toBe(true)
    editor.destroy()
  })

  it('an expired flag leaves the Markdown lane in charge', () => {
    vi.useFakeTimers()
    try {
      const editor = makeEditor({ text: null })
      caretAtEnd(editor)
      armPlainPaste()
      vi.advanceTimersByTime(501)
      expect(handlePaste(editor.view, pasteEvent('', MD_TEXT))).toBe(true)
      expect(hasNode(editor, 'docHeading')).toBe(true)
      editor.destroy()
    } finally {
      vi.useRealTimers()
    }
  })

  it('an armed flag is untouched by an HTML paste', () => {
    const editor = makeEditor()
    caretAtEnd(editor)
    armPlainPaste()
    const html = `<meta charset='utf-8'><h1 style="font-family: Arial">web</h1>`
    // an HTML payload never takes the plain lanes, so the flag stays armed
    expect(handlePaste(editor.view, pasteEvent(html, ''))).toBe(false)
    expect(takePlainPaste()).toBe(true)
    editor.destroy()
  })
})
