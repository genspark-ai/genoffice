import { describe, afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import StarterKit from '@tiptap/starter-kit'
import { Redaction } from '../src/renderer/editor/Redaction'
import { SelectiveEscapeMarkdown } from '../src/renderer/editor/markdownEscape'

/**
 * The span has to survive a round trip through the file. The parser treats a
 * bare <span> as an inert tag and keeps only the words, so without the inline
 * patch in editor/inlineTokens.ts the mark silently disappears on reopen and
 * the document goes back to being fully readable by a model.
 */

const SECRET = '13800138000'

/**
 * Every editor this file builds, torn down after each case.
 *
 * An EditorView left alive keeps ProseMirror's DOMObserver on a timer. Once the
 * file finishes, that timer fires against a jsdom that is already gone, and
 * the `document is not defined` it throws lands as an unhandled error — which
 * fails the whole run even when every assertion passed.
 */
const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

function makeEditor(md = '') {
  const editor = new Editor({
    extensions: [StarterKit, SelectiveEscapeMarkdown, Redaction],
  })
  if (md) editor.commands.setContent(md, { contentType: 'markdown' })
  editors.add(editor)
  return editor
}

function selectText(editor: Editor, needle: string) {
  let start = -1
  editor.state.doc.descendants((node, pos) => {
    if (start >= 0) return
    const at = node.textContent.indexOf(needle)
    if (at >= 0) start = pos + 1 + at
  })
  expect(start).toBeGreaterThan(0)
  editor.view.dispatch(
    editor.state.tr.setSelection(
      TextSelection.create(editor.state.doc, start, start + needle.length),
    ),
  )
}

function redactCount(editor: Editor): number {
  let n = 0
  editor.state.doc.descendants((node) => {
    if (node.marks.some((m) => m.type.name === 'redaction')) n++
  })
  return n
}

function labelsIn(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.descendants((node) => {
    const mark = node.marks.find((m) => m.type.name === 'redaction')
    if (mark) out.push(String((mark.attrs as { label?: unknown }).label ?? ''))
  })
  return out
}

describe('a withheld span survives the file', () => {
  it('writes the reader’s words inside a marked span', () => {
    const editor = makeEditor(`please call ${SECRET} confirm the order`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('client phone')
    const saved = editor.getMarkdown()
    expect(saved).toContain(SECRET)
    expect(saved).toContain('data-redaction')
    expect(saved).toContain('data-label="client phone"')
  })

  it('comes back marked when the file is reopened', () => {
    const editor = makeEditor(`please call ${SECRET} confirm the order`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('client phone')

    const reopened = makeEditor(editor.getMarkdown())
    expect(redactCount(reopened)).toBe(1)
    expect(labelsIn(reopened)).toEqual(['client phone'])
  })

  it('round-trips byte for byte over a second save', () => {
    // the strongest statement available: opening and saving again changes
    // nothing, so the representation is stable rather than merely readable
    const editor = makeEditor(`please call ${SECRET} confirm the order`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('client phone')
    const first = editor.getMarkdown()

    const reopened = makeEditor(first)
    expect(reopened.getMarkdown()).toBe(first)
  })

  it('keeps the words, so the reader still has their data', () => {
    const editor = makeEditor(`please call ${SECRET} confirm the order`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('client phone')
    const reopened = makeEditor(editor.getMarkdown())
    expect(reopened.state.doc.textContent).toBe(`please call ${SECRET} confirm the order`)
  })

  it('restores several spans, each with its own label', () => {
    const editor = makeEditor('call 13800138000 or ship to Pudong New Area, Shanghai')
    selectText(editor, '13800138000')
    editor.commands.setRedaction('client phone')
    selectText(editor, 'Pudong New Area, Shanghai')
    editor.commands.setRedaction('shipping address')

    const reopened = makeEditor(editor.getMarkdown())
    expect(redactCount(reopened)).toBe(2)
    expect(labelsIn(reopened).sort()).toEqual(['client phone', 'shipping address'])
    expect(reopened.state.doc.textContent).toBe(
      'call 13800138000 or ship to Pudong New Area, Shanghai',
    )
  })

  it('recovers a span written by hand, attributes in any order', () => {
    const reopened = makeEditor('A <span data-label="phone" data-redaction>13800138000</span> B')
    expect(redactCount(reopened)).toBe(1)
    expect(labelsIn(reopened)).toEqual(['phone'])
  })

  it('falls back to a neutral label when the attribute is missing', () => {
    const reopened = makeEditor('A <span data-redaction>secret</span> B')
    expect(labelsIn(reopened)).toEqual(['private'])
  })

  it('leaves a plain span alone — it is not a withheld span', () => {
    const reopened = makeEditor('A <span>ordinary</span> B')
    expect(redactCount(reopened)).toBe(0)
    expect(reopened.state.doc.textContent).toContain('ordinary')
  })

  it('survives a span that sits at the very start or end of a line', () => {
    for (const src of [
      '<span data-redaction data-label="head">head</span> tail',
      'lead <span data-redaction data-label="tail">tail</span>',
    ]) {
      const reopened = makeEditor(src)
      expect(redactCount(reopened), src).toBe(1)
    }
  })

  it('survives a span that spans two text runs', () => {
    // bold inside the span splits it in the token stream
    const reopened = makeEditor(
      '<span data-redaction data-label="mixed">plain **bold** plain</span>',
    )
    expect(redactCount(reopened)).toBe(1)
    expect(labelsIn(reopened)).toEqual(['mixed'])
  })

  it('does not resurrect a closed-but-unopened span', () => {
    const reopened = makeEditor('A </span> B')
    expect(redactCount(reopened)).toBe(0)
  })

  it('keeps working after the span is removed', () => {
    const editor = makeEditor(`please call ${SECRET} confirm the order`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('client phone')
    editor.commands.unsetRedaction()
    const reopened = makeEditor(editor.getMarkdown())
    expect(redactCount(reopened)).toBe(0)
    expect(reopened.getMarkdown()).toBe(`please call ${SECRET} confirm the order`)
  })
})

describe('the persisted span is robust to markup in the words', () => {
  it('keeps a literal closing tag inside the hidden words hidden on reopen', () => {
    const words = 'token </span> trailing'
    const editor = makeEditor()
    editor.commands.setContent({
      type: 'doc',
      content: [{ type: 'paragraph', content: [{ type: 'text', text: `prefix ${words} suffix` }] }],
    })
    selectText(editor, words)
    editor.commands.setRedaction('secret')
    const saved = editor.getMarkdown()
    expect(saved).toContain('&lt;/span&gt;')
    const reopened = makeEditor(saved)
    expect(reopened.getText()).toContain(words)
    expect(redactCount(reopened)).toBe(1)
    let marked = ''
    reopened.state.doc.descendants((node) => {
      if (node.marks.some((m) => m.type.name === 'redaction')) marked += node.text ?? ''
    })
    expect(marked).toBe(words)
  })

  it('leaves the span markup alone inside a code span', () => {
    const editor = makeEditor('see `<span data-redaction="" data-label="x">hidden</span>` here')
    expect(redactCount(editor)).toBe(0)
    expect(editor.getText()).toContain('<span data-redaction=""')
  })
})
