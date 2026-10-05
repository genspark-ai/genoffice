import { describe, expect, it } from 'vitest'
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

function makeEditor(md = '') {
  const editor = new Editor({
    extensions: [StarterKit, SelectiveEscapeMarkdown, Redaction],
  })
  if (md) editor.commands.setContent(md, { contentType: 'markdown' })
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
    const editor = makeEditor(`请拨打 ${SECRET} 确认订单`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    const saved = editor.getMarkdown()
    expect(saved).toContain(SECRET)
    expect(saved).toContain('data-redaction')
    expect(saved).toContain('data-label="客户电话"')
  })

  it('comes back marked when the file is reopened', () => {
    const editor = makeEditor(`请拨打 ${SECRET} 确认订单`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')

    const reopened = makeEditor(editor.getMarkdown())
    expect(redactCount(reopened)).toBe(1)
    expect(labelsIn(reopened)).toEqual(['客户电话'])
  })

  it('round-trips byte for byte over a second save', () => {
    // the strongest statement available: opening and saving again changes
    // nothing, so the representation is stable rather than merely readable
    const editor = makeEditor(`请拨打 ${SECRET} 确认订单`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    const first = editor.getMarkdown()

    const reopened = makeEditor(first)
    expect(reopened.getMarkdown()).toBe(first)
  })

  it('keeps the words, so the reader still has their data', () => {
    const editor = makeEditor(`请拨打 ${SECRET} 确认订单`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    const reopened = makeEditor(editor.getMarkdown())
    expect(reopened.state.doc.textContent).toBe(`请拨打 ${SECRET} 确认订单`)
  })

  it('restores several spans, each with its own label', () => {
    const editor = makeEditor('拨打 13800138000 或寄到 上海市浦东新区')
    selectText(editor, '13800138000')
    editor.commands.setRedaction('客户电话')
    selectText(editor, '上海市浦东新区')
    editor.commands.setRedaction('收货地址')

    const reopened = makeEditor(editor.getMarkdown())
    expect(redactCount(reopened)).toBe(2)
    expect(labelsIn(reopened).sort()).toEqual(['客户电话', '收货地址'])
    expect(reopened.state.doc.textContent).toBe('拨打 13800138000 或寄到 上海市浦东新区')
  })

  it('recovers a span written by hand, attributes in any order', () => {
    const reopened = makeEditor('A <span data-label="电话" data-redaction>13800138000</span> B')
    expect(redactCount(reopened)).toBe(1)
    expect(labelsIn(reopened)).toEqual(['电话'])
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
      '<span data-redaction data-label="首">head</span> tail',
      'lead <span data-redaction data-label="尾">tail</span>',
    ]) {
      const reopened = makeEditor(src)
      expect(redactCount(reopened), src).toBe(1)
    }
  })

  it('survives a span that spans two text runs', () => {
    // bold inside the span splits it in the token stream
    const reopened = makeEditor(
      '<span data-redaction data-label="混排">plain **bold** plain</span>',
    )
    expect(redactCount(reopened)).toBe(1)
    expect(labelsIn(reopened)).toEqual(['混排'])
  })

  it('does not resurrect a closed-but-unopened span', () => {
    const reopened = makeEditor('A </span> B')
    expect(redactCount(reopened)).toBe(0)
  })

  it('keeps working after the span is removed', () => {
    const editor = makeEditor(`请拨打 ${SECRET} 确认订单`)
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    editor.commands.unsetRedaction()
    const reopened = makeEditor(editor.getMarkdown())
    expect(redactCount(reopened)).toBe(0)
    expect(reopened.getMarkdown()).toBe(`请拨打 ${SECRET} 确认订单`)
  })
})
