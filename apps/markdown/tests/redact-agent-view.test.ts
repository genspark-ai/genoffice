import { describe, afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import StarterKit from '@tiptap/starter-kit'
import { Redaction } from '../src/renderer/editor/Redaction'
import { SelectiveEscapeMarkdown } from '../src/renderer/editor/markdownEscape'
import { buildDocContext } from '../src/renderer/ai/tools'

/**
 * These go through buildDocContext, the function the agent actually calls,
 * rather than through redactJson directly. A test of the helper would still
 * pass if the call site stopped using it — which is the failure that leaks a
 * secret, so the call site is what needs covering.
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

function makeEditor(md: string) {
  const editor = new Editor({
    extensions: [StarterKit, SelectiveEscapeMarkdown, Redaction],
  })
  editor.commands.setContent(md, { contentType: 'markdown' })
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

describe('the agent context is the redacted view', () => {
  it('never contains a withheld number', () => {
    const editor = makeEditor('请拨打 13800138000 确认订单')
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    expect(buildDocContext(editor)).not.toContain(SECRET)
  })

  it('tells the agent the placeholder is there', () => {
    const editor = makeEditor('请拨打 13800138000 确认订单')
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    const context = buildDocContext(editor)
    expect(context).toContain('客户电话')
    expect(context).toContain('{{客户电话}}')
  })

  it('is untouched for a document with nothing withheld', () => {
    const editor = makeEditor('just an ordinary sentence')
    expect(buildDocContext(editor)).toContain('ordinary sentence')
  })

  it('withholds a number sitting in a list', () => {
    const editor = makeEditor('- 联系电话 13800138000')
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    expect(buildDocContext(editor)).not.toContain(SECRET)
  })

  it('withholds an address in a nested block', () => {
    const editor = makeEditor('> 寄到 上海市浦东新区 100 号')
    selectText(editor, '上海市浦东新区')
    editor.commands.setRedaction('收货地址')
    expect(buildDocContext(editor)).not.toContain('上海市浦东新区')
  })

  it('the save path still has the real text, so the reader keeps their data', () => {
    const editor = makeEditor('请拨打 13800138000 确认订单')
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    expect(editor.getMarkdown()).toContain(SECRET)
  })
})
