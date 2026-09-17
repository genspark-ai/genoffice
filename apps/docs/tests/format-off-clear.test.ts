/**
 * Toggling bold/italic must retire the imported Word off-switch (task#426):
 * runs parsed from explicit w:b/w:i val=0 carry docTextStyle boldOff/italicOff
 * (painted font-weight/style:normal). The off-switch span renders inside the
 * strong/em, so without clearing it a bold toggle changes the model and the
 * saved file but never the pixels — the ribbon lights, the reopened document
 * is bold, and the live text stays regular (classic in table cells, where
 * Word writes b=0 against table-style bold).
 */
import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

// an undestroyed view leaves ProseMirror's 20 ms observer flush behind, which
// then runs against the torn-down jsdom document
const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

const offEditor = (attrs: Record<string, unknown>) =>
  track(
    new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [{ type: 'text', text: 'label', marks: [{ type: 'docTextStyle', attrs }] }],
          },
        ],
      } as never,
    }),
  )

const track = (editor: Editor): Editor => {
  editors.push(editor)
  return editor
}

const textStyleOf = (
  marks: readonly { type: { name: string }; attrs: Record<string, unknown> }[],
) => marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs

describe('format off-switch clearing on toggle (task#426)', () => {
  it('typing after Ctrl+B on a boldOff run produces bold text without the off-switch', () => {
    const editor = offEditor({ boldOff: true, font: 'Manrope' })
    editor.commands.setTextSelection(6) // caret after "label"
    editor.commands.toggleMark('bold')
    editor.view.dispatch(editor.state.tr.insertText('typed'))
    const typed = editor.state.doc.firstChild!.lastChild!
    expect(typed.text).toContain('typed')
    expect(typed.marks.some((m) => m.type.name === 'bold')).toBe(true)
    const style = textStyleOf(typed.marks)!
    expect(style.boldOff).toBe(false)
    expect(style.font).toBe('Manrope') // only the off-switch is dropped
  })

  it('bolding a selected boldOff run clears the off-switch in place', () => {
    const editor = offEditor({ boldOff: true })
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.commands.toggleMark('bold')
    const node = editor.state.doc.firstChild!.firstChild!
    expect(node.marks.some((m) => m.type.name === 'bold')).toBe(true)
    expect(textStyleOf(node.marks)!.boldOff).toBe(false)
  })

  it('italic toggle clears italicOff the same way', () => {
    const editor = offEditor({ italicOff: true })
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.commands.toggleMark('italic')
    const node = editor.state.doc.firstChild!.firstChild!
    expect(node.marks.some((m) => m.type.name === 'italic')).toBe(true)
    expect(textStyleOf(node.marks)!.italicOff).toBe(false)
  })

  it('un-bolding restores the off-switch so inherited style bold does not paint', () => {
    const editor = offEditor({ boldOff: true })
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.commands.toggleMark('bold')
    editor.commands.toggleMark('bold')
    const node = editor.state.doc.firstChild!.firstChild!
    expect(node.marks.some((m) => m.type.name === 'bold')).toBe(false)
    expect(textStyleOf(node.marks)!.boldOff).toBe(true)
  })

  it('caret toggle on then off types text with the off-switch back on', () => {
    const editor = offEditor({ boldOff: true })
    editor.commands.setTextSelection(6)
    editor.commands.toggleMark('bold')
    editor.commands.toggleMark('bold')
    editor.view.dispatch(editor.state.tr.insertText('typed'))
    const typed = editor.state.doc.firstChild!.lastChild!
    expect(typed.text).toContain('typed')
    expect(typed.marks.some((m) => m.type.name === 'bold')).toBe(false)
    expect(textStyleOf(typed.marks)!.boldOff).toBe(true)
  })

  it('editing without a bold toggle keeps the off-switch (untouched runs stay Word-faithful)', () => {
    const editor = offEditor({ boldOff: true })
    editor.commands.setTextSelection(6)
    editor.view.dispatch(editor.state.tr.insertText('typed'))
    const node = editor.state.doc.firstChild!.firstChild!
    expect(textStyleOf(node.marks)!.boldOff).toBe(true)
  })
})
