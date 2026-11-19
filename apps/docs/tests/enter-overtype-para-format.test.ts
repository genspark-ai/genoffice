/**
 * Replacing whole paragraphs must keep their paragraph FORMATTING (r176
 * follow-up): a selection that consumes complete blocks makes ProseMirror
 * fill the hole with a default paragraph, so a document whose Calibri comes
 * from the paragraph style or the paragraph mark (not explicit run marks)
 * fell back to the theme font — and alignment, indents and spacing were
 * silently lost even in run-formatted documents. Word keeps the formatting
 * of the START of the replaced range, the same rule the run-mark carry
 * (r133) already applies.
 */
import { Editor } from '@tiptap/core'
import { AllSelection, TextSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const calibri = { type: 'docTextStyle', attrs: { font: 'Calibri', fontAscii: 'Calibri' } }

const STYLED = {
  styleId: 'BodyCalibri',
  emptyRunFont: 'Calibri',
  align: 'center',
  indentLeft: 48,
  docxIndex: 7,
  bookmarks: ['keep-out'],
}

const makeEditor = (opts?: { runMarks?: boolean; paras?: number }) => {
  const paras = opts?.paras ?? 2
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: Array.from({ length: paras }, (_v, i) => ({
        type: 'docParagraph',
        attrs: STYLED,
        content: [
          {
            type: 'text',
            text: `ligne ${i + 1} de texte`,
            ...(opts?.runMarks === false ? {} : { marks: [calibri] }),
          },
        ],
      })),
    } as never,
  })
}

const pressEnter = (editor: Editor): boolean => {
  const event = new KeyboardEvent('keydown', {
    key: 'Enter',
    code: 'Enter',
    bubbles: true,
    cancelable: true,
  })
  editor.view.dom.dispatchEvent(event)
  return event.defaultPrevented
}

const selectAllContent = (editor: Editor) => {
  const end = editor.state.doc.content.size - 1
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, end)))
}

const expectFormatKept = (editor: Editor) => {
  expect(editor.state.doc.childCount).toBe(2)
  editor.state.doc.forEach((node) => {
    expect(node.content.size).toBe(0)
    expect(node.attrs.styleId).toBe('BodyCalibri')
    expect(node.attrs.emptyRunFont).toBe('Calibri')
    expect(node.attrs.align).toBe('center')
    expect(node.attrs.indentLeft).toBe(48)
  })
}

describe('Enter over whole paragraphs keeps paragraph formatting (r176 follow-up)', () => {
  it('cross-block full selection: both empty halves keep style, mark font and alignment', () => {
    const editor = makeEditor()
    selectAllContent(editor)
    expect(pressEnter(editor)).toBe(true)
    expectFormatKept(editor)
    // the run-mark carry (r133) still rides along
    const stored = editor.state.storedMarks ?? []
    expect(stored.some((m) => m.type.name === 'docTextStyle' && m.attrs.font === 'Calibri')).toBe(
      true,
    )
    editor.destroy()
  })

  it('Ctrl+A AllSelection: both empty halves keep the paragraph formatting', () => {
    const editor = makeEditor({ paras: 3 })
    editor.view.dispatch(editor.state.tr.setSelection(new AllSelection(editor.state.doc)))
    expect(pressEnter(editor)).toBe(true)
    expectFormatKept(editor)
    // the AllSelection replace built a FRESH paragraph: the format restore
    // must not clone identity/annotation attrs onto it
    editor.state.doc.forEach((node) => {
      expect(node.attrs.docxIndex).toBeNull()
      expect(node.attrs.bookmarks).toBeNull()
    })
    editor.destroy()
  })

  it('style-derived font only (no run marks): paragraph attrs still come back', () => {
    const editor = makeEditor({ runMarks: false })
    selectAllContent(editor)
    pressEnter(editor)
    expectFormatKept(editor)
    editor.destroy()
  })

  it('mid-paragraph selection: surviving halves keep their own attrs untouched', () => {
    const editor = makeEditor()
    // from inside para 1 to inside para 2 — both halves keep text
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 7, 30)),
    )
    pressEnter(editor)
    expect(editor.state.doc.childCount).toBe(2)
    editor.state.doc.forEach((node) => {
      expect(node.content.size).toBeGreaterThan(0)
      expect(node.attrs.styleId).toBe('BodyCalibri')
      expect(node.attrs.align).toBe('center')
    })
    editor.destroy()
  })

  it('plain Delete over whole paragraphs restores the filler paragraph format', () => {
    const editor = makeEditor()
    selectAllContent(editor)
    editor.commands.deleteSelection()
    expect(editor.state.doc.childCount).toBe(1)
    const para = editor.state.doc.child(0)
    expect(para.content.size).toBe(0)
    expect(para.attrs.styleId).toBe('BodyCalibri')
    expect(para.attrs.emptyRunFont).toBe('Calibri')
    expect(para.attrs.align).toBe('center')
    // the r172 mark carry still seeds the pilcrow memory
    expect(String(para.attrs.caretMarks)).toContain('Calibri')
    editor.destroy()
  })

  it('Delete over unformatted-default paragraphs stays a no-op', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'plain one' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'plain two' }] },
        ],
      } as never,
    })
    selectAllContent(editor)
    editor.commands.deleteSelection()
    const para = editor.state.doc.child(0)
    expect(para.attrs.styleId).toBeNull()
    expect(para.attrs.caretMarks).toBeNull()
    editor.destroy()
  })
})
