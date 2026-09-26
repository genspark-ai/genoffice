import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyCase, nextCaseMode, transformCase } from '../src/renderer/editor/case-transform'

describe('transformCase', () => {
  it('toggles the case of every cased letter', () => {
    expect(transformCase('Hello World 123', 'toggle')).toBe('hELLO wORLD 123')
    expect(transformCase('éÉ ß', 'toggle')).toBe('Éé SS')
    expect(transformCase('\u4e2d\u6587 kana', 'toggle')).toBe('\u4e2d\u6587 KANA')
  })

  it('full-width converts ASCII, the space and half-width katakana', () => {
    expect(transformCase('Abc 1,2!', 'fullWidth')).toBe('Ａｂｃ　１，２！')
    // ka + voiced mark compose; a stray mark becomes the spacing dakuten
    expect(transformCase('ｶﾞﾊﾟｱｰ｡ﾞ', 'fullWidth')).toBe('ガパアー。゛')
    expect(transformCase('\u4e2d\u6587', 'fullWidth')).toBe('\u4e2d\u6587')
  })

  it('half-width is the inverse for ASCII, the space and katakana', () => {
    expect(transformCase('Ａｂｃ　１，２！', 'halfWidth')).toBe('Abc 1,2!')
    expect(transformCase('ガパアー。゛', 'halfWidth')).toBe('ｶﾞﾊﾟｱｰ｡ﾞ')
    const mixed = 'Word 2024 ｶﾞﾊﾟ!'
    expect(transformCase(transformCase(mixed, 'fullWidth'), 'halfWidth')).toBe(mixed)
  })

  it('leaves the Shift+F3 ring at lower / upper / title', () => {
    expect(nextCaseMode('hello')).toBe('upper')
    expect(nextCaseMode('HELLO')).toBe('title')
    expect(nextCaseMode('Hello')).toBe('lower')
  })
})

describe('applyCase', () => {
  it('rewrites only the selected run and keeps the selection', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'Keep ' },
              { type: 'text', text: 'Toggle Me', marks: [{ type: 'bold' }] },
            ],
          },
        ],
      },
    })
    const from = 6
    const to = from + 'Toggle Me'.length
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
    )
    expect(applyCase(editor, 'toggle')).toBe(true)
    expect(editor.state.doc.textContent).toBe('Keep tOGGLE mE')
    expect(editor.state.selection.from).toBe(from)
    expect(editor.state.selection.to).toBe(to)
    expect(editor.state.doc.textBetween(from, to)).toBe('tOGGLE mE')
    let bold = false
    editor.state.doc.nodesBetween(from, to, (node) => {
      if (node.isText) bold = node.marks.some((mark) => mark.type.name === 'bold')
    })
    expect(bold).toBe(true)
    editor.destroy()
  })
})
