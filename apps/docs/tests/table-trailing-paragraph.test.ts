import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildBlankDocx, parseDocx } from '@genoffice/docx-engine'
import { insertTableAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions, TABLE_TRAILING_SKIP } from '../src/renderer/editor/extensions'

async function openBlank(content?: unknown) {
  const parsed = await parseDocx(await buildBlankDocx())
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: (content ?? blocksToPmDoc(parsed.blocks)) as never,
  })
}

const topLevel = (editor: Editor) => editor.state.doc.content.content.map((n) => n.type.name)

const cellDepth = (editor: Editor) => {
  const { $from } = editor.state.selection
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type.name === 'docTableCell') return d
  return -1
}

describe('a table is never the last body block (issue #266)', () => {
  it('inserting into an empty paragraph keeps that paragraph below the table, caret in the first cell', async () => {
    const editor = await openBlank()
    insertTableAt(editor, 2, 2)
    expect(topLevel(editor)).toEqual(['docTable', 'docParagraph'])
    expect(cellDepth(editor)).toBeGreaterThan(0)
    expect(editor.state.selection.from).toBe(4)
  })

  it('inserting at the end of a text paragraph appends a paragraph after the table', async () => {
    const editor = await openBlank()
    editor.commands.insertContent('hello')
    insertTableAt(editor, 1, 1)
    expect(topLevel(editor)).toEqual(['docParagraph', 'docTable', 'docParagraph'])
    expect(editor.state.doc.firstChild?.textContent).toBe('hello')
  })

  it('deleting the paragraph below a trailing table restores it', async () => {
    const editor = await openBlank()
    insertTableAt(editor, 1, 1)
    const { doc } = editor.state
    const para = doc.lastChild!
    editor.view.dispatch(editor.state.tr.delete(doc.content.size - para.nodeSize, doc.content.size))
    expect(topLevel(editor)).toEqual(['docTable', 'docParagraph'])
  })

  it('undo removes the table together with its appended paragraph', async () => {
    const editor = await openBlank({
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'hello' }] }],
    })
    editor.commands.setTextSelection(6)
    insertTableAt(editor, 1, 1)
    expect(topLevel(editor)).toEqual(['docParagraph', 'docTable', 'docParagraph'])
    editor.commands.undo()
    expect(topLevel(editor)).toEqual(['docParagraph'])
    expect(editor.state.doc.textContent).toBe('hello')
  })

  it('undo back to a table-last body is left alone and redo still works', async () => {
    const table = {
      type: 'docTable',
      content: [
        {
          type: 'docTableRow',
          content: [{ type: 'docTableCell', content: [{ type: 'docParagraph' }] }],
        },
      ],
    }
    const editor = await openBlank({ type: 'doc', content: [table] })
    editor.commands.insertContentAt(editor.state.doc.content.size, { type: 'docParagraph' })
    expect(topLevel(editor)).toEqual(['docTable', 'docParagraph'])
    editor.commands.undo()
    expect(topLevel(editor)).toEqual(['docTable'])
    editor.commands.redo()
    expect(topLevel(editor)).toEqual(['docTable', 'docParagraph'])
  })

  it('load/stream transactions flagged with TABLE_TRAILING_SKIP are left alone', async () => {
    const editor = await openBlank()
    insertTableAt(editor, 1, 1)
    const { doc } = editor.state
    const para = doc.lastChild!
    const tr = editor.state.tr.delete(doc.content.size - para.nodeSize, doc.content.size)
    tr.setMeta(TABLE_TRAILING_SKIP, true)
    editor.view.dispatch(tr)
    expect(topLevel(editor)).toEqual(['docTable'])
  })

  it('a nested table inserted inside a cell is followed by a paragraph', async () => {
    const editor = await openBlank()
    insertTableAt(editor, 1, 1)
    insertTableAt(editor, 1, 1)
    const cell = editor.state.doc.firstChild!.firstChild!.firstChild!
    expect(cell.content.content.map((n) => n.type.name)).toEqual([
      'docParagraph',
      'docNestedTable',
      'docParagraph',
    ])
  })
})
