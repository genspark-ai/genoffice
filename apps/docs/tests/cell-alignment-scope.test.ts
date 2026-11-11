import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { buildBlankDocx, parseDocx } from '@genoffice/docx-engine'
import { insertTableAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { enterSelectedTable, setCellAlignment } from '../src/renderer/editor/table-ops'

async function openWithTable() {
  const parsed = await parseDocx(await buildBlankDocx())
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  insertTableAt(editor, 2, 2)
  return editor
}

const cellAligns = (editor: Editor) => {
  const out: Array<[string | null, string | null]> = []
  editor.state.doc.firstChild!.descendants((n) => {
    if (n.type.name === 'docTableCell')
      out.push([n.attrs.vAlign, n.firstChild?.attrs.align ?? null])
  })
  return out
}

describe('cell alignment on a whole-table selection', () => {
  it('applies to every cell once the table selection is widened to its cells', async () => {
    const editor = await openWithTable()
    const { view } = editor
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, 0)))
    expect(enterSelectedTable(view.state, view.dispatch)).toBe(true)
    setCellAlignment('bottom', 'right')(view.state, view.dispatch)
    expect(cellAligns(editor)).toEqual(Array(4).fill(['bottom', 'right']))
  })

  it('a caret in the first cell only touches that cell (the old context-menu path)', async () => {
    const editor = await openWithTable()
    const { view } = editor
    view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(1))))
    setCellAlignment('bottom', 'right')(view.state, view.dispatch)
    expect(cellAligns(editor).filter(([v]) => v === 'bottom')).toHaveLength(1)
  })
})
