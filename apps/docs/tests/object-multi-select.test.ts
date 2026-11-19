import { Editor } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { insertShapeAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setFloatingWrap } from '../src/renderer/editor/floating-z-order'
import {
  isFloatingObject,
  selectedObjectPositions,
  setObjectSelection,
  toggleObjectSelection,
} from '../src/renderer/editor/object-multi-select'

const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

function shapePositions(editor: Editor): number[] {
  const out: number[] = []
  editor.state.doc.forEach((node, pos) => {
    if (node.type.name === 'docProtected' && node.attrs.textboxes) out.push(pos)
  })
  return out
}

/** a document with `count` shapes, the first `floating` of them in front of the text */
async function openShapes(count: number, floating = count): Promise<Editor> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: para('One') + para('Two') }))
  const element = document.createElement('div')
  document.body.appendChild(element)
  const editor = new Editor({
    element,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  for (let i = 0; i < count; i++) insertShapeAt(editor, 'rect', { atPos: 0 })
  shapePositions(editor)
    .slice(0, floating)
    .forEach((pos) => {
      editor.view.dispatch(
        editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)),
      )
      setFloatingWrap(editor, 'front')
    })
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.atEnd(editor.state.doc)))
  return editor
}

describe('selecting several floating objects', () => {
  it('counts only objects that float', async () => {
    const editor = await openShapes(2, 1)
    const [floating, inline] = shapePositions(editor)
    expect(isFloatingObject(editor.state.doc.nodeAt(floating))).toBe(true)
    expect(isFloatingObject(editor.state.doc.nodeAt(inline))).toBe(false)
    expect(toggleObjectSelection(editor.view, inline)).toBe(false)
    editor.destroy()
  })

  it('adds an object to the selection and takes it out again', async () => {
    const editor = await openShapes(3)
    const [a, b, c] = shapePositions(editor)
    expect(selectedObjectPositions(editor.state)).toEqual([])
    toggleObjectSelection(editor.view, a)
    expect(selectedObjectPositions(editor.state)).toEqual([a])
    toggleObjectSelection(editor.view, b)
    toggleObjectSelection(editor.view, c)
    expect(selectedObjectPositions(editor.state)).toEqual([a, b, c])
    // the node selection stays on the first one: commands on "the selected object" still have it
    expect(editor.state.selection.from).toBe(a)
    toggleObjectSelection(editor.view, b)
    expect(selectedObjectPositions(editor.state)).toEqual([a, c])
    // taking the primary one out hands the role to the next
    toggleObjectSelection(editor.view, a)
    expect(selectedObjectPositions(editor.state)).toEqual([c])
    expect(editor.state.selection.from).toBe(c)
    // the last one stays selected
    expect(toggleObjectSelection(editor.view, c)).toBe(false)
    expect(selectedObjectPositions(editor.state)).toEqual([c])
    editor.destroy()
  })

  it('follows the objects through an edit above them and ends when the selection moves on', async () => {
    const editor = await openShapes(2)
    const [a, b] = shapePositions(editor)
    editor.view.dispatch(setObjectSelection(editor.state.tr, [a, b]))
    const insert = editor.schema.nodes.docParagraph.create(null, editor.schema.text('Lead'))
    editor.view.dispatch(editor.state.tr.insert(0, insert))
    expect(selectedObjectPositions(editor.state)).toEqual(shapePositions(editor))
    expect(shapePositions(editor)).toEqual([a + insert.nodeSize, b + insert.nodeSize])

    // an edit to one of them keeps both selected when the edit selects again
    const [a2, b2] = shapePositions(editor)
    const node = editor.state.doc.nodeAt(b2)!
    const tr = editor.state.tr.setNodeMarkup(b2, undefined, {
      ...node.attrs,
      imageOffsetXEmu: 9525,
    })
    editor.view.dispatch(setObjectSelection(tr, [a2, b2]))
    expect(selectedObjectPositions(editor.state)).toEqual([a2, b2])

    // selecting one object the plain way drops the others
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, b2)))
    expect(selectedObjectPositions(editor.state)).toEqual([b2])
    editor.view.dispatch(setObjectSelection(editor.state.tr, [a2, b2]))
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.atEnd(editor.state.doc)))
    expect(selectedObjectPositions(editor.state)).toEqual([])
    editor.destroy()
  })

  it('deletes every selected object with one key', async () => {
    const editor = await openShapes(3)
    const [a, , c] = shapePositions(editor)
    editor.view.dispatch(setObjectSelection(editor.state.tr, [a, c]))
    editor.view.someProp('handleKeyDown', (handle) =>
      handle(editor.view, new KeyboardEvent('keydown', { key: 'Delete' })),
    )
    expect(shapePositions(editor)).toHaveLength(1)
    editor.destroy()
  })

  it('lets go of an object that is deleted', async () => {
    const editor = await openShapes(3)
    const [a, b, c] = shapePositions(editor)
    editor.view.dispatch(setObjectSelection(editor.state.tr, [a, b, c]))
    const size = editor.state.doc.nodeAt(c)!.nodeSize
    const tr = editor.state.tr.delete(c, c + size)
    editor.view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, a)))
    expect(selectedObjectPositions(editor.state)).toEqual([a, b])
    editor.destroy()
  })
})
