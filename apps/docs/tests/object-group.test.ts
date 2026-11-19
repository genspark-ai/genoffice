import { Editor } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { insertShapeAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setFloatingWrap } from '../src/renderer/editor/floating-z-order'
import {
  canGroupObjects,
  canUngroupObjects,
  ungroupSelectedObjects,
} from '../src/renderer/editor/object-group'
import {
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

/** a document with `count` floating shapes ahead of two paragraphs */
async function openShapes(count: number): Promise<Editor> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: para('One') + para('Two') }))
  const element = document.createElement('div')
  document.body.appendChild(element)
  const editor = new Editor({
    element,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  for (let i = 0; i < count; i++) insertShapeAt(editor, 'rect', { atPos: 0 })
  for (const pos of shapePositions(editor)) {
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    setFloatingWrap(editor, 'front')
  }
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.atEnd(editor.state.doc)))
  return editor
}

/** what Group writes on its members; the boxes come from the page layout in the app */
function markGroup(editor: Editor, positions: number[], id: string): void {
  const tr = editor.state.tr
  positions.forEach((pos, i) => {
    const node = tr.doc.nodeAt(pos)!
    tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      groupId: id,
      groupBox: { x: i * 200, y: i * 50, w: 189, h: 113 },
    })
  })
  editor.view.dispatch(tr.setSelection(TextSelection.atEnd(tr.doc)))
}

describe('groups of floating objects', () => {
  it('offers Group for several objects and Ungroup for a group', async () => {
    const editor = await openShapes(3)
    const [a, b, c] = shapePositions(editor)
    expect(canGroupObjects(editor.state)).toBe(false)
    editor.view.dispatch(setObjectSelection(editor.state.tr, [a]))
    expect(canGroupObjects(editor.state)).toBe(false)
    editor.view.dispatch(setObjectSelection(editor.state.tr, [a, b]))
    expect(canGroupObjects(editor.state)).toBe(true)
    expect(canUngroupObjects(editor.state)).toBe(false)

    markGroup(editor, [a, b], 'g1')
    editor.view.dispatch(setObjectSelection(editor.state.tr, [a, b]))
    // one group already: nothing to group, something to ungroup
    expect(canGroupObjects(editor.state)).toBe(false)
    expect(canUngroupObjects(editor.state)).toBe(true)
    // a group and one more object can become a bigger group
    editor.view.dispatch(setObjectSelection(editor.state.tr, [a, b, c]))
    expect(canGroupObjects(editor.state)).toBe(true)
    editor.destroy()
  })

  it('selects the whole group when one member is picked', async () => {
    const editor = await openShapes(3)
    const [a, b, c] = shapePositions(editor)
    markGroup(editor, [a, c], 'g1')
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, c)))
    expect(selectedObjectPositions(editor.state)).toEqual([c, a])
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, b)))
    expect(selectedObjectPositions(editor.state)).toEqual([b])
    editor.destroy()
  })

  it('ungroups: the members stay, each on its own', async () => {
    const editor = await openShapes(2)
    const [a, b] = shapePositions(editor)
    markGroup(editor, [a, b], 'g1')
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, a)))
    expect(ungroupSelectedObjects(editor)).toBe(true)
    expect(editor.state.doc.nodeAt(a)!.attrs.groupId).toBeNull()
    expect(editor.state.doc.nodeAt(b)!.attrs.groupBox).toBeNull()
    expect(selectedObjectPositions(editor.state)).toEqual([a, b])
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, b)))
    expect(selectedObjectPositions(editor.state)).toEqual([b])
    editor.destroy()
  })

  it('pastes a copy of a group member as an object of its own', async () => {
    const editor = await openShapes(2)
    const [a, b] = shapePositions(editor)
    markGroup(editor, [a, b], 'g1')
    const copy = editor.state.doc.nodeAt(a)!
    editor.view.dispatch(
      editor.state.tr.insert(editor.state.doc.content.size, copy).setMeta('uiEvent', 'paste'),
    )
    const groups = shapePositions(editor).map((pos) => editor.state.doc.nodeAt(pos)!.attrs.groupId)
    expect(groups).toEqual(['g1', 'g1', null])
    editor.destroy()
  })

  it('pastes copies of several members as a group of their own', async () => {
    const editor = await openShapes(2)
    const [a, b] = shapePositions(editor)
    markGroup(editor, [a, b], 'g1')
    const copies = [editor.state.doc.nodeAt(a)!, editor.state.doc.nodeAt(b)!]
    editor.view.dispatch(
      editor.state.tr.insert(editor.state.doc.content.size, copies).setMeta('uiEvent', 'paste'),
    )
    const groups = shapePositions(editor).map((pos) => editor.state.doc.nodeAt(pos)!.attrs.groupId)
    expect(groups.slice(0, 2)).toEqual(['g1', 'g1'])
    expect(groups[2]).toBeTruthy()
    expect(groups[2]).not.toBe('g1')
    expect(groups[3]).toBe(groups[2])
    editor.destroy()
  })

  it('takes a whole group in and out of a selection on Shift+click', async () => {
    const editor = await openShapes(3)
    const [a, b, c] = shapePositions(editor)
    markGroup(editor, [b, c], 'g1')
    editor.view.dispatch(setObjectSelection(editor.state.tr, [a]))
    toggleObjectSelection(editor.view, c)
    expect(selectedObjectPositions(editor.state)).toEqual([a, c, b])
    toggleObjectSelection(editor.view, b)
    expect(selectedObjectPositions(editor.state)).toEqual([a])
    editor.destroy()
  })

  it('cuts every member of a group, and puts them all on the clipboard', async () => {
    const editor = await openShapes(3)
    const [a, b] = shapePositions(editor)
    markGroup(editor, [a, b], 'g1')
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, a)))
    const data = new Map<string, string>()
    const event = Object.assign(new Event('cut', { bubbles: true, cancelable: true }), {
      clipboardData: {
        clearData: () => data.clear(),
        setData: (type: string, value: string) => data.set(type, value),
      },
    })
    editor.view.dom.dispatchEvent(event)
    expect(event.defaultPrevented).toBe(true)
    expect(shapePositions(editor)).toHaveLength(1)
    // both members went to the clipboard as the editor's own content
    const html = data.get('text/html') ?? ''
    expect(html.match(/doc-protected/g)?.length).toBeGreaterThanOrEqual(2)
    editor.destroy()
  })

  it('leaves a cut to the text editor inside a member', async () => {
    const editor = await openShapes(3)
    const [a, b] = shapePositions(editor)
    markGroup(editor, [a, b], 'g1')
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, a)))
    const inner = document.createElement('div')
    inner.className = 'doc-textbox-editor'
    editor.view.dom.appendChild(inner)
    const data = new Map<string, string>()
    const event = Object.assign(new Event('cut', { bubbles: true, cancelable: true }), {
      clipboardData: {
        clearData: () => data.clear(),
        setData: (type: string, value: string) => data.set(type, value),
      },
    })
    inner.dispatchEvent(event)
    // the group was not taken: at most the one object the editor's own cut handles
    expect(shapePositions(editor).length).toBeGreaterThanOrEqual(2)
    expect((data.get('text/html') ?? '').match(/data-doc-protected/g)?.length ?? 0).toBeLessThan(2)
    editor.destroy()
  })

  it('hands the save one group of the members that are in the document', async () => {
    const editor = await openShapes(3)
    const [a, b, c] = shapePositions(editor)
    markGroup(editor, [a, c], 'g1')
    // a group of one is no group
    markGroup(editor, [b], 'g2')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, [])
    expect(plan.groups).toHaveLength(1)
    expect(plan.groups[0].members).toEqual([
      { block: 0, drawing: 0, xEmu: 0, yEmu: 0, widthEmu: 189 * 9525, heightEmu: 113 * 9525 },
      {
        block: 2,
        drawing: 0,
        xEmu: 200 * 9525,
        yEmu: 50 * 9525,
        widthEmu: 189 * 9525,
        heightEmu: 113 * 9525,
      },
    ])
    expect(plan.carried).toEqual([])
    editor.destroy()
  })
})
