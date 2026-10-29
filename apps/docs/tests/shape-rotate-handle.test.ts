/**
 * Round rotation handle of a selected shape: dragging the grip above the box
 * turns the shape about its centre (Shift snaps to 15°). The committed angle
 * lives on the model box as rotDeg — the same field the renderer's
 * transform:rotate and the save path (a:xfrm@rot) read.
 */
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { buildShapeParagraphXml, parseDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

/**
 * A gallery rect with textbox content; the explicit posOffset mirrors a
 * dragged shape (a center-aligned drawing re-resolves its wrap on parse).
 */
const SHAPE_XML = buildShapeParagraphXml({ prst: 'rect', withTextbox: true, id: 42 }).replace(
  '<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>',
  '<wp:positionH relativeFrom="column"><wp:posOffset>500000</wp:posOffset></wp:positionH>',
)

async function openShapeDoc(): Promise<Editor> {
  const source = await buildDocx({ bodyXml: SHAPE_XML })
  const parsed = await parseDocx(source)
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

function selectShape(editor: Editor): number {
  let pos = -1
  editor.state.doc.descendants((node, at) => {
    if (pos === -1 && node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
    return true
  })
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
  return pos
}

const boxOf = (editor: Editor, pos: number): TextboxDisplay =>
  (editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[])[0]

/**
 * The node view's chrome: the visual box and the round rotation grip. jsdom has
 * no layout, so pin the box rect (centre 150,150) for the angle math.
 */
function shapeChrome(editor: Editor, pos: number) {
  const wrapper = editor.view.nodeDOM(pos) as HTMLElement
  const boxEl = wrapper.querySelector('.doc-textbox') as HTMLElement
  const handle = wrapper.querySelector('.box-rotate-handle') as HTMLElement
  boxEl.getBoundingClientRect = () => new DOMRect(100, 100, 100, 100)
  return { wrapper, boxEl, handle }
}

function mouseOn(target: EventTarget, type: string, x: number, y: number, shiftKey = false): void {
  target.dispatchEvent(
    new MouseEvent(type, { bubbles: true, button: 0, clientX: x, clientY: y, shiftKey }),
  )
}

describe('shape rotation handle', () => {
  it('turns the shape about its box centre and keeps it selected', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    const { boxEl, handle } = shapeChrome(editor, pos)
    expect(handle).not.toBeNull()

    // pointer due right of the centre = 0°, then due below it = 90° clockwise
    mouseOn(handle, 'mousedown', 200, 150)
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)
    mouseOn(window, 'mousemove', 150, 200)
    expect(boxEl.style.transform).toBe('rotate(90deg)')

    mouseOn(window, 'mouseup', 150, 200)
    expect(boxOf(editor, pos).rotDeg).toBe(90)
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)
    // the drag preview is gone: the inline transform is re-derived from the
    // committed model attrs by the node view (no stale preview angle)
    expect(boxEl.style.transform).toBe('rotate(90deg)')
    editor.destroy()
  })

  it('snaps to 15-degree steps while Shift is held', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    const { boxEl, handle } = shapeChrome(editor, pos)

    mouseOn(handle, 'mousedown', 200, 150)
    // pointer 40° below the horizontal; plain rounding would commit 40
    const rad = (40 * Math.PI) / 180
    const x = 150 + Math.cos(rad) * 100
    const y = 150 + Math.sin(rad) * 100
    mouseOn(window, 'mousemove', x, y, true)
    expect(boxEl.style.transform).toBe('rotate(45deg)')

    mouseOn(window, 'mouseup', x, y, true)
    expect(boxOf(editor, pos).rotDeg).toBe(45)
    editor.destroy()
  })

  it('clears rotDeg when the shape is turned back to 0 degrees', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    // start from a turned shape (as a re-opened rotated doc would)
    editor.commands.updateAttributes('docProtected', {
      textboxes: [{ ...boxOf(editor, pos), rotDeg: 90 }],
    })
    const { boxEl, handle } = shapeChrome(editor, pos)
    expect(boxEl.style.transform).toBe('rotate(90deg)')

    mouseOn(handle, 'mousedown', 200, 150) // 0°
    mouseOn(window, 'mousemove', 150, 100) // -90° from the start angle
    mouseOn(window, 'mouseup', 150, 100)

    expect(boxOf(editor, pos).rotDeg).toBeUndefined()
    // 0 is the cleared state: the model re-render drops the inline transform
    expect(boxEl.style.transform).toBe('')
    editor.destroy()
  })

  it('leaves the shape untouched when the grip is pressed without a turn', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    const { boxEl, handle } = shapeChrome(editor, pos)

    mouseOn(handle, 'mousedown', 200, 150)
    mouseOn(window, 'mousemove', 205, 150) // same ray: no angle change
    mouseOn(window, 'mouseup', 205, 150)

    expect(boxOf(editor, pos).rotDeg).toBeUndefined()
    // the drag preview is cleared and no transaction is committed
    expect(boxEl.style.transform).toBe('')
    editor.destroy()
  })

  it('is inert in a read-only editor', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    editor.setEditable(false)
    const { boxEl, handle } = shapeChrome(editor, pos)

    mouseOn(handle, 'mousedown', 200, 150)
    mouseOn(window, 'mousemove', 150, 200)
    mouseOn(window, 'mouseup', 150, 200)

    expect(boxEl.style.transform).toBe('')
    expect(boxOf(editor, pos).rotDeg).toBeUndefined()
    editor.destroy()
  })
})
