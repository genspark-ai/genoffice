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
import { buildChartDocx, buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setFloatingWrap } from '../src/renderer/editor/floating-z-order'
import { textboxBoxStyle } from '../src/renderer/editor/protected-render'
import { insertShapeAt } from '../src/renderer/components/ribbon-tabs'

/**
 * A gallery roundRect with textbox content (a plain rect stores no prst, and
 * the drag preview only activates on a shape body that carries one); the
 * explicit posOffset mirrors a dragged shape (a center-aligned drawing
 * re-resolves its wrap on parse).
 */
const SHAPE_XML = buildShapeParagraphXml({ prst: 'roundRect', withTextbox: true, id: 42 }).replace(
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
  const turnEl = (wrapper.querySelector('.box-rotator') ?? boxEl) as HTMLElement
  const handle = wrapper.querySelector('.box-rotate-handle') as HTMLElement
  boxEl.getBoundingClientRect = () => new DOMRect(100, 100, 100, 100)
  return { wrapper, boxEl, turnEl, handle }
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
    const { turnEl, handle } = shapeChrome(editor, pos)
    expect(handle).not.toBeNull()

    // pointer due right of the centre = 0°, then due below it = 90° clockwise
    mouseOn(handle, 'mousedown', 200, 150)
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)
    mouseOn(window, 'mousemove', 150, 200)
    expect(turnEl.style.transform).toBe('rotate(90deg)')

    mouseOn(window, 'mouseup', 150, 200)
    expect(boxOf(editor, pos).rotDeg).toBe(90)
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)
    // the drag preview is gone: the inline transform is re-derived from the
    // committed model attrs by the node view (no stale preview angle)
    expect(turnEl.style.transform).toBe('rotate(90deg)')
    editor.destroy()
  })

  it('snaps to 15-degree steps while Shift is held', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    const { turnEl, handle } = shapeChrome(editor, pos)

    mouseOn(handle, 'mousedown', 200, 150)
    // pointer 40° below the horizontal; plain rounding would commit 40
    const rad = (40 * Math.PI) / 180
    const x = 150 + Math.cos(rad) * 100
    const y = 150 + Math.sin(rad) * 100
    mouseOn(window, 'mousemove', x, y, true)
    expect(turnEl.style.transform).toBe('rotate(45deg)')

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
    const { turnEl, handle } = shapeChrome(editor, pos)
    expect(turnEl.style.transform).toBe('rotate(90deg)')

    mouseOn(handle, 'mousedown', 200, 150) // 0°
    mouseOn(window, 'mousemove', 150, 100) // -90° from the start angle
    mouseOn(window, 'mouseup', 150, 100)

    expect(boxOf(editor, pos).rotDeg).toBeUndefined()
    // 0 is the cleared state: the model re-render drops the inline transform
    expect(turnEl.style.transform).toBe('')
    editor.destroy()
  })

  it('leaves the shape untouched when the grip is pressed without a turn', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    const { turnEl, handle } = shapeChrome(editor, pos)

    mouseOn(handle, 'mousedown', 200, 150)
    mouseOn(window, 'mousemove', 205, 150) // same ray: no angle change
    mouseOn(window, 'mouseup', 205, 150)

    expect(boxOf(editor, pos).rotDeg).toBeUndefined()
    // the drag preview is cleared and no transaction is committed
    expect(turnEl.style.transform).toBe('')
    editor.destroy()
  })

  it('is inert in a read-only editor', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    editor.setEditable(false)
    const { turnEl, handle } = shapeChrome(editor, pos)

    mouseOn(handle, 'mousedown', 200, 150)
    mouseOn(window, 'mousemove', 150, 200)
    mouseOn(window, 'mouseup', 150, 200)

    expect(turnEl.style.transform).toBe('')
    expect(boxOf(editor, pos).rotDeg).toBeUndefined()
    editor.destroy()
  })

  it('turns the chrome rotator, not the box, and flags the turned frame', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    editor.commands.updateAttributes('docProtected', {
      textboxes: [{ ...boxOf(editor, pos), rotDeg: 30 }],
    })
    const { wrapper, boxEl, turnEl } = shapeChrome(editor, pos)
    // the rotation lives on the chrome rotator (handles inside it turn along,
    // Word's frame-follows-shape behavior); the box itself stays unturned in
    // the editor and carries the flag that swaps the frame to the box outline
    expect(turnEl.style.transform).toBe('rotate(30deg)')
    expect(turnEl.querySelector('.box-resize-handle')).not.toBeNull()
    expect(turnEl.querySelector('.box-rotate-handle')).not.toBeNull()
    expect(boxEl.style.transform).toBe('')
    expect(wrapper.classList.contains('has-turned-box')).toBe(true)
    // back to 0: the rotator empties and the flag drops (wrapper frame returns)
    editor.commands.updateAttributes('docProtected', {
      textboxes: [{ ...boxOf(editor, pos), rotDeg: undefined }],
    })
    expect(turnEl.style.transform).toBe('')
    expect(wrapper.classList.contains('has-turned-box')).toBe(false)
    editor.destroy()
  })

  it('mirrors a floating box slot on the rotator so the turn pivots on its centre', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    // a floating (wrap none/front) box positions itself absolutely; the
    // rotator must take that slot over, or transform-origin resolves against
    // the zero-size anchor point and the shape orbits the column
    setFloatingWrap(editor, 'front')
    const wrapper = editor.view.nodeDOM(pos) as HTMLElement
    const boxEl = wrapper.querySelector('.doc-textbox') as HTMLElement
    const rotator = wrapper.querySelector('.box-rotator') as HTMLElement
    const box = boxOf(editor, pos)
    expect(box.floating).toBe(true)

    // same left/top slot as the box's own (chrome-free) rendering
    const reference = document.createElement('div')
    reference.setAttribute('style', textboxBoxStyle(box))
    expect(rotator.style.position).toBe('absolute')
    expect(rotator.style.left).toBe(reference.style.left)
    expect(rotator.style.top).toBe(reference.style.top)
    // auto-sized by the static box inside: the rotator rect is the box rect
    expect(rotator.style.width).toBe('')
    expect(rotator.style.height).toBe('')
    expect(boxEl.style.position).toBe('')
    expect(boxEl.style.left).toBe('')
    expect(boxEl.style.top).toBe('')
    expect(boxEl.style.width).toBe('189px')
    // the measurement/obstacle passes address the box through the rotator
    expect(wrapper.querySelectorAll(':scope > .box-rotator > .doc-textbox')).toHaveLength(1)
    // both grips ride inside the turning layer, so the frame turns with them
    expect(rotator.querySelector(':scope > .box-resize-handle')).not.toBeNull()
    expect(rotator.querySelector(':scope > .box-rotate-handle')).not.toBeNull()

    // a turn spins the rotator and keeps its slot; the box itself stays flat
    editor.commands.updateAttributes('docProtected', {
      textboxes: [{ ...box, rotDeg: 30 }],
    })
    expect(rotator.style.transform).toBe('rotate(30deg)')
    expect(rotator.style.position).toBe('absolute')
    expect(rotator.style.left).toBe(reference.style.left)
    expect(rotator.style.top).toBe(reference.style.top)
    expect(boxEl.style.transform).toBe('')
    editor.destroy()
  })

  it('keeps an inline rotator in flow and turns it without dropping its slot', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    insertShapeAt(editor, 'rect')
    let pos = -1
    editor.state.doc.descendants((node, at) => {
      if (pos === -1 && node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
      return true
    })
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    const wrapper = editor.view.nodeDOM(pos) as HTMLElement
    const rotator = wrapper.querySelector(':scope > .box-rotator') as HTMLElement
    const boxEl = rotator.querySelector(':scope > .doc-textbox') as HTMLElement
    const box = boxOf(editor, pos)
    // an inline shape keeps no float slot inline: the stylesheet shrink-wraps
    // the layer to the box, so its rect (and the pivot) is the box rect
    expect(box.floating).toBeUndefined()
    expect(rotator.style.position).toBe('')
    expect(rotator.style.left).toBe('')
    expect(rotator.style.top).toBe('')
    expect(rotator.style.width).toBe('')
    expect(rotator.style.height).toBe('')
    expect(boxEl.style.width).toBe(`${box.widthPx}px`)
    expect(wrapper.querySelectorAll(':scope > .box-rotator > .doc-textbox')).toHaveLength(1)

    // a turn spins only the layer and keeps the inline (slot-less) style; the
    // flow margins of the box still sit inside it
    editor.commands.updateAttributes('docProtected', {
      textboxes: [{ ...box, rotDeg: 45 }],
    })
    expect(rotator.style.transform).toBe('rotate(45deg)')
    expect(rotator.style.position).toBe('')
    expect(rotator.style.left).toBe('')
    expect(wrapper.classList.contains('has-turned-box')).toBe(true)
    expect(boxEl.style.transform).toBe('')

    editor.commands.updateAttributes('docProtected', {
      textboxes: [{ ...boxOf(editor, pos), rotDeg: undefined }],
    })
    expect(rotator.style.transform).toBe('')
    expect(wrapper.classList.contains('has-turned-box')).toBe(false)
    editor.destroy()
  })

  it('the move-drag preview translates the rotator (screen space), not the box inside it', async () => {
    const editor = await openShapeDoc()
    const pos = selectShape(editor)
    const box = boxOf(editor, pos)
    // turn the shape first: the live preview must slide along screen axes —
    // on the box inside the rotated layer a rightward drag moved it down
    editor.commands.updateAttributes('docProtected', {
      textboxes: [{ ...box, rotDeg: 90 }],
    })
    const wrapper = editor.view.nodeDOM(pos) as HTMLElement
    const turnedBox = wrapper.querySelector('.doc-textbox') as HTMLElement
    const rotator = wrapper.querySelector('.box-rotator') as HTMLElement
    expect(rotator.style.transform).toBe('rotate(90deg)')
    expect(turnedBox.style.transform).toBe('')

    turnedBox.dispatchEvent(
      new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: 100, clientY: 100 }),
    )
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: 140, clientY: 100 }))
    // the translate lands on the rotator, prepended so it moves in screen
    // space ahead of the rotation (same target the rotate plugin's turnEl uses)
    expect(rotator.style.transform).toBe('translate(40px, 0px) rotate(90deg)')
    expect(turnedBox.style.transform).toBe('')
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: 140, clientY: 100 }))
    expect(rotator.style.transform).toBe('rotate(90deg)')
    editor.destroy()
  })

  it('keeps the chart resize handle a direct wrapper child (charts have no rotator)', async () => {
    const parsed = await parseDocx(await buildChartDocx())
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const chart = editor.view.dom.querySelector('.doc-protected-chart') as HTMLElement
    expect(chart).not.toBeNull()
    // no rotator layer on charts: the stylesheet must reveal a direct child
    expect(chart.querySelector('.box-rotator')).toBeNull()
    expect(chart.querySelector(':scope > .box-resize-handle')).not.toBeNull()
    editor.destroy()
  })
})
