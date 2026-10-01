/**
 * The Shape Format tab's Arrange group: wrap text plus the stacking commands,
 * where Word puts them. The commands themselves are unit-tested in
 * floating-z-order.test.ts; this mounts the real ribbon and checks the group
 * is wired to the selected shape.
 */
import { describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { insertShapeAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

async function openBlankDoc() {
  const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>' })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor }
}

/** insert a rect shape and select it as a single click would */
function selectShape(editor: Editor): number {
  insertShapeAt(editor, 'rect')
  let pos = -1
  editor.state.doc.descendants((node, at) => {
    if (node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
    return true
  })
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
  return pos
}

function mountRibbon(editor: Editor) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  const render = () =>
    act(() => root.render(createElement(Ribbon, ribbonProps(editor, computeFormatState(editor)))))
  render()
  return { container, root, render }
}

/** controls are addressed by their tooltip so the test does not pin a locale */
const tipped = (container: HTMLElement, key: Parameters<typeof t>[0]): HTMLButtonElement =>
  container.querySelector(`.ribbon-body [data-tip="${t(key)}"]`) as HTMLButtonElement

/** Drive a shared Dropdown (gs-dd): open the trigger, click the option by value. */
function pickDropdown(container: HTMLElement, dd: HTMLButtonElement, value: string): void {
  act(() => dd.click())
  const item = container.querySelector<HTMLButtonElement>(`.gs-dd-item[data-value="${value}"]`)!
  act(() => item.click())
}

const attrsOf = (editor: Editor, pos: number) => editor.state.doc.nodeAt(pos)!.attrs

const wrapperClass = (editor: Editor, pos: number): string =>
  (editor.view.nodeDOM(pos) as HTMLElement).className

describe('the Shape Format Arrange group', () => {
  it('is only on screen while a shape is selected', async () => {
    const { editor } = await openBlankDoc()
    const { container, root, render } = mountRibbon(editor)
    expect(tipped(container, 'appBringToFront')).toBeNull()

    selectShape(editor)
    // the ribbon switches to Shape Format itself once the shape is selected
    render()
    expect(tipped(container, 'appBringToFront')).not.toBeNull()

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  it('applies wrap text to the selected shape', async () => {
    const { editor } = await openBlankDoc()
    const pos = selectShape(editor)
    const { container, root } = mountRibbon(editor)

    pickDropdown(container, tipped(container, 'ribbonWrapText'), 'behind')
    expect(attrsOf(editor, pos).imageWrap).toBe('behind')
    expect(wrapperClass(editor, pos)).toContain('doc-protected-floating')
    expect(wrapperClass(editor, pos)).toContain('doc-protected-behind')

    pickDropdown(container, tipped(container, 'ribbonWrapText'), 'square-left')
    expect(attrsOf(editor, pos).imageWrap).toBe('square-left')
    expect(wrapperClass(editor, pos)).toContain('doc-protected-wrapside')

    pickDropdown(container, tipped(container, 'ribbonWrapText'), '')
    expect(attrsOf(editor, pos).imageWrap).toBeNull()

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  it('runs the stacking commands on the selected shape', async () => {
    const { editor } = await openBlankDoc()
    const pos = selectShape(editor)
    const { container, root } = mountRibbon(editor)

    act(() => tipped(container, 'appBringToFront').click())
    // a fresh shape is inline; the first reorder floats it in front
    expect(attrsOf(editor, pos).imageWrap).toBe('front')
    expect(attrsOf(editor, pos).imageZOrder).toBe(1)
    expect(wrapperClass(editor, pos)).toContain('doc-protected-floating')

    act(() => tipped(container, 'appBringForward').click())
    expect(attrsOf(editor, pos).imageZOrder).toBe(2)

    act(() => tipped(container, 'appSendBackward').click())
    expect(attrsOf(editor, pos).imageZOrder).toBe(1)

    act(() => tipped(container, 'appSendToBack').click())
    expect(attrsOf(editor, pos).imageZOrder).toBe(0)

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })
})
