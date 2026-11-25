/**
 * The Shape Format tab's Arrange group, the one Layout shows: wrap text plus
 * the stacking commands. The commands themselves are unit-tested in
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
import { SHAPE_INLINE_WRAP } from '../src/renderer/editor/floating-z-order'

// React's act() needs the test environment flag or it warns on every commit
;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

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

/** open one of the group's menu buttons and click the item with this caption */
function pickFrom(
  container: HTMLElement,
  button: Parameters<typeof t>[0],
  item: Parameters<typeof t>[0],
): void {
  act(() => tipped(container, button).click())
  const items = [...container.querySelectorAll<HTMLButtonElement>('[data-rb-panel] button')]
  act(() => items.find((el) => el.textContent?.trim() === t(item))!.click())
}

const attrsOf = (editor: Editor, pos: number) => editor.state.doc.nodeAt(pos)!.attrs

const wrapperClass = (editor: Editor, pos: number): string =>
  (editor.view.nodeDOM(pos) as HTMLElement).className

describe('the Shape Format Arrange group', () => {
  it('is only on screen while a shape is selected', async () => {
    const { editor } = await openBlankDoc()
    const { container, root, render } = mountRibbon(editor)
    expect(tipped(container, 'appBringForward')).toBeNull()

    selectShape(editor)
    // the ribbon switches to Shape Format itself once the shape is selected
    render()
    expect(tipped(container, 'appBringForward')).not.toBeNull()
    // a shape has no preset place, and is turned but not mirrored
    expect(tipped(container, 'ribbonPosition').disabled).toBe(true)
    expect(tipped(container, 'ribbonWrapText').disabled).toBe(false)
    expect(tipped(container, 'ribbonArrangeRotate').disabled).toBe(false)

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  it('applies wrap text to the selected shape', async () => {
    const { editor } = await openBlankDoc()
    const pos = selectShape(editor)
    const { container, root } = mountRibbon(editor)

    pickFrom(container, 'ribbonWrapText', 'appWrapBehind')
    expect(attrsOf(editor, pos).imageWrap).toBe('behind')
    expect(wrapperClass(editor, pos)).toContain('doc-protected-floating')
    expect(wrapperClass(editor, pos)).toContain('doc-protected-behind')

    pickFrom(container, 'ribbonWrapText', 'appWrapSquareLeft')
    expect(attrsOf(editor, pos).imageWrap).toBe('square-left')
    expect(wrapperClass(editor, pos)).toContain('doc-protected-wrapside')

    pickFrom(container, 'ribbonWrapText', 'appWrapInline')
    expect(attrsOf(editor, pos).imageWrap).toBe(SHAPE_INLINE_WRAP)

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  it('runs the stacking commands on the selected shape', async () => {
    const { editor } = await openBlankDoc()
    const pos = selectShape(editor)
    const { container, root } = mountRibbon(editor)

    pickFrom(container, 'appBringForward', 'appBringToFront')
    // a fresh shape is inline; the first reorder floats it in front
    expect(attrsOf(editor, pos).imageWrap).toBe('front')
    expect(attrsOf(editor, pos).imageZOrder).toBe(1)
    expect(wrapperClass(editor, pos)).toContain('doc-protected-floating')

    pickFrom(container, 'appBringForward', 'appBringForward')
    expect(attrsOf(editor, pos).imageZOrder).toBe(2)

    pickFrom(container, 'appSendBackward', 'appSendBackward')
    expect(attrsOf(editor, pos).imageZOrder).toBe(1)

    pickFrom(container, 'appSendBackward', 'appSendToBack')
    expect(attrsOf(editor, pos).imageZOrder).toBe(0)

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })
})
