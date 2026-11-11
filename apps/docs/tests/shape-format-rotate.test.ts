/**
 * The Shape Format tab's rotation commands and outline-weight picker.
 *
 * Word puts two quarter-turn buttons in the Arrange group and the weight
 * gallery beside Shape Outline. Both write the selected box's display model
 * (rotDeg / borderWidthPx); the save pipeline and the renderer are covered by
 * their own tests, so this mounts the real ribbon and checks the attributes.
 */
import { describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { insertShapeAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

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

/** Drive a shared Dropdown (gs-dd): open the trigger, click the option by value. */
function pickDropdown(container: HTMLElement, dd: HTMLButtonElement, value: string): void {
  act(() => dd.click())
  const item = container.querySelector<HTMLButtonElement>(`.gs-dd-item[data-value="${value}"]`)!
  act(() => item.click())
}

const boxesOf = (editor: Editor, pos: number): TextboxDisplay[] =>
  editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]
const boxOf = (editor: Editor, pos: number): TextboxDisplay => boxesOf(editor, pos)[0]

describe('the Shape Format rotate commands', () => {
  it('turns the selected shape by quarter turns and drops a full turn', async () => {
    const { editor } = await openBlankDoc()
    const pos = selectShape(editor)
    const { container, root } = mountRibbon(editor)

    // a fresh shape is unrotated: no attribute to carry into the saved file
    expect('rotDeg' in boxOf(editor, pos)).toBe(false)

    act(() => tipped(container, 'ribbonRotateRight').click())
    expect(boxOf(editor, pos).rotDeg).toBe(90)

    act(() => tipped(container, 'ribbonRotateRight').click())
    expect(boxOf(editor, pos).rotDeg).toBe(180)

    act(() => tipped(container, 'ribbonRotateLeft').click())
    expect(boxOf(editor, pos).rotDeg).toBe(90)

    // 90 - 90 wraps to 0, and 0 is stored as the absence of the attribute
    act(() => tipped(container, 'ribbonRotateLeft').click())
    expect('rotDeg' in boxOf(editor, pos)).toBe(false)

    // below zero normalizes back into 0..359 (a left turn from 0 is 270)
    act(() => tipped(container, 'ribbonRotateLeft').click())
    expect(boxOf(editor, pos).rotDeg).toBe(270)

    // 270 + 90 is a full turn: clean shape again
    act(() => tipped(container, 'ribbonRotateRight').click())
    expect(boxOf(editor, pos).rotDeg).toBeUndefined()
    expect('rotDeg' in boxOf(editor, pos)).toBe(false)

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  it('leaves the sibling shapes of a shared node alone', async () => {
    // a paragraph anchoring several shapes packs them into one docProtected
    // node; the ribbon reads and writes only the first (the one it shows)
    const { editor } = await openBlankDoc()
    const pos = selectShape(editor)
    const node = editor.state.doc.nodeAt(pos)!
    const first = (node.attrs.textboxes as TextboxDisplay[])[0]
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        textboxes: [first, { ...first, shapeId: '999999' }],
      }),
    )
    const { container, root } = mountRibbon(editor)

    act(() => tipped(container, 'ribbonRotateRight').click())
    expect(boxesOf(editor, pos)[0].rotDeg).toBe(90)
    expect(boxesOf(editor, pos)[1].rotDeg).toBeUndefined()

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  it('is disabled while the document is read-only', async () => {
    const { editor } = await openBlankDoc()
    selectShape(editor)
    editor.setEditable(false)
    const { container, root } = mountRibbon(editor)

    expect(tipped(container, 'ribbonRotateRight').disabled).toBe(true)
    expect(tipped(container, 'ribbonRotateLeft').disabled).toBe(true)
    expect(tipped(container, 'ribbonShapeOutlineWeight').disabled).toBe(true)

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })
})

describe('the Shape Format outline-weight picker', () => {
  it("offers Word's weights in points", async () => {
    const { editor } = await openBlankDoc()
    selectShape(editor)
    const { container, root } = mountRibbon(editor)

    act(() => tipped(container, 'ribbonShapeOutlineWeight').click())
    const items = [...container.querySelectorAll<HTMLButtonElement>('.gs-dd-item')]
    expect(items.map((item) => item.dataset.value)).toEqual([
      '0.25',
      '0.5',
      '0.75',
      '1',
      '1.5',
      '2.25',
      '3',
      '4.5',
      '6',
    ])
    expect(items[0].textContent).toBe(t('ribbonPtValue', { n: '0.25' }))
    expect(items[5].textContent).toBe(t('ribbonPtValue', { n: '2.25' }))

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  it('stores the picked weight as CSS px (1 pt = 4/3 px)', async () => {
    const { editor } = await openBlankDoc()
    const pos = selectShape(editor)
    const { container, root } = mountRibbon(editor)

    // a fresh shape carries no explicit width: Word's own default applies
    expect(boxOf(editor, pos).borderWidthPx).toBeUndefined()
    expect(computeFormatState(editor).shapeBorderWidthPx).toBeNull()

    const weights: Array<[string, number]> = [
      ['0.25', 0.33],
      ['0.5', 0.67],
      ['0.75', 1],
      ['1', 1.33],
      ['1.5', 2],
      ['2.25', 3],
      ['3', 4],
      ['4.5', 6],
      ['6', 8],
    ]
    for (const [pt, px] of weights) {
      pickDropdown(container, tipped(container, 'ribbonShapeOutlineWeight'), pt)
      expect(boxOf(editor, pos).borderWidthPx, `${pt} pt`).toBe(px)
      expect(computeFormatState(editor).shapeBorderWidthPx).toBe(px)
    }

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  it('shows the stored weight on the trigger', async () => {
    const { editor } = await openBlankDoc()
    const pos = selectShape(editor)
    const { container, root, render } = mountRibbon(editor)

    // no explicit width: the trigger names the control instead of reading blank
    expect(container.querySelector('.rb-wrap-dd .gs-dd-value')?.textContent).toBe(
      t('ribbonShapeOutlineWeight'),
    )

    pickDropdown(container, tipped(container, 'ribbonShapeOutlineWeight'), '1.5')
    expect(boxOf(editor, pos).borderWidthPx).toBe(2)

    // the ribbon is fed by computeFormatState: re-render and it reads 1.5 pt
    render()
    expect(container.querySelector('.rb-wrap-dd .gs-dd-value')?.textContent).toBe(
      t('ribbonPtValue', { n: '1.5' }),
    )

    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })
})
