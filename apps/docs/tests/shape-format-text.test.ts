/**
 * The Shape Format tab's Text group.
 *
 * Selecting a shape gives the main editor a NodeSelection, not a text selection,
 * so a mark command has nothing to apply itself to. Word formats the shape's
 * whole text in that state, and keeps Shape Format on screen once you double-
 * click into the text — without both, changing the weight or color of text in a
 * shape means a trip back to the Home tab.
 */
import { describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { insertShapeAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setActiveSubEditor } from '../src/renderer/editor/active-editor'
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
  return { editor, parsed, source }
}

/** insert a shape, give it text, and select it as a single click would */
function selectShapeWithText(
  editor: Editor,
  paras: TextboxDisplay['paras'],
  patch: Partial<TextboxDisplay> = {},
): number {
  insertShapeAt(editor, 'rect')
  let pos = -1
  editor.state.doc.descendants((node, at) => {
    if (node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
    return true
  })
  const node = editor.state.doc.nodeAt(pos)!
  const boxes = node.attrs.textboxes as TextboxDisplay[]
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      textboxes: [{ ...boxes[0], ...patch, paras }],
    }),
  )
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
  return pos
}

const boxOf = (editor: Editor, pos: number): TextboxDisplay =>
  (editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[])[0]

const runsOf = (editor: Editor, pos: number) => boxOf(editor, pos).paras.flatMap((p) => p.runs)

function mountRibbon(editor: Editor) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  const render = () =>
    act(() => root.render(createElement(Ribbon, ribbonProps(editor, computeFormatState(editor)))))
  render()
  return { container, root, render }
}

/** buttons are addressed by their tooltip so the test does not pin a locale */
const tipped = (container: HTMLElement, key: Parameters<typeof t>[0]): HTMLButtonElement =>
  container.querySelector(`.ribbon-body [data-tip="${t(key)}"]`) as HTMLButtonElement

describe('what the ribbon reads off a selected shape', () => {
  it('reports the formatting only while every run agrees', async () => {
    const { editor } = await openBlankDoc()
    selectShapeWithText(editor, [
      { runs: [{ text: 'bold one', bold: true }], align: 'center' },
      { runs: [{ text: 'plain two' }], align: 'center' },
    ])

    const mixed = computeFormatState(editor)
    expect(mixed.shapeHasText).toBe(true)
    expect(mixed.shapeTextBold).toBe(false) // one run is not bold
    expect(mixed.shapeTextAlign).toBe('center') // both paragraphs are

    selectShapeWithText(editor, [{ runs: [{ text: 'all bold', bold: true }], align: 'right' }])
    const uniform = computeFormatState(editor)
    expect(uniform.shapeTextBold).toBe(true)
    expect(uniform.shapeTextAlign).toBe('right')
    editor.destroy()
  })

  it('shows the shape style color for runs that carry none of their own', async () => {
    const { editor } = await openBlankDoc()
    // a fresh shape's runs are colorless: the white comes from the style's
    // fontRef, so reading the run alone would leave the swatch blank on blue
    selectShapeWithText(editor, [{ runs: [{ text: 'on blue' }] }], { textColor: 'FFFFFF' })
    expect(computeFormatState(editor).shapeTextColor).toBe('FFFFFF')

    selectShapeWithText(editor, [{ runs: [{ text: 'red', color: 'FF0000' }] }])
    expect(computeFormatState(editor).shapeTextColor).toBe('FF0000')
    editor.destroy()
  })

  it('offers nothing to format for a shape with no text', async () => {
    const { editor } = await openBlankDoc()
    selectShapeWithText(editor, [])
    expect(computeFormatState(editor).shapeHasText).toBe(false)
    editor.destroy()
  })
})

describe('the Text group', () => {
  it('bolds the whole shape, because there is no selection to bold instead', async () => {
    const { editor } = await openBlankDoc()
    const pos = selectShapeWithText(editor, [
      { runs: [{ text: 'first ' }, { text: 'second', italic: true }] },
      { runs: [{ text: 'third' }] },
    ])
    const { container, root, render } = mountRibbon(editor)

    act(() => tipped(container, 'ribbonBoldTip').click())

    expect(runsOf(editor, pos)).toHaveLength(3)
    expect(runsOf(editor, pos).every((r) => r.bold === true)).toBe(true)
    expect(runsOf(editor, pos)[1].italic).toBe(true) // a bold press leaves it alone

    render()
    expect(tipped(container, 'ribbonBoldTip').className).toContain('active')

    // pressing again clears the flag rather than storing an explicit "off"
    act(() => tipped(container, 'ribbonBoldTip').click())
    expect(runsOf(editor, pos).some((r) => 'bold' in r)).toBe(false)

    act(() => root.unmount())
    editor.destroy()
  })

  it('leaves the sibling shapes of a shared node alone', async () => {
    // a paragraph anchoring several shapes packs them into one docProtected node,
    // and the ribbon reads and writes only the first — the rest are shapes it is
    // not showing, so reformatting them would be invisible collateral damage
    const { editor } = await openBlankDoc()
    const pos = selectShapeWithText(editor, [{ runs: [{ text: 'shown' }] }])
    const node = editor.state.doc.nodeAt(pos)!
    const first = (node.attrs.textboxes as TextboxDisplay[])[0]
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        textboxes: [first, { ...first, paras: [{ runs: [{ text: 'sibling' }] }] }],
      }),
    )
    const { container, root } = mountRibbon(editor)

    act(() => tipped(container, 'ribbonBoldTip').click())

    const boxes = editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]
    expect(boxes[0].paras[0].runs[0].bold).toBe(true)
    expect(boxes[1].paras[0].runs[0].bold).toBeUndefined()

    act(() => root.unmount())
    editor.destroy()
  })

  it('aligns every paragraph of the shape at once', async () => {
    const { editor } = await openBlankDoc()
    const pos = selectShapeWithText(editor, [
      { runs: [{ text: 'one' }] },
      { runs: [{ text: 'two' }], align: 'left' },
    ])
    const { container, root } = mountRibbon(editor)

    act(() => tipped(container, 'ribbonAlignRightTip').click())
    expect(boxOf(editor, pos).paras.map((p) => p.align)).toEqual(['right', 'right'])

    act(() => root.unmount())
    editor.destroy()
  })

  it('stays on screen while the text inside the shape is being edited', async () => {
    const { editor } = await openBlankDoc()
    selectShapeWithText(editor, [{ runs: [{ text: 'typing here' }] }])
    const { container, root, render } = mountRibbon(editor)
    expect(tipped(container, 'ribbonBoldTip')).not.toBeNull()

    // double-clicking in focuses a sub-editor while the shape stays selected in
    // the main editor — that combination used to hide the whole tab
    const sub = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] },
    })
    setActiveSubEditor(sub)
    try {
      render()
      expect(computeFormatState(editor).sub).toBe(sub)
      expect(tipped(container, 'ribbonBoldTip')).not.toBeNull()
    } finally {
      setActiveSubEditor(null)
      sub.destroy()
    }

    act(() => root.unmount())
    editor.destroy()
  })
})

describe('a shape formatted this way', () => {
  it('carries the bold into the saved file', async () => {
    const { editor, parsed } = await openBlankDoc()
    selectShapeWithText(editor, [{ runs: [{ text: 'Heading in a box' }] }])
    const { container, root } = mountRibbon(editor)
    act(() => tipped(container, 'ribbonBoldTip').click())
    act(() => root.unmount())

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const box = reparsed.blocks.flatMap((b) => b.textboxes ?? []).find((b) => b.paras.length > 0)
    expect(box?.paras[0].runs[0]).toMatchObject({ text: 'Heading in a box', bold: true })
    editor.destroy()
  })
})
