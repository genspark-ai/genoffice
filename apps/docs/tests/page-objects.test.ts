/** The Selection Pane lists the document's pictures, text boxes and shapes, and selects one. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  isPageObjectSelected,
  pageObjects,
  selectChartForData,
  selectPageObject,
} from '../src/renderer/editor/page-objects'
import { selectedBoxIndex } from '../src/renderer/editor/selected-box'

// the shape of the rule does not matter here: any rule keeps the node out of the list
const RULE = { widthPx: 100, heightPx: 2 }

const box = (text: string, extra: Record<string, unknown> = {}) => ({
  paras: text ? [{ runs: [{ text }] }] : [],
  ...extra,
})

describe('page objects', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'see ' },
              { type: 'docInlineImage', attrs: { dataUrl: 'data:image/png;base64,AA==' } },
              { type: 'docInlineImage', attrs: { rule: RULE } },
            ],
          },
          { type: 'docProtected', attrs: { blockType: 'image', docxIndex: null } },
          {
            type: 'docProtected',
            attrs: {
              docxIndex: null,
              textboxes: [box('A note on the side'), box(''), box('', { picture: true })],
            },
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'end' }] },
        ],
      },
    })
  })
  afterEach(() => editor.destroy())

  it('lists them in document order, numbered within their kind, a horizontal rule left out', () => {
    expect(pageObjects(editor.state.doc).map((o) => [o.kind, o.n, o.index, o.text])).toEqual([
      ['picture', 1, null, ''],
      ['picture', 2, null, ''],
      ['textBox', 1, 0, 'A note on the side'],
      ['shape', 1, 1, ''],
      ['picture', 3, 2, ''],
    ])
  })

  it('selects a picture, and the one box of a node that packs several', () => {
    const objects = pageObjects(editor.state.doc)
    expect(objects.some((o) => isPageObjectSelected(editor.state, o))).toBe(false)
    expect(selectPageObject(editor, objects[0])).toBe(true)
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)
    expect(objects.map((o) => isPageObjectSelected(editor.state, o))).toEqual([
      true,
      false,
      false,
      false,
      false,
    ])
    expect(selectPageObject(editor, objects[3])).toBe(true)
    expect(selectedBoxIndex(editor.state)).toBe(1)
    expect(objects.map((o) => isPageObjectSelected(editor.state, o))).toEqual([
      false,
      false,
      false,
      true,
      false,
    ])
  })

  it('refuses an object the document no longer has there', () => {
    const [first] = pageObjects(editor.state.doc)
    expect(selectPageObject(editor, { ...first, pos: 1 })).toBe(false)
  })

  it('Select Data picks the chart after the caret, keeps one already selected, and reports none', () => {
    expect(selectChartForData(editor)).toBe(false)
    const chart = {
      type: 'docProtected',
      attrs: { docxIndex: null, chartDisplay: { kind: 'bar' } },
    }
    const end = editor.state.doc.content.size
    editor.commands.insertContentAt(end, [chart, chart] as never)
    const charts: number[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.attrs.chartDisplay) charts.push(pos)
    })
    expect(charts).toHaveLength(2)
    editor.commands.setTextSelection(1)
    expect(selectChartForData(editor)).toBe(true)
    expect(editor.state.selection.from).toBe(charts[0])
    // already on a chart: it stays on that one
    editor.view.dispatch(
      editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, charts[1])),
    )
    expect(selectChartForData(editor)).toBe(true)
    expect(editor.state.selection.from).toBe(charts[1])
  })
})
