import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { TableModel } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { tableModelToPmNode } from '../src/renderer/editor/convert'

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function colWidths(model: TableModel): number[] {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [tableModelToPmNode(model)] },
  })
  editors.add(editor)
  return [...editor.getHTML().matchAll(/<col style="width: ?([\d.]+)%;?">/g)].map((m) =>
    Number(m[1]),
  )
}

const cell = (text: string) => ({ paras: [text] })

describe('DocTable colgroup with w:tblCellSpacing', () => {
  it('sizes each column to its gridCol less its share of the gaps', () => {
    // Word: a 2-column table with spacing 15 spans its grid (6281 twips) and the three
    // 30-twip gaps come 45 out of each column; percentages resolve against the table
    // less its border-spacing, so a proportional split shorted the wide column by 2px
    const grid = { colWidthsTwips: [1639, 4642], colWidthsPct: [26.09, 73.91] }
    const widths = colWidths({ rows: [[cell('a'), cell('b')]], ...grid, cellSpacingTwips: 15 })
    const boxes = [1639 - 45, 4642 - 45]
    const total = boxes[0] + boxes[1]
    expect(widths).toHaveLength(2)
    widths.forEach((w, i) => expect(w).toBeCloseTo((boxes[i] / total) * 100, 1))
    // without spacing the grid percentages are used as saved
    expect(colWidths({ rows: [[cell('a'), cell('b')]], ...grid })).toEqual([26.09, 73.91])
  })

  it('charges the outer columns 3s and the inner ones 2s', () => {
    // Word-saved autofit grid of a 4-column table: every cell box is its content
    // plus margins, and the grid holds 45/30/30/45 twips of gap on top of it, so an
    // equal 37.5 share shorts the widest inner column by 7 twips and wraps its line
    const twips = [369, 1947, 3333, 1129]
    const sum = twips.reduce((a, b) => a + b, 0)
    const row = [cell('a'), cell('b'), cell('c'), cell('d')]
    const widths = colWidths({
      rows: [row],
      colWidthsTwips: twips,
      colWidthsPct: twips.map((w) => (w / sum) * 100),
      cellSpacingTwips: 15,
    })
    const boxes = [369 - 45, 1947 - 30, 3333 - 30, 1129 - 45]
    const total = boxes.reduce((a, b) => a + b, 0)
    expect(widths).toHaveLength(4)
    widths.forEach((w, i) => expect(w).toBeCloseTo((boxes[i] / total) * 100, 1))
  })

  it('adds the drawn cell borders to each column box', () => {
    // Word draws the separate cell borders outside the cell box, so the text keeps its
    // full margin inset: a 335-twip column with sz 4 borders still holds "10" at 12 pt
    const twips = [335, 3440, 3190, 2051]
    const sum = twips.reduce((a, b) => a + b, 0)
    const line = { style: 'single', szEighths: 4 }
    const widths = colWidths({
      rows: [[cell('10'), cell('b'), cell('c'), cell('d')]],
      colWidthsTwips: twips,
      colWidthsPct: twips.map((w) => (w / sum) * 100),
      cellSpacingTwips: 15,
      borders: { left: line, right: line, insideV: line },
    })
    const boxes = [335 - 45 + 30, 3440 - 30 + 30, 3190 - 30 + 30, 2051 - 45 + 30]
    const total = boxes.reduce((a, b) => a + b, 0)
    expect(widths).toHaveLength(4)
    widths.forEach((w, i) => expect(w).toBeCloseTo((boxes[i] / total) * 100, 1))
  })
})
