/** anchored shapes inside read-only / nested table cells are drawn, not dropped */
import { DOMSerializer } from '@tiptap/pm/model'
import type { TableModel, TextboxDisplay } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

const render = (spec: unknown): HTMLElement =>
  DOMSerializer.renderSpec(document, spec as never).dom as HTMLElement

// a 24x24 white square with a black outline, anchored 11px below its paragraph
const checkbox: TextboxDisplay = {
  paras: [],
  fill: 'FFFFFF',
  borderColor: '000000',
  borderWidthPx: 1,
  widthPx: 24,
  heightPx: 24,
  readOnly: true,
  offsetXEmu: 155575,
  offsetYEmu: 106045,
  floating: true,
}

describe('renderTableSpec anchored cell boxes', () => {
  it('draws the boxes of a nested table cell before their anchor paragraph', () => {
    const nested: TableModel = {
      rows: [
        [
          {
            paras: ['', 'RELEASED', 'REJECTED'],
            anchoredBoxes: [checkbox, { ...checkbox, offsetYEmu: 400000 }],
            anchoredBoxAnchors: [1, 2],
          },
        ],
      ],
    }
    const model: TableModel = {
      rows: [[{ paras: ['outer'], nestedTables: [nested], nestedTableAnchors: [1] }]],
    }
    const dom = render(renderTableSpec(model))
    const struts = dom.querySelectorAll(
      '.doc-nested-table .doc-cell-boxes, .doc-table .doc-table .doc-cell-boxes',
    )
    const all = dom.querySelectorAll('.doc-cell-boxes')
    expect(all.length).toBe(2)
    expect(struts.length + all.length).toBeGreaterThan(0)
    expect(all[0].querySelectorAll('.doc-textbox').length).toBe(1)
    expect(all[0].nextElementSibling?.textContent).toBe('RELEASED')
    expect(all[1].nextElementSibling?.textContent).toBe('REJECTED')
    expect((all[0] as HTMLElement).style.height).toBe('35.1px')
  })

  it('a cell without anchored boxes renders no strut', () => {
    const dom = render(renderTableSpec({ rows: [[{ paras: ['plain'] }]] }))
    expect(dom.querySelectorAll('.doc-cell-boxes').length).toBe(0)
  })
})
