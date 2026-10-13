/** anchored shapes inside read-only / nested table cells are drawn, not dropped */
import { DOMSerializer } from '@tiptap/pm/model'
import type { TableModel, TextboxDisplay } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import {
  cellBoxesSpec,
  renderTableSpec,
  splitCellBoxGroup,
} from '../src/renderer/editor/protected-render'

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

  it('a wrapNone box reserves no strut height', () => {
    const model: TableModel = {
      rows: [
        [
          {
            paras: ['head'],
            anchoredBoxes: [
              { ...checkbox, offsetYEmu: 217920, heightPx: 193, noWrap: true, behind: true },
              // behind is z-order only: a wrapped behind-text box still grows the row
              { ...checkbox, offsetYEmu: 0, heightPx: 30, behind: true },
            ],
            anchoredBoxAnchors: [0],
          },
        ],
      ],
    }
    const strut = render(renderTableSpec(model)).querySelector<HTMLElement>('.doc-cell-boxes')!
    expect(strut.style.height).toBe('30px')
    expect(strut.querySelectorAll('.doc-textbox').length).toBe(2)
  })

  it('layoutInCell="0" page-positioned boxes get their own strut that never grows the row', () => {
    const avatar: TextboxDisplay = {
      ...checkbox,
      widthPx: 88,
      heightPx: 108,
      offsetXEmu: 212090,
      offsetYEmu: 644525,
      outsideCell: true,
      pageRelV: true,
      pageRelVFrom: 'page',
    }
    const bar: TextboxDisplay = {
      ...avatar,
      offsetXEmu: 635,
      offsetYEmu: 12700,
      pageRelX: true,
      pageRelXFrom: 'page',
    }
    const groups = splitCellBoxGroup([avatar, checkbox, bar])
    expect(groups.map((g) => g.length)).toEqual([2, 1])
    const strut = render(cellBoxesSpec(groups[0]))
    expect(strut.className).toBe('doc-cell-boxes doc-cell-boxes-page')
    expect(strut.style.height).toBe('')
    const [a, b] = Array.from(strut.querySelectorAll<HTMLElement>('.doc-textbox'))
    expect(a.dataset.cellPage).toBe('1')
    expect(a.dataset.cellPageX).toBe('22.3')
    expect(a.dataset.cellPageY).toBe('67.7')
    expect(a.dataset.pageRelFrom).toBe('page')
    expect(a.dataset.pageRelX).toBeUndefined()
    expect(b.dataset.pageRelX).toBe('1')
    expect(b.dataset.pageRelXFrom).toBe('page')
    // the raw cell-origin offsets stay: the page pin pass translates from them
    expect(a.style.left).toBe('22.3px')
    expect(a.style.top).toBe('67.7px')
    const inCell = render(cellBoxesSpec(groups[1]))
    expect(inCell.className).toBe('doc-cell-boxes')
    expect(inCell.style.height).toBe('35.1px')
    // read-only table path splits the same way
    const dom = render(
      renderTableSpec({
        rows: [[{ paras: ['x'], anchoredBoxes: [avatar, checkbox], anchoredBoxAnchors: [0, 0] }]],
      }),
    )
    expect(dom.querySelectorAll('.doc-cell-boxes').length).toBe(2)
    expect(dom.querySelectorAll('.doc-cell-boxes-page').length).toBe(1)
  })

  it('a cell without anchored boxes renders no strut', () => {
    const dom = render(renderTableSpec({ rows: [[{ paras: ['plain'] }]] }))
    expect(dom.querySelectorAll('.doc-cell-boxes').length).toBe(0)
  })
})
