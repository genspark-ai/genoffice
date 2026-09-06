import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  applyLiftTops,
  computeSectionedSlicesF2,
  measureBlocks,
  type BlockBox,
  type PageSlice,
} from '../src/renderer/pagination'

const rectOf = (top: number, height: number) =>
  ({
    top,
    height,
    bottom: top + height,
    left: 0,
    right: 100,
    width: 100,
    x: 0,
    y: top,
    toJSON: () => ({}),
  }) as DOMRect

const cell = { type: 'docTableCell', content: [{ type: 'docParagraph' }] }
const row = { type: 'docTableRow', content: [cell] }

const editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })

function tableSpec(attrs: Record<string, unknown>): [string, Record<string, string>] {
  const table = editor.schema.nodeFromJSON({ type: 'docTable', attrs, content: [row] })
  return editor.schema.nodes.docTable.spec.toDOM!(table) as [string, Record<string, string>]
}

const suppressedFloat = {
  tblFloat: null,
  tblFloatSource: 'left',
  tblFloatSuppressed: true,
  tblFloatHorzAnchor: 'margin',
  tblFloatVertAnchor: 'text',
}

describe('suppressed text-anchored float with a negative w:tblpY', () => {
  it('keeps the lift as a negative top margin', () => {
    const [, attrs] = tableSpec({ ...suppressedFloat, tblFloatYTwips: -1096 })
    expect(attrs.style).toContain('margin-top:-73.1px')
    expect(attrs['data-tblp-lift']).toBe('73.1')
  })

  it('does not lift a positive offset, a page anchor or a live float', () => {
    expect(
      tableSpec({ ...suppressedFloat, tblFloatYTwips: 72 })[1]['data-tblp-lift'],
    ).toBeUndefined()
    expect(
      tableSpec({ ...suppressedFloat, tblFloatVertAnchor: 'page', tblFloatYTwips: -1096 })[1][
        'data-tblp-lift'
      ],
    ).toBeUndefined()
    const live = tableSpec({
      ...suppressedFloat,
      tblFloat: 'left',
      tblFloatSuppressed: false,
      tblFloatYTwips: -1096,
    })[1]
    expect(live['data-tblp-lift']).toBeUndefined()
    expect(live.style).toContain('margin-top:-73.1px')
  })
})

function flow() {
  const pm = document.createElement('div')
  const add = (tag: string, top: number, height: number, html: string) => {
    const el = document.createElement(tag)
    el.innerHTML = html
    el.getBoundingClientRect = () => rectOf(top, height)
    pm.appendChild(el)
    return el
  }
  add('p', 0, 20, 'mark')
  const table = add('table', -50, 140, '<tr><td>logos</td></tr>')
  table.dataset.tblpLift = '70'
  add('p', 90, 20, 'title')
  return pm
}

describe('measureBlocks — lifted table', () => {
  it('records the lift and the lifted top', () => {
    const { blocks } = measureBlocks(flow(), 0, 1)
    expect(blocks.map((b) => [b.top, b.height, b.liftPx ?? null])).toEqual([
      [0, 20, null],
      [-50, 140, 70],
      [90, 20, null],
    ])
  })
})

describe('applyLiftTops', () => {
  it('opens the clip above a page start by the hanging part only', () => {
    const blocks: BlockBox[] = [
      { top: 0, height: 20 },
      { top: -50, height: 140, liftPx: 70 },
      { top: 90, height: 20 },
      { top: 900, height: 140, liftPx: 70 },
    ]
    const slices: PageSlice[] = [
      { start: 0, end: 900, section: 0 },
      { start: 900, end: 1040, section: 0 },
    ]
    applyLiftTops(slices, blocks)
    expect(slices[0].liftTop).toBeCloseTo(50)
    expect(slices[1].liftTop).toBeUndefined()
  })

  it("opens the clip for a lifted table that is the page's first block", () => {
    const blocks: BlockBox[] = [{ top: -70, height: 140, liftPx: 70 }]
    const slices: PageSlice[] = [{ start: 0, end: 900, section: 0 }]
    applyLiftTops(slices, blocks)
    expect(slices[0].liftTop).toBeCloseTo(70)
  })
})

describe('slicer — lifted table charges only the flow it advances', () => {
  const cap = 900
  const geoms = [{ contentHeight: cap, forceBreak: false }]
  const rows = (h: number) => [{ top: 0, height: h }]

  it('credits a page-opening lift up to the top margin', () => {
    const first: BlockBox[] = [
      { top: -70, height: 140, liftPx: 70, tableRows: rows(140) },
      { top: 70, height: cap - 80 },
    ]
    const margin = (topPx: number) => [{ contentHeight: cap, forceBreak: false, topPx }]
    expect(computeSectionedSlicesF2(first, margin(96), cap - 10)).toHaveLength(1)
    expect(computeSectionedSlicesF2(first, margin(40), cap - 10)).toHaveLength(2)
  })

  it('charges the whole table when a later page opens at its lifted top', () => {
    const later: BlockBox[] = [
      { top: 0, height: cap - 10 },
      { top: cap - 80, height: 140, liftPx: 70, tableRows: rows(140) },
      { top: cap + 60, height: cap - 130 },
    ]
    const geomsMargin = [{ contentHeight: cap, forceBreak: false, topPx: 96 }]
    const slices = computeSectionedSlicesF2(later, geomsMargin, cap - 10)
    expect(slices).toHaveLength(3)
    expect(slices[1].start).toBe(cap - 80)
  })

  it('lets the block after a lifted table fit where stacked heights would not', () => {
    const lifted: BlockBox[] = [
      { top: 0, height: 20 },
      { top: -50, height: 140, liftPx: 70, tableRows: rows(140) },
      { top: 90, height: cap - 100 },
    ]
    expect(computeSectionedSlicesF2(lifted, geoms, cap - 10)).toHaveLength(1)
    const stacked: BlockBox[] = [
      { top: 0, height: 20 },
      { top: 20, height: 140, tableRows: rows(140) },
      { top: 160, height: cap - 100 },
    ]
    expect(computeSectionedSlicesF2(stacked, geoms, cap + 60)).toHaveLength(2)
  })
})
