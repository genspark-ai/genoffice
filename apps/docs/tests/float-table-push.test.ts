import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeSectionedSlicesF2, type BlockBox } from '../src/renderer/pagination'

const cell = { type: 'docTableCell', content: [{ type: 'docParagraph' }] }
const row = { type: 'docTableRow', content: [cell] }
const editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })

function tableAttrs(attrs: Record<string, unknown>): Record<string, string> {
  const table = editor.schema.nodeFromJSON({ type: 'docTable', attrs, content: [row] })
  return (editor.schema.nodes.docTable.spec.toDOM!(table) as [string, Record<string, string>])[1]
}

const suppressedFloat = {
  tblFloat: null,
  tblFloatSource: 'left',
  tblFloatSuppressed: true,
  tblFloatHorzAnchor: 'text',
  tblFloatVertAnchor: 'text',
  tblFloatDistance: { top: 180, bottom: 180, left: 180, right: 180 },
}

describe('text-anchored floating table flowed inline turns the page whole when it does not fit', () => {
  it('marks text-anchored floats for the engine', () => {
    expect(tableAttrs({ ...suppressedFloat, tblFloatYTwips: 0 })['data-tblp-push']).toBe('1')
    expect(tableAttrs({ ...suppressedFloat, tblFloatYTwips: 312 })['data-tblp-push']).toBe('1')
    // a lifted table hangs above its anchor instead
    expect(
      tableAttrs({ ...suppressedFloat, tblFloatYTwips: -300 })['data-tblp-push'],
    ).toBeUndefined()
    expect(
      tableAttrs({ ...suppressedFloat, tblFloatVertAnchor: 'margin', tblFloatYTwips: 0 })[
        'data-tblp-push'
      ],
    ).toBeUndefined()
    // a live float turns the page the same way
    expect(
      tableAttrs({ ...suppressedFloat, tblFloat: 'left', tblFloatSuppressed: false })[
        'data-tblp-push'
      ],
    ).toBe('1')
    expect(tableAttrs({ tblFloat: null, tblFloatSource: null })['data-tblp-push']).toBeUndefined()
  })

  it('only a positive w:topFromText turns the page; without it Word splits in place', () => {
    expect(
      tableAttrs({ ...suppressedFloat, tblFloatDistance: { top: 20, bottom: 0 } })[
        'data-tblp-push'
      ],
    ).toBe('1')
    expect(
      tableAttrs({ ...suppressedFloat, tblFloatDistance: { top: 0, bottom: 180 } })[
        'data-tblp-push'
      ],
    ).toBeUndefined()
    expect(
      tableAttrs({ ...suppressedFloat, tblFloatDistance: { left: 180, right: 180 } })[
        'data-tblp-push'
      ],
    ).toBeUndefined()
    expect(
      tableAttrs({ ...suppressedFloat, tblFloatDistance: null })['data-tblp-push'],
    ).toBeUndefined()
  })

  it('starts the table on the next page instead of splitting it in the remainder', () => {
    const cap = 900
    const geoms = [{ contentHeight: cap, forceBreak: false }]
    const rows = Array.from({ length: 12 }, (_, i) => ({ top: i * 50, height: 50 }))
    const table = { top: 500, height: 600, tableRows: rows }
    const inline: BlockBox[] = [{ top: 0, height: 500 }, { ...table }]
    const split = computeSectionedSlicesF2(inline, geoms, cap - 10)
    expect(split).toHaveLength(2)
    expect(split[1].start).toBe(900)

    const pushed: BlockBox[] = [
      { top: 0, height: 500 },
      { ...table, floatPushWhole: true },
    ]
    const slices = computeSectionedSlicesF2(pushed, geoms, cap - 10)
    expect(slices).toHaveLength(2)
    expect(slices[0].end).toBe(500)
    expect(slices[1].start).toBe(500)

    // taller than a page: still opens on a fresh page, then splits at row boundaries
    const tallRows = Array.from({ length: 24 }, (_, i) => ({ top: i * 50, height: 50 }))
    const tall: BlockBox[] = [
      { top: 0, height: 500 },
      { top: 500, height: 1200, tableRows: tallRows, floatPushWhole: true },
    ]
    const tallSlices = computeSectionedSlicesF2(tall, geoms, cap - 10)
    expect(tallSlices.map((s) => s.start)).toEqual([0, 500, 1400])

    // a table that fits the remainder stays; one opening a page is never pushed
    const fitting: BlockBox[] = [
      { top: 0, height: 200 },
      { ...table, top: 200, floatPushWhole: true },
    ]
    expect(computeSectionedSlicesF2(fitting, geoms, cap - 10)).toHaveLength(1)
    const opening: BlockBox[] = [
      { top: 0, height: 1200, tableRows: tallRows, floatPushWhole: true },
    ]
    expect(computeSectionedSlicesF2(opening, geoms, cap - 10).map((s) => s.start)).toEqual([0, 900])
  })
})
