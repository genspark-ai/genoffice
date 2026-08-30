/** Table row height setting + cell/textbox vertical alignment. */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  addTable,
  openPptx,
  savePptx,
  resizeTable,
  setTableRowHeight,
  setTableCellAnchor,
  setTableColWidth,
  editTableStructure,
  setElementTextAnchor,
  addElement,
  createBlankPptx,
} from '../src/index'
import type { TableElement, TextElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('setTableRowHeight', () => {
  it('row height written to a:tr h, outer frame cy kept in sync', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 2,
      offset: { x: 0, y: 0, cx: 2000000, cy: 800000 },
    })!
    const tbl = r.slide.elements.find((e) => e.id === r.elementId) as TableElement
    const h0 = tbl.rowHeights[0]!
    expect(setTableRowHeight(r.slide, r.elementId, 0, h0 * 2)).toBe(true)
    expect(tbl.rowHeights[0]).toBe(h0 * 2)
    expect(tbl.anchor.originalXml).toContain(`<a:tr h="${h0 * 2}"`)
    const sum = tbl.rowHeights.reduce((a, b) => a + b, 0)
    expect(tbl.transform.offset.cy).toBe(sum)
    expect(setTableRowHeight(r.slide, r.elementId, 9, 100)).toBe(false)
  })
})

describe('resizeTable', () => {
  it('redistributes gridCol widths and tr heights so XML sums match the frame', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 3,
      cols: 4,
      offset: { x: 0, y: 0, cx: 2000000, cy: 900000 },
    })!
    const tbl = r.slide.elements.find((e) => e.id === r.elementId) as TableElement

    expect(resizeTable(r.slide, r.elementId, 1000000, 600000)).toBe(true)
    expect(tbl.colWidths.reduce((a, b) => a + b, 0)).toBe(1000000)
    expect(tbl.rowHeights.reduce((a, b) => a + b, 0)).toBe(600000)
    expect(tbl.transform.offset.cx).toBe(1000000)
    expect(tbl.transform.offset.cy).toBe(600000)

    // XML agrees with the model: gridCol/tr h rewritten, ext synced
    const xml = tbl.anchor.originalXml
    const colSum = [...xml.matchAll(/<a:gridCol w="(\d+)"/g)].reduce((a, m) => a + Number(m[1]), 0)
    const rowSum = [...xml.matchAll(/<a:tr h="(\d+)"/g)].reduce((a, m) => a + Number(m[1]), 0)
    expect(colSum).toBe(1000000)
    expect(rowSum).toBe(600000)
    expect(xml).toMatch(/<a:ext cx="1000000" cy="600000"\/>/)
  })

  it('proportions survive save → reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 2,
      offset: { x: 0, y: 0, cx: 2000000, cy: 800000 },
    })!
    // make col0 twice col1, then shrink the table by half
    expect(resizeTable(r.slide, r.elementId, 2000000, 800000)).toBe(true)
    const tbl = r.slide.elements.find((e) => e.id === r.elementId) as TableElement
    tbl.colWidths = [1333334, 666666]
    expect(resizeTable(r.slide, r.elementId, 1000000, 400000)).toBe(true)

    const reopened = await openPptx(await savePptx(opened))
    const tbl2 = reopened.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    expect(tbl2.colWidths.reduce((a, b) => a + b, 0)).toBe(1000000)
    expect(tbl2.colWidths[0]).toBeGreaterThan(tbl2.colWidths[1]! * 1.9)
    expect(tbl2.rowHeights.reduce((a, b) => a + b, 0)).toBe(400000)
  })
})

describe('setTableCellAnchor', () => {
  it('tcPr anchor written / replaced', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 2,
      offset: { x: 0, y: 0, cx: 2000000, cy: 800000 },
    })!
    expect(setTableCellAnchor(r.slide, r.elementId, 0, 0, 'middle')).toBe(true)
    const tbl = r.slide.elements.find((e) => e.id === r.elementId) as TableElement
    expect(tbl.anchor.originalXml).toMatch(/<a:tcPr anchor="ctr"/)
    expect(setTableCellAnchor(r.slide, r.elementId, 0, 0, 'bottom')).toBe(true)
    expect(tbl.anchor.originalXml).toMatch(/<a:tcPr anchor="b"/)
    expect(tbl.anchor.originalXml).not.toMatch(/anchor="ctr"[^>]*anchor=/)
  })
})

describe('setElementTextAnchor', () => {
  it('bodyPr anchor written, model kept in sync', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 1000, cy: 1000 },
      paragraphs: [{ runs: [{ text: 'x' }] }],
    })
    expect(setElementTextAnchor(slide, el.id, 'middle')).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr anchor="ctr"/)
    expect((el as TextElement).text!.anchor).toBe('middle')
    expect(setElementTextAnchor(slide, el.id, 'top')).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr anchor="t"/)
  })
})

// Google Slides exports attach an <a:extLst> a16:colId to every gridCol, so
// the columns are paired tags, not self-closing — the surgery paths must not
// assume the self-closing form (they used to, and missed or mis-indexed cols).
function pairGridCols(tbl: TableElement): void {
  let i = 0
  tbl.anchor.originalXml = tbl.anchor.originalXml.replace(
    /<a:gridCol w="(\d+)"\/>/g,
    (_m, w: string) =>
      `<a:gridCol w="${w}"><a:extLst><a:ext uri="{9D8B030D-6E8A-4147-A177-3AD203B41FA5}">` +
      `<a16:colId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="${20000 + i++}"/>` +
      `</a:ext></a:extLst></a:gridCol>`,
  )
}

describe('paired-form gridCol (extLst children)', () => {
  it('setTableColWidth patches the right column and keeps the extLst', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 3,
      offset: { x: 0, y: 0, cx: 3000000, cy: 800000 },
    })!
    const tbl = r.slide.elements.find((e) => e.id === r.elementId) as TableElement
    pairGridCols(tbl)
    expect(setTableColWidth(r.slide, r.elementId, 1, 999000)).toBe(true)
    const xml = tbl.anchor.originalXml
    const cols = [...xml.matchAll(/<a:gridCol w="(\d+)">/g)].map((m) => Number(m[1]))
    expect(cols[1]).toBe(999000)
    expect(tbl.colWidths[1]).toBe(999000)
    expect([...xml.matchAll(/a16:colId/g)]).toHaveLength(3)
    const sum = tbl.colWidths.reduce((a, b) => a + b, 0)
    expect(xml).toContain(`cx="${sum}"`)
  })

  it('resizeTable scales paired-form columns too', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 2,
      offset: { x: 0, y: 0, cx: 2000000, cy: 800000 },
    })!
    const tbl = r.slide.elements.find((e) => e.id === r.elementId) as TableElement
    pairGridCols(tbl)
    expect(resizeTable(r.slide, r.elementId, 1000000, 400000)).toBe(true)
    const xml = tbl.anchor.originalXml
    const colSum = [...xml.matchAll(/<a:gridCol w="(\d+)">/g)].reduce((a, m) => a + Number(m[1]), 0)
    expect(colSum).toBe(1000000)
  })

  it('editTableStructure counts paired-form columns and inserts a fresh colId-free clone', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 2,
      offset: { x: 0, y: 0, cx: 2000000, cy: 800000 },
    })!
    const tbl = r.slide.elements.find((e) => e.id === r.elementId) as TableElement
    pairGridCols(tbl)
    const res = editTableStructure(opened, 0, r.elementId, { kind: 'insert-col', index: 1 })
    expect(res).not.toBeNull()
    const fresh = res!.slide.elements.find((e) => e.id === res!.elementId) as TableElement
    expect(fresh.colWidths).toHaveLength(3)
    // the clone must not duplicate the reference column's unique a16:colId
    expect([...fresh.anchor.originalXml.matchAll(/a16:colId/g)]).toHaveLength(2)
    const del = editTableStructure(opened, 0, res!.elementId, { kind: 'delete-col', index: 0 })
    expect(del).not.toBeNull()
    const after = del!.slide.elements.find((e) => e.id === del!.elementId) as TableElement
    expect(after.colWidths).toHaveLength(2)
  })
})
