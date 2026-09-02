/**
 * setTableCell (the op behind edit_table_cell, the cell editor and script
 * setTableCell) used to write the caller's paragraphs into the cell verbatim.
 * The AI tool sends plain text, so a rewritten cell fell back to the
 * table-style defaults — size, color, font and bold gone. The op now rebuilds
 * the paragraphs onto the cell's current ones like setText does for shapes.
 */
import { describe, expect, it, beforeEach } from 'vitest'
import {
  addTable,
  createBlankPptx,
  editTableCellText,
  openPptx,
  savePptx,
  type OpenedPptx,
  type TableElement,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/main/ops'

let opened: OpenedPptx
let tableId: string

// element ids are regenerated on reopen: the deck's only table is the one
const table = (o: OpenedPptx = opened) =>
  o.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
const cellRuns = (r: number, c: number, o: OpenedPptx = opened) =>
  table(o).rows[r]![c]!.text!.paragraphs.map((p) => p.runs)

beforeEach(async () => {
  opened = await openPptx(await createBlankPptx())
  const r = addTable(opened, 0, {
    rows: 2,
    cols: 2,
    offset: { x: 0, y: 0, cx: 3810000, cy: 952500 },
  })!
  tableId = r.elementId
  const slide = opened.deck.slides[0]!
  // a styled header cell and a two-run body cell, written with explicit run props
  editTableCellText(slide, tableId, 0, 0, [
    {
      runs: [{ text: 'Revenue', bold: true, fontSize: 14, color: '1F4E78', fontFamily: 'Calibri' }],
      align: 'center',
    },
  ])
  editTableCellText(slide, tableId, 1, 0, [
    {
      runs: [
        { text: 'Note: ', bold: true, fontSize: 10 },
        { text: 'unaudited figures for the quarter', fontSize: 10, color: '595959' },
      ],
    },
  ])
})

describe('setTableCell keeps the cell formatting the caller cannot express', () => {
  it('a plain-text rewrite (AI edit_table_cell) keeps size, color, font, bold and alignment', async () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setTableCell',
          target: { slide: 0, el: tableId },
          row: 0,
          col: 0,
          paragraphs: [{ runs: [{ text: '收入' }] }],
        },
      ],
    })
    expect(r.applied).toBe(true)
    const [runs] = cellRuns(0, 0)
    expect(runs).toHaveLength(1)
    expect(runs![0]).toMatchObject({
      text: '收入',
      bold: true,
      fontSize: 14,
      color: '1F4E78',
      fontFamily: 'Calibri',
    })
    expect(table().rows[0]![0]!.text!.paragraphs[0]!.align).toBe('center')
    const xml = table().anchor.originalXml
    expect(xml).toContain('收入')
    expect(xml).toMatch(/sz="1400"/)
    expect(xml).toMatch(/1F4E78/)
    expect(xml).toMatch(/typeface="Calibri"/)

    // the bytes survive save and reopen
    const reopened = await openPptx(await savePptx(opened))
    const [again] = cellRuns(0, 0, reopened)
    expect(again![0]).toMatchObject({ text: '收入', bold: true, fontSize: 14, color: '#1F4E78' })
  })

  it("the caller's explicit run props still win over the inherited ones", () => {
    runTxn(opened, {
      ops: [
        {
          op: 'setTableCell',
          target: { slide: 0, el: tableId },
          row: 0,
          col: 0,
          paragraphs: [{ runs: [{ text: 'Revenue', bold: false, fontSize: 18 }], align: 'left' }],
        },
      ],
    })
    const [runs] = cellRuns(0, 0)
    expect(runs![0]).toMatchObject({ text: 'Revenue', bold: false, fontSize: 18, color: '1F4E78' })
    expect(table().rows[0]![0]!.text!.paragraphs[0]!.align).toBe('left')
  })

  it('a single-run rewrite of a "Label: text" cell takes the text run formatting; extra paragraphs continue it', () => {
    runTxn(opened, {
      ops: [
        {
          op: 'setTableCell',
          target: { slide: 0, el: tableId },
          row: 1,
          col: 0,
          paragraphs: [{ runs: [{ text: 'Unaudited.' }] }, { runs: [{ text: 'Restated in Q3.' }] }],
        },
      ],
    })
    const paras = cellRuns(1, 0)
    expect(paras).toHaveLength(2)
    expect(paras[0]![0]).toMatchObject({ text: 'Unaudited.', fontSize: 10, color: '595959' })
    expect(paras[0]![0]!.bold).toBeUndefined()
    expect(paras[1]![0]).toMatchObject({ text: 'Restated in Q3.', fontSize: 10, color: '595959' })
  })

  it('paragraph format toggled in the cell editor (bullet, spacing, rtl) is applied on top of the rebuild', async () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setTableCell',
          target: { slide: 0, el: tableId },
          row: 1,
          col: 0,
          paragraphs: [
            {
              runs: [{ text: 'Unaudited figures', srcRun: 1 }],
              srcPara: 0,
              bullet: 'char',
              bulletChar: '–',
              lineSpacingPct: 150,
              spaceAfterPt: 6,
            },
            { runs: [{ text: 'מספרים' }], rtl: true },
          ],
        },
      ],
    })
    expect(r.applied).toBe(true)
    const paras = table().rows[1]![0]!.text!.paragraphs
    expect(paras).toHaveLength(2)
    // the run formatting still comes from the cell, the paragraph format from the patch
    expect(paras[0]!.runs[0]).toMatchObject({
      text: 'Unaudited figures',
      fontSize: 10,
      color: '595959',
    })
    expect(paras[0]!.bullet).toMatchObject({ type: 'char', char: '–' })
    expect(paras[0]!.lineHeight).toBe(150)
    expect(paras[0]!.spaceAfter).toBe(6)
    expect(paras[1]!.rtl).toBe(true)
    expect(paras[1]!.runs[0]).toMatchObject({ text: 'מספרים', fontSize: 10 })
    const xml = table().anchor.originalXml
    expect(xml).toContain('<a:buChar char="–"/>')
    expect(xml).toMatch(/<a:lnSpc><a:spcPct val="150000"\/><\/a:lnSpc>/)
    expect(xml).toMatch(/<a:spcAft><a:spcPts val="600"\/><\/a:spcAft>/)
    expect(xml).toMatch(/<a:pPr[^>]*\brtl="1"/)

    // and they survive save and reopen
    const reopened = await openPptx(await savePptx(opened))
    const again = table(reopened).rows[1]![0]!.text!.paragraphs
    expect(again[0]!.bullet).toMatchObject({ type: 'char', char: '–' })
    expect(again[0]!.lineHeight).toBe(150)
    expect(again[0]!.spaceAfter).toBe(6)
    expect(again[1]!.rtl).toBe(true)
  })

  it('an empty cell stays a plain write', () => {
    runTxn(opened, {
      ops: [
        {
          op: 'setTableCell',
          target: { slide: 0, el: tableId },
          row: 1,
          col: 1,
          paragraphs: [{ runs: [{ text: 'fresh' }] }],
        },
      ],
    })
    const [runs] = cellRuns(1, 1)
    expect(runs![0]!.text).toBe('fresh')
    expect(runs![0]!.fontSize).toBeUndefined()
  })
})
