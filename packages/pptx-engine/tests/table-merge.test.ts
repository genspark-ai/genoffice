/** Table cell merge/split: byte surgery + reparse + save/reopen round-trip. */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { addTable, editTableCellText, mergeTableCells, openPptx, savePptx } from '../src/index'
import type { TableElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

async function makeTable() {
  const opened = await openPptx(fx('01_standard_business.pptx'))
  const r = addTable(opened, 0, {
    rows: 3,
    cols: 3,
    offset: { x: 914400, y: 914400, cx: 4572000, cy: 1828800 },
  })!
  editTableCellText(r.slide, r.elementId, 0, 0, [{ runs: [{ text: 'A1' }] }])
  editTableCellText(r.slide, r.elementId, 0, 1, [{ runs: [{ text: 'B1' }] }])
  return { opened, slideIndex: 0, tblId: r.elementId }
}

const tableOf = (opened: Awaited<ReturnType<typeof makeTable>>['opened'], id: string) =>
  opened.deck.slides[0]!.elements.find((e) => e.id === id) as TableElement

describe('mergeTableCells', () => {
  it('merge right: gridSpan + hMerge + text merged into the anchor cell', async () => {
    const { opened, tblId } = await makeTable()
    const r = mergeTableCells(opened, 0, tblId, { kind: 'merge-right', row: 0, col: 0 })
    expect(r).not.toBeNull()
    const tbl = tableOf(opened, r!.elementId)
    expect(tbl.rows[0]![0]!.gridSpan).toBe(2)
    expect(tbl.rows[0]![1]!.merged).toBe(true)
    const anchorText = tbl.rows[0]![0]!.text!.paragraphs.map((p) =>
      p.runs.map((x) => x.text).join(''),
    ).join('|')
    expect(anchorText).toContain('A1')
    expect(anchorText).toContain('B1')
  })

  it('merge down + save/reopen preserves rowSpan/vMerge', async () => {
    const { opened, tblId } = await makeTable()
    const r = mergeTableCells(opened, 0, tblId, { kind: 'merge-down', row: 0, col: 0 })!
    const saved = await savePptx(opened)
    const reopened = await openPptx(saved)
    const tbl = reopened.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    expect(tbl.rows[0]![0]!.rowSpan).toBe(2)
    expect(tbl.rows[1]![0]!.merged).toBe(true)
    expect(r.elementId).toBeTruthy()
  })

  it('consecutive merge-right extends gridSpan', async () => {
    const { opened, tblId } = await makeTable()
    const r1 = mergeTableCells(opened, 0, tblId, { kind: 'merge-right', row: 0, col: 0 })!
    const r2 = mergeTableCells(opened, 0, r1.elementId, { kind: 'merge-right', row: 0, col: 0 })!
    const tbl = tableOf(opened, r2.elementId)
    expect(tbl.rows[0]![0]!.gridSpan).toBe(3)
  })

  it('split: clears spans and continuation markers', async () => {
    const { opened, tblId } = await makeTable()
    const r1 = mergeTableCells(opened, 0, tblId, { kind: 'merge-right', row: 1, col: 0 })!
    const r2 = mergeTableCells(opened, 0, r1.elementId, { kind: 'split', row: 1, col: 0 })!
    const tbl = tableOf(opened, r2.elementId)
    expect(tbl.rows[1]![0]!.gridSpan ?? 1).toBe(1)
    expect(tbl.rows[1]![1]!.merged ?? false).toBe(false)
  })

  it('invalid targets rejected: merging on a continuation cell / splitting a plain cell', async () => {
    const { opened, tblId } = await makeTable()
    const r1 = mergeTableCells(opened, 0, tblId, { kind: 'merge-right', row: 0, col: 0 })!
    expect(
      mergeTableCells(opened, 0, r1.elementId, { kind: 'merge-right', row: 0, col: 1 }),
    ).toBeNull()
    expect(mergeTableCells(opened, 0, r1.elementId, { kind: 'split', row: 2, col: 0 })).toBeNull()
  })
})
