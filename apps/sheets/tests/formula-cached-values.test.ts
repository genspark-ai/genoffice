/**
 * Recalculated values lived only on screen, so a saved file had new
 * inputs and stale outputs — readers without a formula engine (openpyxl data_only,
 * pandas, preview services) silently got wrong numbers. The save now refreshes each
 * formula cell's cached <v> while leaving its <f> untouched.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { applyCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import { buildStructureFixture } from './fixture-builder'

const SHEET = 'Data'

/** The fixture's D2 is `=SUM(A1:A2)` with a cached value. */
async function saveWithFormulaValues(
  cells: { row: number; column: number; value: string | number | boolean | null }[],
) {
  const mutation = await applyCellEditsToXlsx(
    await buildStructureFixture(),
    // one unrelated value edit so the save has something to do
    [{ sheetName: SHEET, row: 0, column: 0, writeValue: true, cell: { value: 1 } }],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [{ sheetName: SHEET, cells }],
  )
  const zip = await JSZip.loadAsync(mutation.buffer)
  return (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
}

describe('formula cached values', () => {
  it('updates <v> and keeps <f> for a numeric result', async () => {
    const xml = await saveWithFormulaValues([{ row: 1, column: 3, value: 811.4 }])
    const cell = /<c[^>]*r="D2"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''
    expect(cell).toContain('<f>')
    expect(cell).toContain('<v>811.4</v>')
    // exactly one cached value, and no stale one left behind
    expect((cell.match(/<v>/g) ?? []).length).toBe(1)
  })

  it('a text result gets t="str"', async () => {
    const xml = await saveWithFormulaValues([{ row: 1, column: 3, value: 'N/A & more' }])
    const cell = /<c[^>]*r="D2"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''
    expect(cell).toContain('t="str"')
    expect(cell).toContain('<v>N/A &amp; more</v>')
    expect(cell).toContain('<f>')
  })

  it('a null result clears the cached value but keeps the formula', async () => {
    const xml = await saveWithFormulaValues([{ row: 1, column: 3, value: null }])
    const cell = /<c[^>]*r="D2"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''
    expect(cell).toContain('<f>')
    expect(cell).not.toContain('<v>')
  })

  it('non-formula cells and missing cells are left alone', async () => {
    const before = await saveWithFormulaValues([])
    // A1 is a plain number, ZZ99 does not exist
    const after = await saveWithFormulaValues([
      { row: 0, column: 0, value: 999 },
      { row: 98, column: 701, value: 999 },
    ])
    expect(after).toBe(before)
    expect(after).not.toContain('<v>999</v>')
  })
})
