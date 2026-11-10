import { describe, expect, it } from 'vitest'

import { applyStructuralOps, type StructuralOp } from '../src/gateway/xlsx-structure'

const SHEET = 'Data'

function worksheet(autoFilter: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:D10"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1"><v>4</v></c></row>
    <row r="10"><c r="A10"><v>5</v></c><c r="B10"><v>6</v></c><c r="C10"><v>7</v></c><c r="D10"><v>8</v></c></row>
  </sheetData>
${autoFilter}
</worksheet>`
}

const FILTERED =
  `<autoFilter ref="A1:D10">` +
  `<filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn>` +
  `<sortState ref="A2:D10"><sortCondition ref="C2:C10"/></sortState>` +
  `</autoFilter>`

function replay(ops: readonly StructuralOp[]): string {
  return applyStructuralOps(worksheet(FILTERED), ops, SHEET)
}

describe('auto-filter criteria replay on column edits', () => {
  it('re-bases filterColumn colId when a column is inserted inside the range', () => {
    const xml = replay([{ kind: 'insert-cols', index: 1, count: 1 }])
    expect(xml).toContain('<autoFilter ref="A1:E10">')
    expect(xml).toContain('<filterColumn colId="3">')
    expect(xml).toContain('<sortCondition ref="D2:D10"/>')
  })

  it('drops the criterion and sort target when their column is deleted', () => {
    const xml = replay([{ kind: 'remove-cols', index: 2, count: 1 }])
    expect(xml).toContain('<autoFilter ref="A1:C10">')
    expect(xml).not.toContain('<filterColumn')
    expect(xml).not.toContain('sortState')
    expect(xml).not.toContain('sortCondition')
  })

  it('keeps colId and moves only the row extents on a row insert', () => {
    const xml = replay([{ kind: 'insert-rows', index: 0, count: 2 }])
    expect(xml).toContain('<autoFilter ref="A3:D12">')
    expect(xml).toContain('<filterColumn colId="2">')
    expect(xml).toContain('<sortState ref="A4:D12">')
    expect(xml).toContain('<sortCondition ref="C4:C12"/>')
  })

  it('leaves a self-closing auto-filter with no criteria untouched apart from its range', () => {
    const xml = applyStructuralOps(
      worksheet('<autoFilter ref="A1:D10"/>'),
      [{ kind: 'insert-cols', index: 1, count: 1 }],
      SHEET,
    )
    expect(xml).toContain('<autoFilter ref="A1:E10"/>')
  })
})
