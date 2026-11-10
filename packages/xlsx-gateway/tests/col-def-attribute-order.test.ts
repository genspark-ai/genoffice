import { describe, expect, it } from 'vitest'

import { applyStructuralOps, type StructuralOp } from '../src/gateway/xlsx-structure'

const SHEET = 'Sheet1'

function sheet(cols: string): string {
  // CT_Worksheet order: <cols> precedes <sheetData>.
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>${cols}</cols>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row>
  </sheetData>
</worksheet>`
}

/// Insert one column before B (index 1), so C and everything after it moves right.
const INSERT_B: StructuralOp = { kind: 'insert-cols', index: 1, count: 1 }

function colsAfter(xml: string): string {
  return /<cols>([\s\S]*?)<\/cols>/.exec(xml)?.[1] ?? ''
}

describe('column definition shifting', () => {
  it('shifts a <col> written max before min', () => {
    // Attribute order is not significant, so this is a valid <col>. The old
    // pattern required min first and skipped it entirely.
    const before = sheet('<col max="3" min="3" width="20"/>')
    const after = applyStructuralOps(before, [INSERT_B], SHEET)
    expect(colsAfter(after)).toBe('<col max="4" min="4" width="20"/>')
  })

  it('shifts a <col> written in schema order too', () => {
    const before = sheet('<col min="3" max="3" width="20"/>')
    const after = applyStructuralOps(before, [INSERT_B], SHEET)
    expect(colsAfter(after)).toBe('<col min="4" max="4" width="20"/>')
  })

  it('shifts a range and keeps every other attribute in place', () => {
    const before = sheet('<col max="5" min="3" width="20" style="2" hidden="1"/>')
    const after = applyStructuralOps(before, [INSERT_B], SHEET)
    expect(colsAfter(after)).toBe('<col max="6" min="4" width="20" style="2" hidden="1"/>')
  })

  it('drops a <col> the delete fully covers', () => {
    // Column 2's <col> is deleted outright; column 3's takes its place at 2.
    const before = sheet('<col min="2" max="2" width="20"/><col min="3" max="3" width="8"/>')
    const after = applyStructuralOps(before, [{ kind: 'remove-cols', index: 1, count: 1 }], SHEET)
    expect(colsAfter(after)).toBe('<col min="2" max="2" width="8"/>')
  })
})
