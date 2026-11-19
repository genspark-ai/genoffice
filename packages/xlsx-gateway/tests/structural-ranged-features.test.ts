import { describe, expect, it } from 'vitest'

import { applyStructuralOps, type StructuralOp } from '../src/gateway/xlsx-structure'

const SHEET = 'Data'

function worksheet(trailing: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
    <row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>
  </sheetData>
${trailing}
</worksheet>`
}

const PAIRED_DV =
  '<dataValidations count="2">' +
  '<dataValidation type="list" sqref="A1" operator="equal"/>' +
  '<dataValidation type="list" sqref="B1:B2"><formula1>"x,y"</formula1></dataValidation>' +
  '</dataValidations>'

const PAIRED_CF =
  '<conditionalFormatting sqref="A1"/>' +
  '<conditionalFormatting sqref="B1:B2">' +
  '<cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule>' +
  '</conditionalFormatting>'

function replay(trailing: string, ops: readonly StructuralOp[]): string {
  return applyStructuralOps(worksheet(trailing), ops, SHEET)
}

const INSERT_COLUMN: StructuralOp[] = [{ kind: 'insert-cols', index: 0, count: 1 }]
const REMOVE_COLUMN: StructuralOp[] = [{ kind: 'remove-cols', index: 0, count: 1 }]

describe('ranged feature replay with adjacent self-closing and paired elements', () => {
  it('shifts both data validations when a column is inserted', () => {
    const xml = replay(PAIRED_DV, INSERT_COLUMN)
    expect(xml).toContain('<dataValidation type="list" sqref="B1" operator="equal"/>')
    expect(xml).toContain('<dataValidation type="list" sqref="C1:C2">')
    expect(xml).toContain('<dataValidations count="2">')
  })

  it('keeps the paired data validation when the leading one is deleted', () => {
    const xml = replay(PAIRED_DV, REMOVE_COLUMN)
    expect(xml).not.toContain('sqref="A1"')
    expect(xml).toContain('<dataValidation type="list" sqref="A1:A2">')
    expect(xml).toContain('<dataValidations count="1">')
  })

  it('shifts both conditional formats when a column is inserted', () => {
    const xml = replay(PAIRED_CF, INSERT_COLUMN)
    expect(xml).toContain('<conditionalFormatting sqref="B1"/>')
    expect(xml).toContain('<conditionalFormatting sqref="C1:C2">')
  })
})
