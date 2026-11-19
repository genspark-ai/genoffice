import { describe, expect, it } from 'vitest'

import {
  shiftCrossSheetFormulas,
  StructuralShiftError,
  type StructuralOp,
} from '../src/gateway/xlsx-structure'

const EDITED = 'Data'

function otherWorksheet(target: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><f>${target}+Report!$A$1</f></c></row></sheetData>
  <conditionalFormatting sqref="B1:B9">
    <cfRule type="expression" priority="1"><formula>${target}&gt;3</formula></cfRule>
  </conditionalFormatting>
  <dataValidations count="1">
    <dataValidation type="list" sqref="C1:C9"><formula1>${target}</formula1></dataValidation>
  </dataValidations>
  <hyperlinks><hyperlink ref="A2" location="${target}" display="go"/></hyperlinks>
</worksheet>`
}

const ROW_INSERT: StructuralOp[] = [{ kind: 'insert-rows', index: 0, count: 1 }]
const COLUMN_INSERT: StructuralOp[] = [{ kind: 'insert-cols', index: 0, count: 1 }]

describe('cross-sheet replay of qualified targets', () => {
  it('moves conditional-format, validation, and hyperlink targets on a row insert', () => {
    const xml = shiftCrossSheetFormulas(otherWorksheet('Data!$A$5'), EDITED, ROW_INSERT)
    expect(xml).toContain('<f>Data!$A$6+Report!$A$1</f>')
    expect(xml).toContain('<formula>Data!$A$6&gt;3</formula>')
    expect(xml).toContain('<formula1>Data!$A$6</formula1>')
    expect(xml).toContain('location="Data!$A$6"')
  })

  it('expands a qualified range target on a row insert', () => {
    const xml = shiftCrossSheetFormulas(otherWorksheet('Data!$A$1:$A$9'), EDITED, ROW_INSERT)
    expect(xml).toContain('<formula1>Data!$A$2:$A$10</formula1>')
    expect(xml).toContain('location="Data!$A$2:$A$10"')
  })

  it('moves them sideways on a column insert', () => {
    const xml = shiftCrossSheetFormulas(otherWorksheet('Data!$B$9'), EDITED, COLUMN_INSERT)
    expect(xml).toContain('<f>Data!$C$9+Report!$A$1</f>')
    expect(xml).toContain('<formula>Data!$C$9&gt;3</formula>')
    expect(xml).toContain('location="Data!$C$9"')
  })

  it('leaves targets qualified with another sheet alone', () => {
    const xml = shiftCrossSheetFormulas(otherWorksheet('Report!$A$5'), EDITED, ROW_INSERT)
    expect(xml).toContain('<f>Report!$A$5+Report!$A$1</f>')
    expect(xml).toContain('<formula>Report!$A$5&gt;3</formula>')
    expect(xml).toContain('location="Report!$A$5"')
  })

  it('fails closed when a target is deleted outright', () => {
    expect(() =>
      shiftCrossSheetFormulas(otherWorksheet('Data!$A$1'), EDITED, [
        { kind: 'remove-cols', index: 0, count: 1 },
      ]),
    ).toThrow(StructuralShiftError)
  })
})

describe('hyperlink location attribute escaping', () => {
  it('keeps a quote in a quoted sheet name attribute-encoded', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <hyperlinks>
    <hyperlink ref="A2" location="Data!$A$5" display="go"/>
    <hyperlink ref="A3" location="'Re&quot;port'!$A$5" display="stay"/>
  </hyperlinks>
</worksheet>`
    const shifted = shiftCrossSheetFormulas(xml, EDITED, ROW_INSERT)
    expect(shifted).toContain('location="Data!$A$6"')
    expect(shifted).toContain(`location="'Re&quot;port'!$A$5"`)
    expect(shifted).not.toContain(`'Re"port'`)
  })
})
