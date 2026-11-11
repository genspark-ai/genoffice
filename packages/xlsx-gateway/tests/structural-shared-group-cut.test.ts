import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, assertOnlyTouchedEntriesChanged } from '../src/gateway/xlsx-gateway'
import {
  applyStructuralOps,
  shiftCrossSheetFormulas,
  type StructuralOp,
  StructuralShiftError,
} from '../src/gateway/xlsx-structure'

const SHEET = 'Data'

const worksheet = (rows: string) => `<worksheet><sheetData>${rows}</sheetData></worksheet>`

const shift = (xml: string, op: StructuralOp) => applyStructuralOps(xml, [op], SHEET)

const filledDown = worksheet(
  '<row r="3"><c r="D3"><f t="shared" ref="D3:D6" si="0">B1</f><v>1</v></c></row>' +
    '<row r="4"><c r="D4"><f t="shared" si="0"/><v>2</v></c></row>' +
    '<row r="5"><c r="D5"><f t="shared" si="0"/><v>3</v></c></row>' +
    '<row r="6"><c r="D6"><f t="shared" si="0"/><v>4</v></c></row>',
)

const filledRight = worksheet(
  '<row r="4">' +
    '<c r="C4"><f t="shared" ref="C4:F4" si="0">A1</f><v>1</v></c>' +
    '<c r="D4"><f t="shared" si="0"/><v>2</v></c>' +
    '<c r="E4"><f t="shared" si="0"/><v>3</v></c>' +
    '<c r="F4"><f t="shared" si="0"/><v>4</v></c>' +
    '</row>',
)

describe('a row or column shift that cuts through a shared formula group', () => {
  it('materializes the group when a removed row falls inside it', () => {
    expect(shift(filledDown, { kind: 'remove-rows', index: 4, count: 1 })).toBe(
      worksheet(
        '<row r="3"><c r="D3"><f>B1</f><v>1</v></c></row>' +
          '<row r="4"><c r="D4"><f>B2</f><v>2</v></c></row>' +
          '<row r="5"><c r="D5"><f>B4</f><v>4</v></c></row>',
      ),
    )
  })

  it('materializes the group when rows are inserted inside it', () => {
    expect(shift(filledDown, { kind: 'insert-rows', index: 4, count: 1 })).toBe(
      worksheet(
        '<row r="3"><c r="D3"><f>B1</f><v>1</v></c></row>' +
          '<row r="4"><c r="D4"><f>B2</f><v>2</v></c></row>' +
          '<row r="6"><c r="D6"><f>B3</f><v>3</v></c></row>' +
          '<row r="7"><c r="D7"><f>B4</f><v>4</v></c></row>',
      ),
    )
  })

  it('materializes the group when a removed column falls inside it', () => {
    expect(shift(filledRight, { kind: 'remove-cols', index: 4, count: 1 })).toBe(
      worksheet(
        '<row r="4">' +
          '<c r="C4"><f>A1</f><v>1</v></c>' +
          '<c r="D4"><f>B1</f><v>2</v></c>' +
          '<c r="E4"><f>D1</f><v>4</v></c>' +
          '</row>',
      ),
    )
  })

  it('materializes the group when columns are inserted inside it', () => {
    expect(shift(filledRight, { kind: 'insert-cols', index: 4, count: 1 })).toBe(
      worksheet(
        '<row r="4">' +
          '<c r="C4"><f>A1</f><v>1</v></c>' +
          '<c r="D4"><f>B1</f><v>2</v></c>' +
          '<c r="F4"><f>C1</f><v>3</v></c>' +
          '<c r="G4"><f>D1</f><v>4</v></c>' +
          '</row>',
      ),
    )
  })

  it('materializes a two-dimensional group cut by the shift', () => {
    const block = worksheet(
      '<row r="2">' +
        '<c r="B2"><f t="shared" ref="B2:C3" si="0">A1*10</f><v>1</v></c>' +
        '<c r="C2"><f t="shared" si="0"/><v>2</v></c>' +
        '</row>' +
        '<row r="3">' +
        '<c r="B3"><f t="shared" si="0"/><v>3</v></c>' +
        '<c r="C3"><f t="shared" si="0"/><v>4</v></c>' +
        '</row>',
    )
    expect(shift(block, { kind: 'insert-rows', index: 2, count: 1 })).toBe(
      worksheet(
        '<row r="2">' +
          '<c r="B2"><f>A1*10</f><v>1</v></c>' +
          '<c r="C2"><f>B1*10</f><v>2</v></c>' +
          '</row>' +
          '<row r="4">' +
          '<c r="B4"><f>A2*10</f><v>3</v></c>' +
          '<c r="C4"><f>B2*10</f><v>4</v></c>' +
          '</row>',
      ),
    )
  })

  it('translates an entity-encoded anchor body and keeps the anchor attributes', () => {
    const encoded = worksheet(
      '<row r="3"><c r="D3"><f t="shared" ref="D3:D5" si="0" ca="1">IF(B1&gt;0,&quot;a&lt;b&quot;,B1)</f><v>1</v></c></row>' +
        '<row r="4"><c r="D4"><f t="shared" si="0"/><v>2</v></c></row>' +
        '<row r="5"><c r="D5"><f t="shared" si="0"/><v>3</v></c></row>',
    )
    expect(shift(encoded, { kind: 'insert-rows', index: 4, count: 1 })).toBe(
      worksheet(
        '<row r="3"><c r="D3"><f ca="1">IF(B1&gt;0,"a&lt;b",B1)</f><v>1</v></c></row>' +
          '<row r="4"><c r="D4"><f>IF(B2&gt;0,"a&lt;b",B2)</f><v>2</v></c></row>' +
          '<row r="6"><c r="D6"><f>IF(B3&gt;0,"a&lt;b",B3)</f><v>3</v></c></row>',
      ),
    )
  })

  it('materializes a group past the boundary whose references straddle it', () => {
    const straddling = worksheet(
      '<row r="10"><c r="D10"><f t="shared" ref="D10:D11" si="0">B4</f><v>1</v></c></row>' +
        '<row r="11"><c r="D11"><f t="shared" si="0"/><v>2</v></c></row>',
    )
    expect(shift(straddling, { kind: 'insert-rows', index: 4, count: 1 })).toBe(
      worksheet(
        '<row r="11"><c r="D11"><f>B4</f><v>1</v></c></row>' +
          '<row r="12"><c r="D12"><f>B6</f><v>2</v></c></row>',
      ),
    )
  })

  it('leaves a group entirely above the boundary byte-identical', () => {
    expect(shift(filledDown, { kind: 'insert-rows', index: 9, count: 2 })).toBe(filledDown)
    expect(shift(filledDown, { kind: 'remove-rows', index: 9, count: 2 })).toBe(filledDown)
  })

  it('keeps a group entirely below the boundary shared and moves it whole', () => {
    expect(shift(filledDown, { kind: 'insert-rows', index: 0, count: 2 })).toBe(
      worksheet(
        '<row r="5"><c r="D5"><f t="shared" ref="D5:D8" si="0">B3</f><v>1</v></c></row>' +
          '<row r="6"><c r="D6"><f t="shared" si="0"/><v>2</v></c></row>' +
          '<row r="7"><c r="D7"><f t="shared" si="0"/><v>3</v></c></row>' +
          '<row r="8"><c r="D8"><f t="shared" si="0"/><v>4</v></c></row>',
      ),
    )
    const sameRow = worksheet(
      '<row r="3"><c r="D3"><f t="shared" ref="D3:D4" si="0">B3*2</f><v>1</v></c></row>' +
        '<row r="4"><c r="D4"><f t="shared" si="0"/><v>2</v></c></row>',
    )
    expect(shift(sameRow, { kind: 'remove-rows', index: 0, count: 1 })).toBe(
      worksheet(
        '<row r="2"><c r="D2"><f t="shared" ref="D2:D3" si="0">B2*2</f><v>1</v></c></row>' +
          '<row r="3"><c r="D3"><f t="shared" si="0"/><v>2</v></c></row>',
      ),
    )
    expect(shift(filledRight, { kind: 'insert-cols', index: 0, count: 1 })).toBe(
      worksheet(
        '<row r="4">' +
          '<c r="D4"><f t="shared" ref="D4:G4" si="0">B1</f><v>1</v></c>' +
          '<c r="E4"><f t="shared" si="0"/><v>2</v></c>' +
          '<c r="F4"><f t="shared" si="0"/><v>3</v></c>' +
          '<c r="G4"><f t="shared" si="0"/><v>4</v></c>' +
          '</row>',
      ),
    )
  })

  it('materializes a group on another sheet whose references the shift splits', () => {
    const other = worksheet(
      '<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">Data!B3</f><v>1</v></c></row>' +
        '<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row>' +
        '<row r="3"><c r="A3"><f t="shared" si="0"/><v>3</v></c></row>',
    )
    expect(
      shiftCrossSheetFormulas(other, SHEET, [{ kind: 'insert-rows', index: 3, count: 1 }]),
    ).toBe(
      worksheet(
        '<row r="1"><c r="A1"><f>Data!B3</f><v>1</v></c></row>' +
          '<row r="2"><c r="A2"><f>Data!B5</f><v>2</v></c></row>' +
          '<row r="3"><c r="A3"><f>Data!B6</f><v>3</v></c></row>',
      ),
    )
    expect(
      shiftCrossSheetFormulas(other, SHEET, [{ kind: 'insert-rows', index: 9, count: 1 }]),
    ).toBe(other)
  })

  it('refuses the shift when a follower cannot be translated inside the sheet', () => {
    const offSheet = worksheet(
      '<row r="3"><c r="D3"><f t="shared" ref="D3:D4" si="0">B1048576</f><v>1</v></c></row>' +
        '<row r="4"><c r="D4"><f t="shared" si="0"/><v>2</v></c></row>',
    )
    expect(() => shift(offSheet, { kind: 'insert-rows', index: 3, count: 1 })).toThrow(
      StructuralShiftError,
    )
    expect(() => shift(offSheet, { kind: 'remove-rows', index: 2, count: 1 })).toThrow(
      /shared formula in D4 would reference cells outside the sheet/,
    )
  })
})

describe('saving a workbook whose shift cuts a shared formula group', () => {
  it('rewrites only the cut group and keeps every other part and cell byte-identical', async () => {
    const dataSheet =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData>' +
      '<row r="1"><c r="B1"><v>10</v></c><c r="F1" s="1"><f t="shared" ref="F1:F2" si="1">B1*2</f><v>20</v></c></row>' +
      '<row r="2"><c r="B2"><v>20</v></c><c r="F2" s="1"><f t="shared" si="1"/><v>40</v></c></row>' +
      '<row r="3"><c r="B3"><v>30</v></c><c r="D3" s="1"><f t="shared" ref="D3:D6" si="0">B1</f><v>10</v></c></row>' +
      '<row r="4"><c r="B4"><v>40</v></c><c r="D4" s="1"><f t="shared" si="0"/><v>20</v></c></row>' +
      '<row r="5"><c r="A5" t="inlineStr"><is><t>gone</t></is></c><c r="D5" s="1"><f t="shared" si="0"/><v>30</v></c></row>' +
      '<row r="6"><c r="D6" s="1"><f t="shared" si="0"/><v>40</v></c></row>' +
      '<row r="8"><c r="E8"><v>1</v></c><c r="G8"><f t="shared" ref="G8:G9" si="2">E8+1</f><v>2</v></c></row>' +
      '<row r="9"><c r="E9"><v>2</v></c><c r="G9"><f t="shared" si="2"/><v>3</v></c></row>' +
      '</sheetData>' +
      '</worksheet>'
    const otherSheet =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData>' +
      '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">Data!B1</f><v>10</v></c></row>' +
      '<row r="2"><c r="A2"><f t="shared" si="0"/><v>20</v></c></row>' +
      '</sheetData>' +
      '</worksheet>'
    const source = await buildWorkbook(dataSheet, otherSheet)

    const mutation = await applyCellEditsToXlsx(
      source,
      [],
      [{ sheetName: SHEET, ops: [{ kind: 'remove-rows', index: 4, count: 1 }] }],
    )

    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    expect(mutation.touchedEntries).toEqual(['xl/workbook.xml', 'xl/worksheets/sheet1.xml'])
    const before = await JSZip.loadAsync(source)
    const after = await JSZip.loadAsync(mutation.buffer)
    const touched = new Set(mutation.touchedEntries)
    for (const name of Object.keys(before.files)) {
      if (touched.has(name)) continue
      expect(await after.file(name)?.async('nodebuffer'), name).toEqual(
        await before.file(name)?.async('nodebuffer'),
      )
    }
    expect(await after.file('xl/worksheets/sheet1.xml')?.async('text')).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<sheetData>' +
        '<row r="1"><c r="B1"><v>10</v></c><c r="F1" s="1"><f t="shared" ref="F1:F2" si="1">B1*2</f><v>20</v></c></row>' +
        '<row r="2"><c r="B2"><v>20</v></c><c r="F2" s="1"><f t="shared" si="1"/><v>40</v></c></row>' +
        '<row r="3"><c r="B3"><v>30</v></c><c r="D3" s="1"><f>B1</f><v>10</v></c></row>' +
        '<row r="4"><c r="B4"><v>40</v></c><c r="D4" s="1"><f>B2</f><v>20</v></c></row>' +
        '<row r="5"><c r="D5" s="1"><f>B4</f><v>40</v></c></row>' +
        '<row r="7"><c r="E7"><v>1</v></c><c r="G7"><f t="shared" ref="G7:G8" si="2">E7+1</f><v>2</v></c></row>' +
        '<row r="8"><c r="E8"><v>2</v></c><c r="G8"><f t="shared" si="2"/><v>3</v></c></row>' +
        '</sheetData>' +
        '</worksheet>',
    )
  })
})

async function buildWorkbook(dataSheet: string, otherSheet: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>' +
      '</workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts>' +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<borders count="1"><border/></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
      '</styleSheet>',
  )
  zip.file('xl/worksheets/sheet1.xml', dataSheet)
  zip.file('xl/worksheets/sheet2.xml', otherSheet)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}
