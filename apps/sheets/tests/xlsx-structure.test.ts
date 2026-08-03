import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, assertOnlyTouchedEntriesChanged } from '../src/gateway/xlsx-gateway'
import {
  applyStructuralOps,
  shiftCrossSheetFormulas,
  shiftDefinedNames,
  shiftDrawingAnchors,
  shiftFormulaText,
  shiftTablePart,
  StructuralShiftError,
} from '../src/gateway/xlsx-structure'
import { buildStructureFixture } from './fixture-builder'

const SHEET = 'Data'

async function fixtureWorksheet(): Promise<string> {
  const zip = await JSZip.loadAsync(await buildStructureFixture())
  return zip.file('xl/worksheets/sheet1.xml')?.async('text') ?? ''
}

describe('applyStructuralOps rows', () => {
  it('inserting rows renumbers rows, cells, merges, dimension, and scoped ranges', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'insert-rows', index: 1, count: 2 }],
      SHEET,
    )
    expect(xml).toContain('<row r="1" spans="1:4">')
    expect(xml).toContain('<row r="4"><c r="A4"><v>2</v></c><c r="D4">')
    expect(xml).toContain('<row r="12"><c r="A12"><v>10</v></c></row>')
    expect(xml).toContain('<dimension ref="A1:D12"/>')
    expect(xml).toContain('<mergeCell ref="A7:B8"/>')
    expect(xml).toContain('<mergeCell ref="C1:D1"/>')
    expect(xml).toContain('sqref="A1:A12"')
    expect(xml).toContain('sqref="B4:B5"')
  })

  it('shifts formulas: relative, absolute, ranges, and leaves other sheets and strings alone', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'insert-rows', index: 1, count: 2 }],
      SHEET,
    )
    // Range A1:A2 spans the insert point, so its end expands.
    expect(xml).toContain('<f>SUM(A1:A4)</f>')
    // $A$2 shifts (Excel moves absolute references too); Other!B9 and the
    // "A1" string literal stay; whole-column B:B is row-op invariant.
    expect(xml).toContain('<f>$A$4+Other!B9+SUM(B:B)&amp;"A1"</f>')
    // Conditional-formatting rule formulas shift with their ranges.
    expect(xml).toContain('<formula>$A4&gt;5</formula>')
  })

  it('deleting rows removes their cells, shrinks merges, and clamps ranges', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'remove-rows', index: 4, count: 2 }],
      SHEET,
    )
    // Rows 5-6 (0-based 4-5) removed: the A5:B6 merge dies entirely.
    expect(xml).not.toContain('A5:B6')
    expect(xml).toContain('<mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells>')
    expect(xml).toContain('<row r="8"><c r="A8"><v>10</v></c></row>')
    expect(xml).toContain('sqref="A1:A8"')
  })

  it('aborts when a formula references only deleted cells', async () => {
    expect(() =>
      shiftFormulaText(
        'A5*2',
        SHEET,
        { boundary: 4, delta: -2, deleted: { start: 4, end: 5 } },
        'row',
      ),
    ).toThrow(StructuralShiftError)
  })

  it('clamps a range formula that straddles the deleted rows', () => {
    const shifted = shiftFormulaText(
      'SUM(A1:A10)',
      SHEET,
      { boundary: 4, delta: -2, deleted: { start: 4, end: 5 } },
      'row',
    )
    expect(shifted).toBe('SUM(A1:A8)')
  })
})

describe('applyStructuralOps columns', () => {
  it('inserting a column shifts cell letters, col defs, and drops span hints', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'insert-cols', index: 1, count: 1 }],
      SHEET,
    )
    expect(xml).not.toContain('spans=')
    expect(xml).toContain('<c r="D1"><v>7</v></c>')
    expect(xml).toContain('<c r="E2"><f>SUM(A1:A2)</f>')
    expect(xml).toContain('<col min="3" max="4" width="20" customWidth="1"/>')
    // B4's own cell moves to C4 and its whole-column reference shifts.
    expect(xml).toContain('<c r="C4"><f>$A$2+Other!B9+SUM(C:C)&amp;"A1"</f>')
    expect(xml).toContain('<mergeCell ref="A5:C6"/>')
    expect(xml).toContain('sqref="C2:C3"')
  })

  it('deleting a column removes its cells and shifts the rest left', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'remove-cols', index: 2, count: 1 }],
      SHEET,
    )
    expect(xml).not.toContain('<c r="C1">')
    expect(xml).toContain('<c r="C2"><f>SUM(A1:A2)</f>')
    expect(xml).toContain('<mergeCell ref="A5:B6"/>')
    // C1:D1 merge loses its deleted column and collapses to a single cell.
    expect(xml).toContain('<mergeCell ref="C1:C1"/>')
  })

  it('ops compose in order', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [
        { kind: 'insert-rows', index: 0, count: 1 },
        { kind: 'remove-rows', index: 0, count: 1 },
      ],
      SHEET,
    )
    expect(xml).toContain('<row r="1" spans="1:4"><c r="A1"><v>1</v></c>')
    expect(xml).toContain('<f>SUM(A1:A2)</f>')
  })
})

describe('merge operations', () => {
  it('adds a merge to an existing mergeCells section and updates the count', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'merge-cells', range: { startRow: 7, endRow: 8, startColumn: 0, endColumn: 1 } }],
      SHEET,
    )
    expect(xml).toContain('<mergeCell ref="A8:B9"/></mergeCells>')
    expect(xml).toContain('<mergeCells count="3">')
  })

  it('creates the mergeCells section after sheetData when absent', () => {
    const xml = applyStructuralOps(
      '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
      [{ kind: 'merge-cells', range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 } }],
      SHEET,
    )
    expect(xml).toContain('</sheetData><mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells>')
  })

  it('unmerges by removing the exact range and drops an emptied section', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [
        { kind: 'unmerge-cells', range: { startRow: 4, endRow: 5, startColumn: 0, endColumn: 1 } },
        { kind: 'unmerge-cells', range: { startRow: 0, endRow: 0, startColumn: 2, endColumn: 3 } },
      ],
      SHEET,
    )
    expect(xml).not.toContain('<mergeCells')
  })

  it('a merge recorded before a row insert lands at the shifted position', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [
        { kind: 'merge-cells', range: { startRow: 7, endRow: 7, startColumn: 0, endColumn: 3 } },
        { kind: 'insert-rows', index: 0, count: 1 },
      ],
      SHEET,
    )
    expect(xml).toContain('<mergeCell ref="A9:D9"/>')
  })
})

describe('cross-sheet reference rewriting', () => {
  const otherSheet =
    '<worksheet><sheetData><row r="9"><c r="B9"><f>SUM(Data!A5:A10)+Data!$A$8+A5</f><v>1</v></c></row></sheetData></worksheet>'

  it('shifts only references qualified to the edited sheet', () => {
    const shifted = shiftCrossSheetFormulas(otherSheet, SHEET, [
      { kind: 'insert-rows', index: 2, count: 2 },
    ])
    // Data!A5:A10 and Data!$A$8 shift; the local A5 does not.
    expect(shifted).toContain('<f>SUM(Data!A7:A12)+Data!$A$10+A5</f>')
  })

  it('skips self-closing shared formulas and still shifts the formula after them', () => {
    // `<f t="shared" si="0"/>` is not an opening tag: matching it once
    // swallowed everything up to the next `</f>` and re-escaped it (suppliers
    // regression: `<v>` became `&lt;v&gt;`).
    const sharedFormulaSheet =
      '<worksheet><sheetData><row r="5">' +
      '<c r="F5" s="26"><f t="shared" si="0"/><v>0.24</v></c>' +
      '<c r="G5"><f>SUM(Data!A5:A9)</f><v>7</v></c>' +
      '</row></sheetData></worksheet>'
    const shifted = shiftCrossSheetFormulas(sharedFormulaSheet, SHEET, [
      { kind: 'insert-rows', index: 2, count: 1 },
    ])
    expect(shifted).toContain('<f t="shared" si="0"/><v>0.24</v>')
    expect(shifted).toContain('<f>SUM(Data!A6:A10)</f>')
    expect(shifted).not.toContain('&lt;v&gt;')
  })

  it('fails closed when a qualified reference lands in a deleted range', () => {
    expect(() =>
      shiftCrossSheetFormulas(otherSheet, SHEET, [{ kind: 'remove-rows', index: 7, count: 1 }]),
    ).toThrow(StructuralShiftError)
  })

  it('shifts defined names', () => {
    const workbook =
      '<workbook><definedNames><definedName name="x">Data!$A$8</definedName><definedName name="y">Other!$A$8</definedName></definedNames></workbook>'
    const shifted = shiftDefinedNames(workbook, SHEET, [
      { kind: 'insert-rows', index: 2, count: 3 },
    ])
    expect(shifted).toContain('<definedName name="x">Data!$A$11</definedName>')
    expect(shifted).toContain('<definedName name="y">Other!$A$8</definedName>')
  })
})

const DRAWING_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<xdr:twoCellAnchor>' +
  '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>100</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>50</xdr:rowOff></xdr:from>' +
  '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>200</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>60</xdr:rowOff></xdr:to>' +
  '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Shape 2"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp>' +
  '<xdr:clientData/>' +
  '</xdr:twoCellAnchor>' +
  '<xdr:oneCellAnchor>' +
  '<xdr:from><xdr:col>2</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>20</xdr:rowOff></xdr:from>' +
  '<xdr:ext cx="100" cy="100"/>' +
  '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Pic 3"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill/><xdr:spPr/></xdr:pic>' +
  '<xdr:clientData/>' +
  '</xdr:oneCellAnchor>' +
  '<xdr:absoluteAnchor><xdr:pos x="10" y="10"/><xdr:ext cx="5" cy="5"/><xdr:clientData/></xdr:absoluteAnchor>' +
  '</xdr:wsDr>'

describe('shiftDrawingAnchors', () => {
  it('row insert shifts markers past the boundary; absoluteAnchor stays', () => {
    const xml = shiftDrawingAnchors(DRAWING_XML, [{ kind: 'insert-rows', index: 6, count: 2 }])
    // from row 5 is before the boundary; to row 8 shifts (object grows).
    expect(xml).toContain(
      '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>100</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>50</xdr:rowOff></xdr:from>',
    )
    expect(xml).toContain(
      '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>200</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>60</xdr:rowOff></xdr:to>',
    )
    expect(xml).toContain('<xdr:row>8</xdr:row><xdr:rowOff>20</xdr:rowOff>')
    expect(xml).toContain('<xdr:absoluteAnchor><xdr:pos x="10" y="10"/>')
  })

  it('row delete clamps markers inside the span and zeroes their offsets', () => {
    const xml = shiftDrawingAnchors(DRAWING_XML, [{ kind: 'remove-rows', index: 5, count: 4 }])
    // The whole object (rows 5-8) sat in the deleted span: kept, compressed.
    expect(xml).toContain(
      '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>100</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
    )
    expect(xml).toContain(
      '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>200</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>',
    )
    expect(xml).toContain('name="Shape 2"')
    // oneCellAnchor's from clamps too; its ext is untouched.
    expect(xml).toContain(
      '<xdr:col>2</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff>',
    )
    expect(xml).toContain('<xdr:ext cx="100" cy="100"/>')
  })

  it('rows past a deleted span shift up without clamping', () => {
    const xml = shiftDrawingAnchors(DRAWING_XML, [{ kind: 'remove-rows', index: 0, count: 2 }])
    expect(xml).toContain('<xdr:row>3</xdr:row><xdr:rowOff>50</xdr:rowOff>')
    expect(xml).toContain('<xdr:row>6</xdr:row><xdr:rowOff>60</xdr:rowOff>')
  })

  it('column insert and delete move the col markers', () => {
    const inserted = shiftDrawingAnchors(DRAWING_XML, [{ kind: 'insert-cols', index: 2, count: 1 }])
    expect(inserted).toContain('<xdr:from><xdr:col>1</xdr:col>')
    expect(inserted).toContain('<xdr:to><xdr:col>4</xdr:col>')
    const removed = shiftDrawingAnchors(DRAWING_XML, [{ kind: 'remove-cols', index: 2, count: 2 }])
    // to col 3 was inside the deleted span: clamps to 2, colOff zeroed.
    expect(removed).toContain('<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff>')
  })
})

const TABLE_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="B2:D5" totalsRowShown="0">' +
  '<autoFilter ref="B2:D5"/>' +
  '<tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns>' +
  '<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>' +
  '</table>'

describe('shiftTablePart', () => {
  it('inserting rows inside the table grows ref and autoFilter', () => {
    const xml = shiftTablePart(TABLE_XML, [{ kind: 'insert-rows', index: 2, count: 2 }])
    expect(xml).toContain('ref="B2:D7" totalsRowShown="0"')
    expect(xml).toContain('<autoFilter ref="B2:D7"/>')
  })

  it('inserting rows above shifts the table; below leaves it alone', () => {
    const above = shiftTablePart(TABLE_XML, [{ kind: 'insert-rows', index: 0, count: 1 }])
    expect(above).toContain('ref="B3:D6"')
    expect(above).toContain('<autoFilter ref="B3:D6"/>')
    expect(shiftTablePart(TABLE_XML, [{ kind: 'insert-rows', index: 5, count: 3 }])).toBe(TABLE_XML)
  })

  it('deleting data rows shrinks the table', () => {
    const xml = shiftTablePart(TABLE_XML, [{ kind: 'remove-rows', index: 3, count: 1 }])
    expect(xml).toContain('ref="B2:D4"')
  })

  it('fails closed on header-row deletion, whole-table deletion, and emptying deletions', () => {
    expect(() => shiftTablePart(TABLE_XML, [{ kind: 'remove-rows', index: 1, count: 1 }])).toThrow(
      /header row/,
    )
    expect(() => shiftTablePart(TABLE_XML, [{ kind: 'remove-rows', index: 0, count: 10 }])).toThrow(
      StructuralShiftError,
    )
    expect(() => shiftTablePart(TABLE_XML, [{ kind: 'remove-rows', index: 2, count: 3 }])).toThrow(
      /without data rows/,
    )
  })

  it('fails closed when a totals row would be deleted', () => {
    const withTotals = TABLE_XML.replace('totalsRowShown="0"', 'totalsRowCount="1"')
    expect(() => shiftTablePart(withTotals, [{ kind: 'remove-rows', index: 4, count: 1 }])).toThrow(
      /totals row/,
    )
  })

  it('column shifts outside the table move ref; inside fail closed', () => {
    const before = shiftTablePart(TABLE_XML, [{ kind: 'insert-cols', index: 0, count: 2 }])
    expect(before).toContain('ref="D2:F5"')
    expect(before).toContain('<autoFilter ref="D2:F5"/>')
    const removedLeft = shiftTablePart(TABLE_XML, [{ kind: 'remove-cols', index: 0, count: 1 }])
    expect(removedLeft).toContain('ref="A2:C5"')
    expect(shiftTablePart(TABLE_XML, [{ kind: 'insert-cols', index: 4, count: 1 }])).toBe(TABLE_XML)
    expect(() => shiftTablePart(TABLE_XML, [{ kind: 'insert-cols', index: 2, count: 1 }])).toThrow(
      /Inserting columns inside table/,
    )
    expect(() => shiftTablePart(TABLE_XML, [{ kind: 'remove-cols', index: 3, count: 1 }])).toThrow(
      /Deleting columns of table/,
    )
  })

  it('fails closed on merges over the table, allows merges elsewhere', () => {
    expect(() =>
      shiftTablePart(TABLE_XML, [
        { kind: 'merge-cells', range: { startRow: 2, endRow: 2, startColumn: 1, endColumn: 3 } },
      ]),
    ).toThrow(/Merging cells over table/)
    expect(
      shiftTablePart(TABLE_XML, [
        { kind: 'merge-cells', range: { startRow: 8, endRow: 9, startColumn: 0, endColumn: 1 } },
      ]),
    ).toBe(TABLE_XML)
  })

  it('ops compose: an insert above then rows inside use replayed coordinates', () => {
    const xml = shiftTablePart(TABLE_XML, [
      { kind: 'insert-rows', index: 0, count: 1 },
      { kind: 'insert-rows', index: 3, count: 1 },
    ])
    expect(xml).toContain('ref="B3:D7"')
  })
})

describe('calcChain invalidation', () => {
  it('a plain cell edit — formula cell turned into a literal — also drops calcChain', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildStructureFixture(),
      // D2 is the fixture's formula cell; overwriting it with a number leaves
      // calcChain pointing at a cell that no longer has an <f>
      [{ sheetName: SHEET, row: 1, column: 3, writeValue: true, cell: { value: 4242.5 } }],
      [],
    )
    expect(mutation.removedEntries).toEqual(['xl/calcChain.xml'])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const zip = await JSZip.loadAsync(mutation.buffer)
    expect(zip.file('xl/calcChain.xml')).toBeNull()
    expect(await zip.file('[Content_Types].xml')?.async('text')).not.toContain('calcChain')
    expect(await zip.file('xl/_rels/workbook.xml.rels')?.async('text')).not.toContain('calcChain')
  })

  it('the drop is tied to the worksheet rewrite, not to the kind of edit', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildStructureFixture(),
      // a value edit on a plain (non-formula) cell: no shift, no sheet-set change
      [{ sheetName: SHEET, row: 0, column: 0, writeValue: true, cell: { value: 'plain' } }],
      [],
    )
    expect(mutation.touchedEntries).toContain('xl/worksheets/sheet1.xml')
    expect(mutation.removedEntries).toEqual(['xl/calcChain.xml'])
  })
})

describe('structural save integration', () => {
  it('drops calcChain with its content-type and relationship, then passes preservation', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildStructureFixture(),
      [
        {
          sheetName: SHEET,
          row: 1,
          column: 0,
          writeValue: true,
          cell: { value: 'inserted row value' },
        },
      ],
      [{ sheetName: SHEET, ops: [{ kind: 'insert-rows', index: 1, count: 1 }] }],
    )
    expect(mutation.removedEntries).toEqual(['xl/calcChain.xml'])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const zip = await JSZip.loadAsync(mutation.buffer)
    expect(zip.file('xl/calcChain.xml')).toBeNull()
    const contentTypes = await zip.file('[Content_Types].xml')?.async('text')
    expect(contentTypes).not.toContain('calcChain')
    const rels = await zip.file('xl/_rels/workbook.xml.rels')?.async('text')
    expect(rels).not.toContain('calcChain')
    const worksheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    // The cell edit was journaled at post-insert coordinates: it lands in the
    // newly inserted empty row 2, between the untouched row 1 and shifted row 3.
    expect(worksheet).toContain(
      '<row r="2"><c r="A2" t="inlineStr"><is><t xml:space="preserve">inserted row value</t></is></c></row>',
    )
    expect(worksheet).toContain('<row r="3"><c r="A3"><v>2</v></c><c r="D3"><f>SUM(A1:A3)</f>')
    expect(worksheet).toContain('<row r="11"><c r="A11"><v>10</v></c></row>')
    const workbook = await zip.file('xl/workbook.xml')?.async('text')
    expect(workbook).toContain('fullCalcOnLoad="1"')
  })

  it('drops calcChain on a plain cell edit too — a formula overwritten by a literal invalidates it', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildStructureFixture(),
      [
        {
          sheetName: SHEET,
          row: 1,
          column: 3, // D2 carries a formula in the fixture's calcChain
          writeValue: true,
          cell: { value: 4242.5 },
        },
      ],
      [],
    )
    expect(mutation.removedEntries).toEqual(['xl/calcChain.xml'])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const zip = await JSZip.loadAsync(mutation.buffer)
    expect(zip.file('xl/calcChain.xml')).toBeNull()
    expect(await zip.file('[Content_Types].xml')?.async('text')).not.toContain('calcChain')
    expect(await zip.file('xl/_rels/workbook.xml.rels')?.async('text')).not.toContain('calcChain')
  })

  it('rewrites another sheet referencing the edited one and touches it', async () => {
    const zip = await JSZip.loadAsync(await buildStructureFixture())
    zip.file(
      'xl/worksheets/sheet2.xml',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>Data!A10</f><v>1</v></c></row></sheetData></worksheet>',
      { createFolders: false },
    )
    const buffer = await zip.generateAsync({ type: 'nodebuffer' })
    const mutation = await applyCellEditsToXlsx(
      buffer,
      [],
      [{ sheetName: SHEET, ops: [{ kind: 'remove-rows', index: 4, count: 2 }] }],
    )
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    expect(mutation.touchedEntries).toContain('xl/worksheets/sheet2.xml')
    const saved = await JSZip.loadAsync(mutation.buffer)
    const other = await saved.file('xl/worksheets/sheet2.xml')?.async('text')
    expect(other).toContain('<f>Data!A8</f>')
  })

  it('still fails closed when a cross-sheet reference is deleted', async () => {
    const zip = await JSZip.loadAsync(await buildStructureFixture())
    zip.file(
      'xl/worksheets/sheet2.xml',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>Data!A1</f><v>1</v></c></row></sheetData></worksheet>',
      { createFolders: false },
    )
    const buffer = await zip.generateAsync({ type: 'nodebuffer' })
    await expect(
      applyCellEditsToXlsx(
        buffer,
        [],
        [{ sheetName: SHEET, ops: [{ kind: 'remove-rows', index: 0, count: 1 }] }],
      ),
    ).rejects.toThrow('deleted range')
  })

  const anchoredTable =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:B3" totalsRowShown="0">' +
    '<autoFilter ref="A1:B3"/>' +
    '<tableColumns count="2"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns>' +
    '<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>' +
    '</table>'

  const anchoredChart =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart><c:plotArea><c:barChart><c:ser><c:f>Data!$A$1:$A$10</c:f></c:ser></c:barChart></c:plotArea></c:chart>' +
    '</c:chartSpace>'

  async function buildAnchoredFixture(): Promise<Buffer> {
    const zip = await JSZip.loadAsync(await buildStructureFixture())
    const sheet = (await zip.file('xl/worksheets/sheet1.xml')!.async('text'))
      .replace(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      )
      .replace(
        '</worksheet>',
        '<drawing r:id="rIdD"/><tableParts count="1"><tablePart r:id="rIdT"/></tableParts></worksheet>',
      )
    zip.file('xl/worksheets/sheet1.xml', sheet, { createFolders: false })
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rIdD" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
        '<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
        '</Relationships>',
      { createFolders: false },
    )
    zip.file('xl/drawings/drawing1.xml', DRAWING_XML, { createFolders: false })
    zip.file('xl/tables/table1.xml', anchoredTable, { createFolders: false })
    zip.file('xl/charts/chart1.xml', anchoredChart, { createFolders: false })
    return zip.generateAsync({ type: 'nodebuffer' })
  }

  it('a row insert inside the table shifts table, drawing anchors, and chart refs together', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildAnchoredFixture(),
      [],
      [{ sheetName: SHEET, ops: [{ kind: 'insert-rows', index: 1, count: 1 }] }],
    )
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    for (const path of [
      'xl/drawings/drawing1.xml',
      'xl/tables/table1.xml',
      'xl/charts/chart1.xml',
    ]) {
      expect(mutation.touchedEntries).toContain(path)
    }
    const zip = await JSZip.loadAsync(mutation.buffer)
    const table = await zip.file('xl/tables/table1.xml')?.async('text')
    expect(table).toContain('ref="A1:B4"')
    expect(table).toContain('<autoFilter ref="A1:B4"/>')
    const drawing = await zip.file('xl/drawings/drawing1.xml')?.async('text')
    expect(drawing).toContain('<xdr:row>6</xdr:row><xdr:rowOff>50</xdr:rowOff>')
    expect(drawing).toContain('<xdr:row>9</xdr:row><xdr:rowOff>60</xdr:rowOff>')
    expect(drawing).toContain('<xdr:row>7</xdr:row><xdr:rowOff>20</xdr:rowOff>')
    const chart = await zip.file('xl/charts/chart1.xml')?.async('text')
    expect(chart).toContain('<c:f>Data!$A$1:$A$11</c:f>')
  })

  it('a row deletion compresses covered drawings, leaves the table above alone, and clamps chart refs', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildAnchoredFixture(),
      [],
      [{ sheetName: SHEET, ops: [{ kind: 'remove-rows', index: 4, count: 5 }] }],
    )
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    expect(mutation.touchedEntries).not.toContain('xl/tables/table1.xml')
    const zip = await JSZip.loadAsync(mutation.buffer)
    const drawing = await zip.file('xl/drawings/drawing1.xml')?.async('text')
    // The shape sat entirely in the deleted rows: kept, compressed to the start.
    expect(drawing).toContain(
      '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>100</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
    )
    expect(drawing).toContain(
      '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>200</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>',
    )
    expect(drawing).toContain('name="Shape 2"')
    const chart = await zip.file('xl/charts/chart1.xml')?.async('text')
    expect(chart).toContain('<c:f>Data!$A$1:$A$5</c:f>')
    const table = await zip.file('xl/tables/table1.xml')?.async('text')
    expect(table).toContain('ref="A1:B3"')
  })

  it('a column insert left of the table shifts its ref and the drawing columns', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildAnchoredFixture(),
      [],
      [{ sheetName: SHEET, ops: [{ kind: 'insert-cols', index: 0, count: 1 }] }],
    )
    const zip = await JSZip.loadAsync(mutation.buffer)
    const table = await zip.file('xl/tables/table1.xml')?.async('text')
    expect(table).toContain('ref="B1:C3"')
    const drawing = await zip.file('xl/drawings/drawing1.xml')?.async('text')
    expect(drawing).toContain('<xdr:from><xdr:col>2</xdr:col>')
    expect(drawing).toContain('<xdr:to><xdr:col>4</xdr:col>')
  })

  it('still fails closed when a deletion would hit the table header row', async () => {
    await expect(
      applyCellEditsToXlsx(
        await buildAnchoredFixture(),
        [],
        [{ sheetName: SHEET, ops: [{ kind: 'remove-rows', index: 0, count: 1 }] }],
      ),
    ).rejects.toThrow('header row')
  })

  it('shifts fine when the rels put Target before Id (openpyxl attribute order)', async () => {
    const zip = await JSZip.loadAsync(await buildAnchoredFixture())
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Target="../drawings/drawing1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Id="rIdD"/>' +
        '<Relationship Target="../tables/table1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Id="rIdT"/>' +
        '</Relationships>',
      { createFolders: false },
    )
    const mutation = await applyCellEditsToXlsx(
      await zip.generateAsync({ type: 'nodebuffer' }),
      [],
      [{ sheetName: SHEET, ops: [{ kind: 'insert-rows', index: 1, count: 1 }] }],
    )
    expect(mutation.touchedEntries).toContain('xl/drawings/drawing1.xml')
    expect(mutation.touchedEntries).toContain('xl/tables/table1.xml')
    const out = await JSZip.loadAsync(mutation.buffer)
    expect(await out.file('xl/tables/table1.xml')?.async('text')).toContain('ref="A1:B4"')
  })

  it('fails closed when the referenced relationship is genuinely missing', async () => {
    const zip = await JSZip.loadAsync(await buildAnchoredFixture())
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
        '</Relationships>',
      { createFolders: false },
    )
    await expect(
      applyCellEditsToXlsx(
        await zip.generateAsync({ type: 'nodebuffer' }),
        [],
        [{ sheetName: SHEET, ops: [{ kind: 'insert-rows', index: 1, count: 1 }] }],
      ),
    ).rejects.toThrow('relationship is missing')
  })

  it('fails closed with a distinct message when the rels part itself is missing', async () => {
    const zip = await JSZip.loadAsync(await buildAnchoredFixture())
    zip.remove('xl/worksheets/_rels/sheet1.xml.rels')
    await expect(
      applyCellEditsToXlsx(
        await zip.generateAsync({ type: 'nodebuffer' }),
        [],
        [{ sheetName: SHEET, ops: [{ kind: 'insert-rows', index: 1, count: 1 }] }],
      ),
    ).rejects.toThrow('sheet1.xml.rels is missing')
  })
})

describe('axis attribute operations', () => {
  it('sets row heights on existing rows and creates missing rows in order', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'set-row-size', start: 1, end: 2, size: 24.75 }],
      SHEET,
    )
    expect(xml).toContain('<row r="2" ht="24.75" customHeight="1">')
    // Row 3 has no element in the fixture; it is created between rows 2 and 4.
    expect(xml).toMatch(/<row r="3" ht="24.75" customHeight="1"\/>\s*<row r="4">/)
    expect(xml).toContain('<row r="1" spans="1:4">')
  })

  it('resets row heights by dropping ht/customHeight without creating rows', async () => {
    const withHeights = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'set-row-size', start: 1, end: 2, size: 30 }],
      SHEET,
    )
    const xml = applyStructuralOps(
      withHeights,
      [{ kind: 'set-row-size', start: 1, end: 5, size: null }],
      SHEET,
    )
    expect(xml).not.toContain('ht="30"')
    expect(xml).not.toContain('customHeight')
    expect(xml).not.toContain('<row r="5"')
    expect(xml).toContain('<row r="2"><c r="A2">')
  })

  it('hides and unhides rows declaratively', async () => {
    const hidden = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'set-rows-hidden', start: 1, end: 3, hidden: true }],
      SHEET,
    )
    expect(hidden).toContain('<row r="2" hidden="1">')
    expect(hidden).toMatch(/<row r="3" hidden="1"\/>/)
    const unhidden = applyStructuralOps(
      hidden,
      [{ kind: 'set-rows-hidden', start: 1, end: 3, hidden: false }],
      SHEET,
    )
    expect(unhidden).not.toContain('hidden="1"')
  })

  it('splits an existing col range for a width change and keeps the rest verbatim', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'set-col-size', start: 2, end: 4, size: 12.5 }],
      SHEET,
    )
    // Original col covered B:C (min=2,max=3): B keeps width 20, C joins the new range.
    expect(xml).toContain('<col min="2" max="2" width="20" customWidth="1"/>')
    expect(xml).toContain('<col min="3" max="3" width="12.5" customWidth="1"/>')
    expect(xml).toContain('<col min="4" max="5" width="12.5" customWidth="1"/>')
  })

  it('creates the cols section before sheetData when absent and drops it when emptied', async () => {
    const bare = (await fixtureWorksheet()).replace(/<cols>[\s\S]*?<\/cols>/, '')
    const withWidth = applyStructuralOps(
      bare,
      [{ kind: 'set-col-size', start: 0, end: 0, size: 9 }],
      SHEET,
    )
    expect(withWidth).toMatch(
      /<cols><col min="1" max="1" width="9" customWidth="1"\/><\/cols>\s*<sheetData>/,
    )
    const reset = applyStructuralOps(
      withWidth,
      [{ kind: 'set-col-size', start: 0, end: 0, size: null }],
      SHEET,
    )
    expect(reset).not.toContain('<cols>')
  })

  it('hides columns while preserving unrelated col attributes', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'set-cols-hidden', start: 1, end: 2, hidden: true }],
      SHEET,
    )
    expect(xml).toContain('<col min="2" max="3" width="20" customWidth="1" hidden="1"/>')
  })

  it('applies in order with row inserts: a later height op uses post-insert coordinates', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [
        { kind: 'insert-rows', index: 0, count: 1 },
        { kind: 'set-row-size', start: 0, end: 0, size: 40 },
      ],
      SHEET,
    )
    // The height lands on the inserted screen row 1, not the shifted old row 1.
    expect(xml).toMatch(/<row r="1" ht="40" customHeight="1"\/>/)
    expect(xml).toContain('<row r="2" spans="1:4">')
  })

  it('sizing and visibility ops are allowed on sheets with tables or drawings', () => {
    const guarded =
      '<worksheet><sheetData><row r="1"/></sheetData><tableParts count="1"><tablePart r:id="rId2"/></tableParts><drawing r:id="rId3"/></worksheet>'
    const xml = applyStructuralOps(
      guarded,
      [{ kind: 'set-row-size', start: 0, end: 0, size: 20 }],
      SHEET,
    )
    expect(xml).toContain('<row r="1" ht="20" customHeight="1"/>')
  })
})

describe('outline operations', () => {
  it('sets row outline levels, creates missing rows, and syncs sheetFormatPr', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'set-rows-outline', start: 1, end: 2, level: 2 }],
      SHEET,
    )
    expect(xml).toContain('<row r="2" outlineLevel="2">')
    expect(xml).toMatch(/<row r="3" outlineLevel="2"\/>\s*<row r="4">/)
    expect(xml).toMatch(/<sheetFormatPr[^>]* outlineLevelRow="2"/)
  })

  it('removes the attribute at level 0 and clears the sheetFormatPr maximum', async () => {
    const grouped = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'set-rows-outline', start: 1, end: 2, level: 1 }],
      SHEET,
    )
    const ungrouped = applyStructuralOps(
      grouped,
      [{ kind: 'set-rows-outline', start: 1, end: 2, level: 0 }],
      SHEET,
    )
    expect(ungrouped).not.toContain('outlineLevel')
    expect(ungrouped).not.toContain('outlineLevelRow')
    // A level-0 reset applied directly never creates missing rows.
    const reset = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'set-rows-outline', start: 1, end: 2, level: 0 }],
      SHEET,
    )
    expect(reset).not.toContain('<row r="3"')
  })

  it('writes the collapsed flag only when the op carries one', async () => {
    const collapsed = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'set-rows-outline', start: 3, end: 3, level: 0, collapsed: true }],
      SHEET,
    )
    expect(collapsed).toContain('<row r="4" collapsed="1">')
    const levelOnly = applyStructuralOps(
      collapsed,
      [{ kind: 'set-rows-outline', start: 3, end: 3, level: 1 }],
      SHEET,
    )
    // A level-only op leaves the file's collapsed flag untouched.
    expect(levelOnly).toContain('collapsed="1"')
    expect(levelOnly).toContain('outlineLevel="1"')
  })

  it('sets column outline levels preserving unrelated col attributes', async () => {
    const xml = applyStructuralOps(
      await fixtureWorksheet(),
      [{ kind: 'set-cols-outline', start: 1, end: 2, level: 1 }],
      SHEET,
    )
    expect(xml).toContain('<col min="2" max="3" width="20" customWidth="1" outlineLevel="1"/>')
    expect(xml).toMatch(/<sheetFormatPr[^>]* outlineLevelCol="1"/)
  })

  it('creates a sheetFormatPr when the worksheet has none', () => {
    const bare = '<worksheet><sheetData><row r="1"/></sheetData></worksheet>'
    const xml = applyStructuralOps(
      bare,
      [{ kind: 'set-rows-outline', start: 0, end: 0, level: 1 }],
      SHEET,
    )
    expect(xml).toMatch(/<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"\/>\s*<sheetData>/)
  })

  it('outline ops are allowed on sheets with tables or drawings', () => {
    const guarded =
      '<worksheet><sheetData><row r="1"/></sheetData><tableParts count="1"><tablePart r:id="rId2"/></tableParts><drawing r:id="rId3"/></worksheet>'
    const xml = applyStructuralOps(
      guarded,
      [{ kind: 'set-rows-outline', start: 0, end: 0, level: 1 }],
      SHEET,
    )
    expect(xml).toContain('<row r="1" outlineLevel="1"/>')
  })
})
