import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, type CellEdit } from '../src/gateway/xlsx-gateway'
import { translateSharedFormula } from '../src/gateway/xlsx-structure'

describe('translateSharedFormula', () => {
  it('shifts relative references by the follower offset', () => {
    expect(translateSharedFormula('SUM(F6:F10095)', 0, 7)).toBe('SUM(M6:M10095)')
    expect(translateSharedFormula('D3+1', 2, 0)).toBe('D5+1')
  })

  it('keeps $-anchored components in place', () => {
    expect(translateSharedFormula('$A$1+B2', 3, 3)).toBe('$A$1+E5')
    expect(translateSharedFormula('A$1+$B2', 3, 3)).toBe('D$1+$B5')
  })

  it('shifts whole-column and whole-row ranges', () => {
    expect(translateSharedFormula('SUM(B:B)', 0, 2)).toBe('SUM(D:D)')
    expect(translateSharedFormula('SUM(3:4)', 2, 0)).toBe('SUM(5:6)')
  })

  it('leaves string literals and function names alone', () => {
    expect(translateSharedFormula('IF(A1="B2","C3",LOG10(D4))', 1, 0)).toBe(
      'IF(A2="B2","C3",LOG10(D5))',
    )
  })

  it('keeps sheet-qualified relative references moving', () => {
    expect(translateSharedFormula('Other!C3*2', 1, 1)).toBe('Other!D4*2')
  })

  it('returns null when a shifted reference leaves the sheet', () => {
    expect(translateSharedFormula('A1+B1', -1, 0)).toBeNull()
  })
})

/// Row 2 mirrors the production corruption: F2 is the shared master of
/// F2:P2; the edit rewrites F2..L2, and the untouched M2..P2 followers must
/// come out as plain translated formulas, not orphaned si references.
describe('shared formula group materialization on save', () => {
  it('materializes untouched followers when their master is rewritten', async () => {
    const source = await buildSharedFormulaFixture()
    const edits: CellEdit[] = ['F', 'G', 'H', 'I', 'J', 'K', 'L'].map((column, index) => ({
      sheetName: 'Data',
      row: 1,
      column: 5 + index,
      writeValue: true,
      cell: { value: null, formula: `=SUM(${column}6:${column}10)` },
    }))
    const mutation = await applyCellEditsToXlsx(source, edits)
    const zip = await JSZip.loadAsync(mutation.buffer)
    const worksheet = await zip.file('xl/worksheets/sheet1.xml')?.async('string')
    expect(worksheet).toContain('<c r="M2" s="5"><f>SUM(M6:M10)</f><v>7</v></c>')
    expect(worksheet).toContain('<c r="N2" s="5"><f>SUM(N6:N10)</f><v>8</v></c>')
    expect(worksheet).toContain('<c r="F2" s="5"><f>SUM(F6:F10)</f></c>')
    expect(worksheet).not.toContain('t="shared"')
  })

  it('keeps the group intact when only followers are edited', async () => {
    const source = await buildSharedFormulaFixture()
    const mutation = await applyCellEditsToXlsx(source, [
      {
        sheetName: 'Data',
        row: 1,
        column: 6,
        writeValue: true,
        cell: { value: null, formula: '=SUM(G6:G10)' },
      },
    ])
    const zip = await JSZip.loadAsync(mutation.buffer)
    const worksheet = await zip.file('xl/worksheets/sheet1.xml')?.async('string')
    expect(worksheet).toContain('ref="F2:N2"')
    expect(worksheet).toContain('<c r="M2" s="5"><f t="shared" si="0"/><v>7</v></c>')
  })
})

async function buildSharedFormulaFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )
  const followers = ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
    .map((column, index) => `<c r="${column}2" s="5"><f t="shared" si="0"/><v>${index + 1}</v></c>`)
    .join('')
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:N10"/>
  <sheetData><row r="2"><c r="F2" s="5"><f t="shared" ref="F2:N2" si="0">SUM(F6:F10)</f><v>0</v></c>${followers}<c r="P2"><f>SUM(P6:P10)</f><v>9</v></c></row></sheetData>
</worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/// A follower whose translation would leave the sheet must not stay a bare
/// shared tag once its master is gone — it degrades to its cached value.
it('drops the shared tag when a follower cannot be translated', async () => {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )
  // Master references A1 relatively from B2; the follower at A2 would shift
  // the reference off-sheet (column -1) — translation fails by design.
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C2"/>
  <sheetData><row r="2"><c r="A2"><f t="shared" si="0"/><v>5</v></c><c r="B2"><f t="shared" ref="A2:B2" si="0">A1*2</f><v>6</v></c></row></sheetData>
</worksheet>`,
  )
  const source = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  const mutation = await applyCellEditsToXlsx(source, [
    {
      sheetName: 'Data',
      row: 1,
      column: 1,
      writeValue: true,
      cell: { value: null, formula: '=A1*3' },
    },
  ])
  const out = await JSZip.loadAsync(mutation.buffer)
  const worksheet = await out.file('xl/worksheets/sheet1.xml')?.async('string')
  expect(worksheet).toContain('<c r="A2"><v>5</v></c>')
  expect(worksheet).not.toContain('t="shared"')
})
