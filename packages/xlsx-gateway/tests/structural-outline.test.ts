import { describe, expect, it } from 'vitest'

import { applyStructuralOps, type StructuralOp } from '../src/gateway/xlsx-structure'

const SHEET = 'Data'

function worksheet(prefix: string, rows: string, cols = ''): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${prefix}<dimension ref="A1:D10"/><sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rows}</sheetData></worksheet>`
}

const ROWS =
  '<row r="1"><c r="A1"><v>1</v></c></row><row r="2" ht="20" customHeight="1"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>'

describe('row/column outline attributes', () => {
  it('writes outlineLevel on existing rows, materializes missing ones and syncs sheetFormatPr', () => {
    const ops: StructuralOp[] = [{ kind: 'set-rows-outline', start: 0, end: 3, level: 2 }]
    const xml = applyStructuralOps(worksheet('', ROWS), ops, SHEET)
    expect(xml).toContain('<row r="1" outlineLevel="2">')
    expect(xml).toContain('<row r="2" ht="20" customHeight="1" outlineLevel="2">')
    expect(xml).toContain('<row r="4" outlineLevel="2"/>')
    expect(xml).toContain('<sheetFormatPr defaultRowHeight="15" outlineLevelRow="2"/>')
  })

  it('level 0 strips the attribute and the collapsed flag follows the op', () => {
    const base = worksheet(
      '',
      '<row r="5" outlineLevel="1" collapsed="1"><c r="A5"><v>1</v></c></row>',
    )
    const cleared = applyStructuralOps(
      base,
      [{ kind: 'set-rows-outline', start: 4, end: 4, level: 0, collapsed: false }],
      SHEET,
    )
    expect(cleared).toContain('<row r="5">')
    expect(cleared).not.toContain('outlineLevelRow')
    const untouched = applyStructuralOps(
      base,
      [{ kind: 'set-rows-outline', start: 4, end: 4, level: 3 }],
      SHEET,
    )
    expect(untouched).toContain('<row r="5" collapsed="1" outlineLevel="3">')
  })

  it('splits <col> spans around the grouped columns and records outlineLevelCol', () => {
    const base = worksheet(
      '',
      ROWS,
      '<cols><col min="1" max="6" width="12" customWidth="1"/></cols>',
    )
    const xml = applyStructuralOps(
      base,
      [{ kind: 'set-cols-outline', start: 1, end: 2, level: 1, collapsed: true }],
      SHEET,
    )
    expect(xml).toContain('<col min="1" max="1" width="12" customWidth="1"/>')
    expect(xml).toContain(
      '<col min="2" max="3" width="12" customWidth="1" outlineLevel="1" collapsed="1"/>',
    )
    expect(xml).toContain('<col min="4" max="6" width="12" customWidth="1"/>')
    expect(xml).toContain('outlineLevelCol="1"')
  })
})

describe('sheetPr/outlinePr summary placement', () => {
  const op = (summaryBelow: boolean, summaryRight: boolean): StructuralOp => ({
    kind: 'set-outline-pr',
    summaryBelow,
    summaryRight,
  })

  it('creates sheetPr and outlinePr when the summary moves above/left', () => {
    const xml = applyStructuralOps(worksheet('', ROWS), [op(false, true)], SHEET)
    expect(xml).toContain('<sheetPr><outlinePr summaryBelow="0"/></sheetPr><dimension')
  })

  it('keeps other sheetPr children and outlinePr attributes, in schema order', () => {
    const base = worksheet(
      '<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>',
      ROWS,
    )
    const xml = applyStructuralOps(base, [op(false, false)], SHEET)
    expect(xml).toContain(
      '<sheetPr><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0" summaryRight="0"/><pageSetUpPr fitToPage="1"/></sheetPr>',
    )
    const existing = worksheet(
      '<sheetPr><outlinePr applyStyles="1" summaryBelow="0"/></sheetPr>',
      ROWS,
    )
    expect(applyStructuralOps(existing, [op(true, false)], SHEET)).toContain(
      '<sheetPr><outlinePr applyStyles="1" summaryRight="0"/></sheetPr>',
    )
  })

  it('drops an outlinePr (and an emptied sheetPr) at the defaults', () => {
    const base = worksheet('<sheetPr><outlinePr summaryBelow="0"/></sheetPr>', ROWS)
    const xml = applyStructuralOps(base, [op(true, true)], SHEET)
    expect(xml).not.toContain('sheetPr')
    expect(applyStructuralOps(worksheet('', ROWS), [op(true, true)], SHEET)).not.toContain(
      'sheetPr',
    )
    const withTab = worksheet(
      '<sheetPr><tabColor rgb="FFFF0000"/><outlinePr summaryRight="0"/></sheetPr>',
      ROWS,
    )
    expect(applyStructuralOps(withTab, [op(true, true)], SHEET)).toContain(
      '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>',
    )
  })
})
