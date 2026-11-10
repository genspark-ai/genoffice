import { describe, expect, it } from 'vitest'

import { applyStructuralOps, inferWorksheetAddresses } from '../src/gateway/xlsx-structure'

function worksheet(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`
}

describe('inferWorksheetAddresses with self-closing rows and cells', () => {
  it('leaves a custom-height empty row self-closing', () => {
    const xml = worksheet(
      '<row r="1"><c r="A1"><v>1</v></c></row>' +
        '<row r="2" ht="20" customHeight="1"/>' +
        '<row><c><v>3</v></c></row>',
    )
    const inferred = inferWorksheetAddresses(xml)
    expect(inferred).toContain('<row r="2" ht="20" customHeight="1"/>')
    expect(inferred).not.toContain('/></row>')
    expect(inferred).toContain('<row r="3"><c r="A3"><v>3</v></c></row>')
  })

  it('infers the address of a styled empty cell without closing it twice', () => {
    const xml = worksheet(
      '<row r="1"><c r="A1"><v>1</v></c></row><row><c s="1"/><c><v>2</v></c></row>',
    )
    const inferred = inferWorksheetAddresses(xml)
    expect(inferred).toContain('<row r="2"><c r="A2" s="1"/><c r="B2"><v>2</v></c></row>')
    expect(inferred).not.toContain('/></c>')
  })

  it('keeps an addressed self-closing row untouched when another row needs inference', () => {
    const xml = worksheet('<row r="1" hidden="1"/><row><c s="2"/></row>')
    const inferred = inferWorksheetAddresses(xml)
    expect(inferred).toBe(worksheet('<row r="1" hidden="1"/><row r="2"><c r="A2" s="2"/></row>'))
  })

  it('replays a structural op over inferred self-closing rows and cells', () => {
    const xml = worksheet(
      '<row r="1" ht="20" customHeight="1"/><row><c s="1"/><c><v>2</v></c></row>',
    )
    const shifted = applyStructuralOps(xml, [{ kind: 'insert-rows', index: 0, count: 1 }], 'Data')
    expect(shifted).toContain('<row r="2" ht="20" customHeight="1"/>')
    expect(shifted).toContain('<c r="A3" s="1"/><c r="B3"><v>2</v></c>')
    expect(shifted).not.toContain('/></row>')
    expect(shifted).not.toContain('/></c>')
  })
})
