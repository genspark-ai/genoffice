import { describe, expect, it } from 'vitest'
import { parseRange } from '../src/domain/cell-address'
import { refIntersects, renameRefSheet, splitSheetRef } from '../src/domain/chart-visual'

describe('splitSheetRef', () => {
  it('splits a plain absolute cell ref', () => {
    expect(splitSheetRef('Sheet1!$A$1')).toEqual({ sheetName: 'Sheet1', range: 'A1' })
    expect(splitSheetRef("'My Sheet'!$B$2:$B$13")).toEqual({
      sheetName: 'My Sheet',
      range: 'B2:B13',
    })
    expect(splitSheetRef("'It''s'!C3:C9")).toEqual({ sheetName: "It's", range: 'C3:C9' })
  })

  it('accepts a whole-column ref, whose c:f form Excel writes', () => {
    expect(splitSheetRef('Data!$A:$A')).toEqual({ sheetName: 'Data', range: 'A:A' })
    expect(splitSheetRef('Data!A:A')).toEqual({ sheetName: 'Data', range: 'A:A' })
    expect(splitSheetRef('Data!$A:C')).toEqual({ sheetName: 'Data', range: 'A:C' })
    expect(splitSheetRef("'My Sheet'!$A:$C")).toEqual({
      sheetName: 'My Sheet',
      range: 'A:C',
    })
  })

  it('accepts a whole-row ref the same way', () => {
    expect(splitSheetRef('Data!$1:$1')).toEqual({ sheetName: 'Data', range: '1:1' })
    expect(splitSheetRef("'My Sheet'!2:5")).toEqual({ sheetName: 'My Sheet', range: '2:5' })
    expect(splitSheetRef('Data!$1:$A')).toBeNull()
  })

  it('accepts a lowercase ref and normalises the range to the case parseRange wants', () => {
    expect(splitSheetRef('Data!$a$1')).toEqual({ sheetName: 'Data', range: 'A1' })
    expect(splitSheetRef('Data!$a$1:$b$13')).toEqual({ sheetName: 'Data', range: 'A1:B13' })
    expect(splitSheetRef('Data!$a:$a')).toEqual({ sheetName: 'Data', range: 'A:A' })
    // parseAddress is case-strict by contract, so normalising here is what
    // makes the lowercase support usable downstream rather than half-supported.
    expect(parseRange(splitSheetRef('Data!$a$1:$b$13')!.range)).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 12,
      endColumn: 1,
    })
    expect(() => parseRange('a1:b13')).toThrow('Invalid cell address')
  })

  it('leaves a defined name alone: it names no sheet', () => {
    expect(splitSheetRef('Data!Q1Sales')).toBeNull()
    expect(splitSheetRef('Q1Sales')).toBeNull()
  })

  it('rejects a cross-sheet ref and a bare range with no sheet qualifier', () => {
    // A chart c:f is qualified by exactly one sheet.
    expect(splitSheetRef('Sheet1!Sheet2!$A$1')).toBeNull()
    expect(splitSheetRef('$A$1')).toBeNull()
  })

  it('cannot match mid-identifier text', () => {
    expect(splitSheetRef('not-a-ref')).toBeNull()
    expect(splitSheetRef('Data!A1extra')).toBeNull()
    expect(splitSheetRef('Data!$A:$A1')).toBeNull()
    expect(splitSheetRef('Data!')).toBeNull()
    // A name that merely starts like a ref must not be split.
    expect(splitSheetRef('Data!A')).toBeNull()
  })
})

describe('renameRefSheet', () => {
  it('rewrites the sheet qualifier and leaves the range text byte-for-byte', () => {
    expect(renameRefSheet('Sheet1!$A$1', 'Sheet1', 'Renamed')).toBe("'Renamed'!$A$1")
    expect(renameRefSheet("'My Sheet'!$B$2:$B$13", 'My Sheet', 'New')).toBe("'New'!$B$2:$B$13")
    expect(renameRefSheet('Data!$A$1', 'Other', 'New')).toBe('Data!$A$1')
  })

  it('rewrites a whole-column ref, which used to stay dangling', () => {
    // Regression: the old pattern returned null here, so renameRefSheet passed
    // the ref through untouched and c:f kept pointing at the old sheet name.
    expect(renameRefSheet('Data!$A:$A', 'Data', 'Renamed')).toBe("'Renamed'!$A:$A")
    expect(renameRefSheet('Data!A:C', 'Data', 'Renamed')).toBe("'Renamed'!A:C")
  })

  it('rewrites a whole-row ref', () => {
    expect(renameRefSheet('Data!$1:$5', 'Data', 'Renamed')).toBe("'Renamed'!$1:$5")
  })

  it('rewrites a lowercase ref', () => {
    expect(renameRefSheet('Data!$a$1', 'Data', 'Renamed')).toBe("'Renamed'!$a$1")
    expect(renameRefSheet('Data!$a:$a', 'Data', 'Renamed')).toBe("'Renamed'!$a:$a")
  })

  it('escapes an apostrophe in the new sheet name', () => {
    expect(renameRefSheet('Data!$A$1', 'Data', "It's")).toBe("'It''s'!$A$1")
  })

  it('leaves a defined-name ref alone', () => {
    expect(renameRefSheet('Data!Q1Sales', 'Data', 'Renamed')).toBe('Data!Q1Sales')
  })
})

describe('refIntersects after a ref-shape fix', () => {
  const bounds = { startRow: 5, endRow: 5, startColumn: 0, endColumn: 0 }

  it('intersects a lowercase ref now that the range is normalised', () => {
    expect(refIntersects('Data!$a$1', 'Data', bounds)).toBe(false)
    expect(refIntersects('Data!$a$6', 'Data', bounds)).toBe(true)
  })

  it('still reports no intersection for another sheet', () => {
    expect(refIntersects('Data!$A$6', 'Other', bounds)).toBe(false)
  })

  it('still reports no intersection for a whole-column ref it cannot bound', () => {
    // parseRange has no row extent for $A:$A, so the existing catch holds and
    // this stays false rather than throwing.
    expect(refIntersects('Data!$A:$A', 'Data', bounds)).toBe(false)
  })
})
