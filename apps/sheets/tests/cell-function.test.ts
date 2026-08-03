import { BaseFunction, ErrorValueObject, StringValueObject } from '@univerjs/engine-formula'
import { describe, expect, it, vi } from 'vitest'

import { CellWithFilename } from '../src/renderer/cell-function'

function infoType(value: string) {
  return StringValueObject.create(value)
}

function refOn(sheetId: string | undefined, defaultSheetId?: string) {
  return {
    isError: () => false,
    isReferenceObject: () => true,
    getForcedSheetId: () => sheetId,
    getDefaultSheetId: () => defaultSheetId,
  }
}

function makeCell(path: string | null, original?: BaseFunction) {
  const cell = new CellWithFilename('CELL', original, () => path)
  cell.setRefInfo('unit1', 'sheet-a', 3, 2)
  cell.setSheetsInfo({
    sheetOrder: ['sheet-a', 'sheet-b'],
    sheetNameMap: { 'sheet-a': 'COUNT Functions', 'sheet-b': 'Sheet 2' },
  } as never)
  return cell
}

describe('CELL("filename")', () => {
  it('returns dir/[book]sheet for the referenced sheet', () => {
    const result = makeCell('/Users/x/Downloads/Workshop v1.xlsx').calculate(
      infoType('filename'),
      refOn(undefined, 'sheet-a') as never,
    ) as StringValueObject
    expect(result.getValue()).toBe('/Users/x/Downloads/[Workshop v1.xlsx]COUNT Functions')
  })

  it('uses the explicit sheet of a cross-sheet reference', () => {
    const result = makeCell('/tmp/a.xlsx').calculate(
      infoType('FILENAME'),
      refOn('sheet-b') as never,
    ) as StringValueObject
    expect(result.getValue()).toBe('/tmp/[a.xlsx]Sheet 2')
  })

  it('falls back to the calling cell sheet when ref is omitted', () => {
    const result = makeCell('/tmp/a.xlsx').calculate(infoType('filename')) as StringValueObject
    expect(result.getValue()).toBe('/tmp/[a.xlsx]COUNT Functions')
  })

  it('returns an empty string for never-saved workbooks', () => {
    const result = makeCell(null).calculate(
      infoType('filename'),
      refOn(undefined, 'sheet-a') as never,
    ) as StringValueObject
    expect(result.getValue()).toBe('')
  })

  it('handles Windows separators', () => {
    const result = makeCell('C:\\Users\\x\\a.xlsx').calculate(
      infoType('filename'),
    ) as StringValueObject
    expect(result.getValue()).toBe('C:\\Users\\x\\[a.xlsx]COUNT Functions')
  })

  it('propagates info_type errors', () => {
    const error = ErrorValueObject.create('#NAME?' as never)
    expect(makeCell('/tmp/a.xlsx').calculate(error)).toBe(error)
  })
})

describe('CELL delegation', () => {
  it('routes every other info_type to the builtin executor', () => {
    const original = { calculate: vi.fn(() => StringValueObject.create('$B$4')) }
    const ref = refOn(undefined, 'sheet-a')
    const result = makeCell('/tmp/a.xlsx', original as never).calculate(
      infoType('address'),
      ref as never,
    )
    expect(original.calculate).toHaveBeenCalledWith(expect.anything(), ref)
    expect((result as StringValueObject).getValue()).toBe('$B$4')
  })

  it('returns #N/A for ref-less info_types other than filename', () => {
    const original = { calculate: vi.fn() }
    const result = makeCell('/tmp/a.xlsx', original as never).calculate(infoType('row'))
    expect((result as ErrorValueObject).getValue()).toBe('#N/A')
    expect(original.calculate).not.toHaveBeenCalled()
  })
})
