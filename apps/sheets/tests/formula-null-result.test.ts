import {
  ArrayValueObject,
  ErrorValueObject,
  NullValueObject,
  NumberValueObject,
  StringValueObject,
} from '@univerjs/engine-formula'
import { describe, expect, it } from 'vitest'

import { coerceNullResult } from '../src/renderer/formula-null-result'

describe('coerceNullResult', () => {
  it('turns a scalar null result into 0', () => {
    const result = coerceNullResult(NullValueObject.create())
    expect(result).toBeInstanceOf(NumberValueObject)
    expect((result as NumberValueObject).getValue()).toBe(0)
  })

  // `=Sheet2!B3` where B3 holds no value: Excel shows 0 (serial 0 under a
  // date format renders 1900/1/0), a NullValueObject root left the cell
  // blank.
  it('turns a single-cell reference resolving to an empty cell into 0', () => {
    const reference = {
      isReferenceObject: () => true,
      getRangePosition: () => ({ startRow: 2, endRow: 2, startColumn: 1, endColumn: 1 }),
      getFirstCell: () => NullValueObject.create(),
    }
    const result = coerceNullResult(reference)
    expect(result).toBeInstanceOf(NumberValueObject)
    expect((result as NumberValueObject).getValue()).toBe(0)
  })

  it('leaves multi-cell and value-bearing references alone', () => {
    const multi = {
      isReferenceObject: () => true,
      getRangePosition: () => ({ startRow: 0, endRow: 4, startColumn: 1, endColumn: 1 }),
      getFirstCell: () => NullValueObject.create(),
    }
    expect(coerceNullResult(multi)).toBe(multi)
    const filled = {
      isReferenceObject: () => true,
      getRangePosition: () => ({ startRow: 2, endRow: 2, startColumn: 1, endColumn: 1 }),
      getFirstCell: () => NumberValueObject.create(7),
    }
    expect(coerceNullResult(filled)).toBe(filled)
  })

  it('leaves real values, errors, arrays, and references alone', () => {
    const number = NumberValueObject.create(42)
    expect(coerceNullResult(number)).toBe(number)
    const text = StringValueObject.create('NA')
    expect(coerceNullResult(text)).toBe(text)
    const error = ErrorValueObject.create('#DIV/0!' as never)
    expect(coerceNullResult(error)).toBe(error)
    const array = ArrayValueObject.createByArray([[null, 1]])
    expect(coerceNullResult(array)).toBe(array)
    expect(coerceNullResult(null)).toBe(null)
    expect(coerceNullResult(undefined)).toBe(undefined)
  })
})
