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
