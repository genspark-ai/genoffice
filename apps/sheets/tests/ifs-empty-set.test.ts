import { ArrayValueObject, NumberValueObject, StringValueObject } from '@univerjs/engine-formula'
import { describe, expect, it } from 'vitest'

import { createIfsEmptySetExecutors, fixIfsEmptySetResult } from '../src/renderer/ifs-empty-set'

// Univer's zero-match branch: ArrayValueObject.create('0') parses the string
// as an array literal and brace-strips it to '' — a 1x1 empty-string array.
const artifact = () => ArrayValueObject.create('0')

describe('fixIfsEmptySetResult', () => {
  it('rewrites the 1x1 empty-string artifact to 0', () => {
    const fixed = fixIfsEmptySetResult(artifact() as never) as NumberValueObject
    expect(fixed.isArray?.()).toBeFalsy()
    expect(fixed.getValue()).toBe(0)
  })

  it('passes real numbers and errors through untouched', () => {
    const seven = NumberValueObject.create(7)
    expect(fixIfsEmptySetResult(seven as never)).toBe(seven)
    const text = StringValueObject.create('')
    expect(fixIfsEmptySetResult(text as never)).toBe(text)
  })

  it('rewrites artifact slots inside an array-criteria expansion', () => {
    const outer = ArrayValueObject.create({
      calculateValueList: [[NumberValueObject.create(5), artifact()]],
      rowCount: 1,
      columnCount: 2,
      unitId: '',
      sheetId: '',
      row: 0,
      column: 0,
    })
    const fixed = fixIfsEmptySetResult(outer as never) as ArrayValueObject
    const values = fixed.getArrayValue()
    expect((values[0]![0] as NumberValueObject).getValue()).toBe(5)
    expect((values[0]![1] as NumberValueObject).getValue()).toBe(0)
    expect((values[0]![1] as NumberValueObject).isArray?.()).toBeFalsy()
  })

  it('leaves clean arrays untouched', () => {
    const outer = ArrayValueObject.create('{1;2;3}')
    expect(fixIfsEmptySetResult(outer as never)).toBe(outer)
  })
})

describe('createIfsEmptySetExecutors', () => {
  it('builds MINIFS and MAXIFS wrappers on the current Univer bundle', () => {
    const executors = createIfsEmptySetExecutors()
    expect(executors.map((executor) => executor.name).sort()).toEqual(['MAXIFS', 'MINIFS'])
  })
})
