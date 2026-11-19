import { ArrayValueObject, NumberValueObject, StringValueObject } from '@univerjs/engine-formula'
import type { BaseValueObject } from '@univerjs/engine-formula'
import { afterEach, describe, expect, it } from 'vitest'

import { installCriteriaCompareCacheFix } from '../src/renderer/criteria-compare-cache'

// The inverted-index cache is a module singleton keyed by unit/sheet/column;
// unique ids per case keep runs independent.
let unit = 0

function column(values: readonly (string | number)[], columnIndex: number, unitId: string) {
  return ArrayValueObject.create({
    calculateValueList: values.map((value) => [
      typeof value === 'number'
        ? (NumberValueObject.create(value) as BaseValueObject)
        : (StringValueObject.create(value) as BaseValueObject),
    ]),
    rowCount: values.length,
    columnCount: 1,
    unitId,
    sheetId: 'sheet1',
    row: 4,
    column: columnIndex,
  } as unknown as Parameters<typeof ArrayValueObject.create>[0]) as ArrayValueObject
}

/// The number-sensitive repair pass SUMIFS runs after the compare
/// (filterSameValueObjectResult) only visits materialized slots; emulate it
/// to show the end-to-end criteria verdict.
function repaired(result: ArrayValueObject, range: ArrayValueObject, criteria: BaseValueObject) {
  const mapped = (
    result as unknown as {
      mapValue(
        map: (value: BaseValueObject, r: number, c: number) => BaseValueObject,
      ): ArrayValueObject
    }
  ).mapValue((value, r, c) => {
    const cell = (range as unknown as { get(r: number, c: number): BaseValueObject }).get(r, c)
    if (cell?.isString() && criteria.isNumber()) {
      const coerced = cell.convertToNumberObjectValue()
      if (coerced.isNumber()) return coerced.compare(criteria, '=' as never)
    }
    return value
  })
  return mapped.getArrayValue().map((row) => row[0]?.getValue() === true)
}

const disposables: { dispose(): void }[] = []

afterEach(() => {
  for (const disposable of disposables.splice(0)) disposable.dispose()
})

describe('installCriteriaCompareCacheFix', () => {
  it('keeps number-coerced criteria matching text cells on cached re-compares', () => {
    disposables.push(installCriteriaCompareCacheFix())
    const unitId = `wb-fix-${unit++}`
    const months = ['2026-06', '2026-05', '2026-06', '2026-06', '2026-05']
    // findCompareToken('2026-06') hands SUMIFS this date serial.
    const criteria = NumberValueObject.create(46174) as BaseValueObject
    const expected = [true, false, true, true, false]
    for (let call = 0; call < 3; call++) {
      const range = column(months, 7, unitId)
      const result = range.compare(criteria, '=' as never) as ArrayValueObject
      expect(repaired(result, range, criteria), `call ${call}`).toEqual(expected)
    }
  })

  it('reproduces the sparse-cache miss without the fix', () => {
    const unitId = `wb-bug-${unit++}`
    const months = ['2026-06', '2026-05', '2026-06']
    const criteria = NumberValueObject.create(46174) as BaseValueObject
    const first = column(months, 7, unitId)
    expect(
      repaired(first.compare(criteria, '=' as never) as ArrayValueObject, first, criteria),
    ).toEqual([true, false, true])
    const second = column(months, 7, unitId)
    expect(
      repaired(second.compare(criteria, '=' as never) as ArrayValueObject, second, criteria),
    ).toEqual([false, false, false])
  })

  it('keeps the cache fast path for same-type compares', () => {
    disposables.push(installCriteriaCompareCacheFix())
    const unitId = `wb-same-${unit++}`
    const categories = ['Housing', 'Utilities', 'Housing']
    const criteria = StringValueObject.create('Utilities') as BaseValueObject
    // The cached path returns a sparse array (positives + default false);
    // read through get() like the consumers do.
    const verdicts = (result: ArrayValueObject, rows: number) =>
      Array.from(
        { length: rows },
        (_unused, r) =>
          (result as unknown as { get(r: number, c: number): BaseValueObject | null })
            .get(r, 0)
            ?.getValue() === true,
      )
    for (let call = 0; call < 2; call++) {
      const range = column(categories, 2, unitId)
      const result = range.compare(criteria, '=' as never) as ArrayValueObject
      expect(verdicts(result, 3), `call ${call}`).toEqual([false, true, false])
    }
    const numbers = [10, 20, 10]
    const numberCriteria = NumberValueObject.create(10) as BaseValueObject
    for (let call = 0; call < 2; call++) {
      const range = column(numbers, 3, unitId)
      const result = range.compare(numberCriteria, '=' as never) as ArrayValueObject
      expect(verdicts(result, 3), `numeric call ${call}`).toEqual([true, false, true])
    }
  })

  it('dispose restores the original compare', () => {
    const proto = ArrayValueObject.prototype as unknown as Record<string, unknown>
    const original = proto._batchOperatorValue
    const fix = installCriteriaCompareCacheFix()
    expect(proto._batchOperatorValue).not.toBe(original)
    fix.dispose()
    expect(proto._batchOperatorValue).toBe(original)
  })
})
