/**
 * The formula stream hold vetoes Univer's calculation kick-off while file
 * chunks land and releases one merged cycle afterwards; the flush must only
 * dispatch when the accumulated dirty data actually holds something.
 */
import { describe, expect, it } from 'vitest'

import { hasDirtyData } from '../src/renderer/formula-stream-hold'

describe('hasDirtyData', () => {
  it('treats empty maps and arrays as nothing to calculate', () => {
    expect(hasDirtyData(undefined)).toBe(false)
    expect(
      hasDirtyData({
        dirtyRanges: [],
        dirtyNameMap: {},
        dirtyDefinedNameMap: {},
        dirtyUnitFeatureMap: {},
        clearDependencyTreeCache: {},
        forceCalculation: false,
      }),
    ).toBe(false)
  })

  it('detects ranges, name maps and the force flag', () => {
    expect(hasDirtyData({ dirtyRanges: [{ unitId: 'u', sheetId: 's', range: {} }] })).toBe(true)
    expect(hasDirtyData({ dirtyNameMap: { u: { s: true } } })).toBe(true)
    expect(hasDirtyData({ forceCalculation: true })).toBe(true)
  })
})
