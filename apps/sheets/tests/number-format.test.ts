import { describe, expect, it } from 'vitest'

import { categoryOptionForPattern, numberFormatCategories } from '../src/renderer/number-format'

describe('numberFormatCategories', () => {
  it('every category pattern echoes back as its own dropdown option', () => {
    for (const category of numberFormatCategories()) {
      expect(categoryOptionForPattern(category.pattern)).toBe(category.label)
    }
  })
})

describe('categoryOptionForPattern', () => {
  it.each([
    ['', 'General'],
    ['m/d/yy', 'Short Date'],
    ['yyyy-mm-dd', 'Short Date'],
    ['dddd, mmmm dd, yyyy', 'Long Date'],
    ['[$-409]d-mmmm-yy', 'Long Date'],
    ['0.0', 'Number'],
    ['0.00_);[Red](0.00)', 'Number'],
    ['#,##0 "units"', 'Number'],
  ])('%s → %s', (pattern, label) => {
    expect(categoryOptionForPattern(pattern)).toBe(label)
  })
})
