import { describe, expect, it } from 'vitest'

import { isNumericIdentifierText } from '../src/renderer/cell-warning'

describe('isNumericIdentifierText', () => {
  it('recognizes phone numbers and long digit-only identifiers', () => {
    expect(isNumericIdentifierText('13800138000')).toBe(true)
    expect(isNumericIdentifierText(' 0123456 ')).toBe(true)
  })

  it('keeps warnings for short numbers and numeric expressions', () => {
    expect(isNumericIdentifierText('007')).toBe(false)
    expect(isNumericIdentifierText('123.45')).toBe(false)
    expect(isNumericIdentifierText('20%')).toBe(false)
    expect(isNumericIdentifierText('+8613800138000')).toBe(false)
  })
})
