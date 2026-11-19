import { describe, expect, it } from 'vitest'

import { appendSymbol } from '../src/renderer/SymbolDialog'

describe('appendSymbol', () => {
  it('puts the symbol alone into an empty cell', () => {
    expect(appendSymbol(null, 'Ω')).toBe('Ω')
    expect(appendSymbol(undefined, '€')).toBe('€')
    expect(appendSymbol('', '±')).toBe('±')
  })

  it('appends to existing text', () => {
    expect(appendSymbol('Total', '™')).toBe('Total™')
    expect(appendSymbol('α + β', '→')).toBe('α + β→')
  })

  it('coerces numbers and booleans to text before appending', () => {
    expect(appendSymbol(42, '°')).toBe('42°')
    expect(appendSymbol(3.14, 'π')).toBe('3.14π')
    expect(appendSymbol(true, '✓')).toBe('true✓')
  })
})
