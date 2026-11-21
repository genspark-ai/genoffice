/** Word's Symbol dialog: the grid's characters and the character a typed code stands for. */
import { describe, expect, it } from 'vitest'
import { SYMBOL_CHARS, symbolCode, symbolFromCode } from '../src/renderer/components/SymbolDialog'

describe('symbol dialog', () => {
  it('lists each character once and none that draws nothing', () => {
    expect(new Set(SYMBOL_CHARS).size).toBe(SYMBOL_CHARS.length)
    expect(SYMBOL_CHARS).toContain('©')
    expect(SYMBOL_CHARS).not.toContain('­')
  })

  it('reads a typed character code and writes the code of a picked character', () => {
    expect(symbolFromCode('00A9')).toBe('©')
    expect(symbolFromCode('u+2192')).toBe('→')
    expect(symbolFromCode('1F600')).toBe('\u{1f600}')
    expect(symbolFromCode('')).toBeNull()
    expect(symbolFromCode('xyz')).toBeNull()
    expect(symbolFromCode('0009')).toBeNull()
    expect(symbolFromCode('D800')).toBeNull()
    expect(symbolFromCode('110000')).toBeNull()
    expect(symbolCode('©')).toBe('00A9')
    expect(symbolCode('\u{1f600}')).toBe('1F600')
  })
})
