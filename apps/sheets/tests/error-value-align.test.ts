import { CellValueType, HorizontalAlign } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { centeredErrorStyle } from '../src/renderer/error-value-align'
import { nonTextDisplayLabel } from '../src/renderer/numfmt-fix'

describe('centeredErrorStyle', () => {
  it('centers error literals under General alignment', () => {
    expect(centeredErrorStyle({ v: '#REF!', t: CellValueType.STRING }, { fs: 11 })).toEqual({
      fs: 11,
      ht: HorizontalAlign.CENTER,
    })
    expect(centeredErrorStyle({ v: '#N/A' }, null)).toEqual({ ht: HorizontalAlign.CENTER })
    expect(centeredErrorStyle({ v: '#DIV/0!' }, { ht: HorizontalAlign.UNSPECIFIED })).toEqual({
      ht: HorizontalAlign.CENTER,
    })
  })

  it('keeps an explicit alignment', () => {
    expect(centeredErrorStyle({ v: '#VALUE!' }, { ht: HorizontalAlign.LEFT })).toBeNull()
    expect(centeredErrorStyle({ v: '#VALUE!' }, { ht: HorizontalAlign.RIGHT })).toBeNull()
  })

  it('ignores text, numbers, rich text and hashed errors', () => {
    expect(centeredErrorStyle({ v: '#REF', t: CellValueType.STRING }, null)).toBeNull()
    expect(centeredErrorStyle({ v: 'text' }, null)).toBeNull()
    expect(centeredErrorStyle({ v: 12 }, null)).toBeNull()
    expect(centeredErrorStyle({ v: '####', t: CellValueType.NUMBER }, null)).toBeNull()
    expect(centeredErrorStyle({ v: '#REF!', p: { id: 'd' } as never }, null)).toBeNull()
  })
})

describe('nonTextDisplayLabel', () => {
  it('labels booleans and error literals, not numbers or text', () => {
    expect(nonTextDisplayLabel({ v: 1, t: CellValueType.BOOLEAN })).toBe('TRUE')
    expect(nonTextDisplayLabel({ v: false, t: CellValueType.BOOLEAN })).toBe('FALSE')
    expect(nonTextDisplayLabel({ v: '#VALUE!', t: CellValueType.STRING })).toBe('#VALUE!')
    expect(nonTextDisplayLabel({ v: '#N/A' })).toBe('#N/A')
    expect(nonTextDisplayLabel({ v: 'TRUE', t: CellValueType.STRING })).toBeNull()
    expect(nonTextDisplayLabel({ v: 1 })).toBeNull()
  })
})
