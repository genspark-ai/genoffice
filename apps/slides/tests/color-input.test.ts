import { describe, expect, it } from 'vitest'
import { armColorInput, toPickerHex } from '../src/renderer/color-input'

function fakeInput(value: string): HTMLInputElement {
  return { value } as HTMLInputElement
}

describe('armColorInput', () => {
  it('flips the lowest blue bit so re-picking the shown color still fires change', () => {
    const input = fakeInput('#000000')
    armColorInput(input)
    expect(input.value).toBe('#000001')
    expect(input.value).not.toBe('#000000')
  })

  it('stays within range at channel max', () => {
    const input = fakeInput('#ffffff')
    armColorInput(input)
    expect(input.value).toBe('#fffffe')
  })

  it('round-trips: arming twice restores the original value', () => {
    const input = fakeInput('#4472c4')
    armColorInput(input)
    const nudged = input.value
    expect(nudged).not.toBe('#4472c4')
    armColorInput(input)
    expect(input.value).toBe('#4472c4')
  })

  it('ignores values that are not 6-digit hex', () => {
    const input = fakeInput('')
    armColorInput(input)
    expect(input.value).toBe('')
  })
})

describe('toPickerHex', () => {
  it('lowercases and keeps the # prefix', () => {
    expect(toPickerHex('#FF8800')).toBe('#ff8800')
    expect(toPickerHex('4472C4')).toBe('#4472c4')
  })

  it('drops the alpha byte of #RRGGBBAA', () => {
    expect(toPickerHex('#FF880080')).toBe('#ff8800')
  })

  it('returns null for missing or non-hex colors (gradient/image backgrounds)', () => {
    expect(toPickerHex(undefined)).toBeNull()
    expect(toPickerHex('')).toBeNull()
    expect(toPickerHex('red')).toBeNull()
    expect(toPickerHex('#fff')).toBeNull()
  })
})
