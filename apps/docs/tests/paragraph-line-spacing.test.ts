import { describe, expect, it } from 'vitest'
import {
  DEFAULT_MULTIPLE,
  lineSpacingAttrs,
  lineSpacingFromAttrs,
  pickLineSpacing,
} from '../src/renderer/components/paragraph-line-spacing'

describe('lineSpacingFromAttrs', () => {
  it('preselects Single / 1.5 lines / Double for exact multiples', () => {
    expect(lineSpacingFromAttrs({}).choice).toBe('single')
    expect(lineSpacingFromAttrs({ lineSpacing: 1 }).choice).toBe('single')
    expect(lineSpacingFromAttrs({ lineSpacing: 1.5 }).choice).toBe('oneHalf')
    expect(lineSpacingFromAttrs({ lineSpacing: 2 }).choice).toBe('double')
  })

  it('loads any other multiple as Multiple with the value shown', () => {
    expect(lineSpacingFromAttrs({ lineSpacing: 1.15 })).toMatchObject({
      choice: 'multiple',
      multiple: 1.15,
    })
    expect(lineSpacingFromAttrs({ lineSpacing: 3 })).toMatchObject({
      choice: 'multiple',
      multiple: 3,
    })
    expect(lineSpacingFromAttrs({ lineRule: 'auto', lineRawTwips: 276 })).toMatchObject({
      choice: 'multiple',
      multiple: 1.15,
    })
  })

  it('keeps At least / Exactly with their pt value', () => {
    expect(lineSpacingFromAttrs({ lineRule: 'exact', lineRawTwips: 360 })).toEqual({
      choice: 'exact',
      multiple: 1,
      pt: 18,
    })
    expect(lineSpacingFromAttrs({ lineRule: 'atLeast', lineRawTwips: 250 }).choice).toBe('atLeast')
  })
})

describe('lineSpacingAttrs', () => {
  it('writes the same attrs as the old presets', () => {
    const base = { multiple: 1, pt: 12 }
    expect(lineSpacingAttrs({ ...base, choice: 'single' })).toEqual({
      lineSpacing: null,
      lineRule: null,
      lineRawTwips: null,
    })
    expect(lineSpacingAttrs({ ...base, choice: 'oneHalf' }).lineSpacing).toBe(1.5)
    expect(lineSpacingAttrs({ ...base, choice: 'double' }).lineSpacing).toBe(2)
    expect(lineSpacingAttrs({ ...base, choice: 'multiple', multiple: 2.5 }).lineSpacing).toBe(2.5)
    expect(lineSpacingAttrs({ ...base, choice: 'multiple', multiple: 1 }).lineSpacing).toBeNull()
    expect(lineSpacingAttrs({ ...base, choice: 'exact', pt: 18 })).toEqual({
      lineSpacing: null,
      lineRule: 'exact',
      lineRawTwips: 360,
    })
    expect(lineSpacingAttrs({ ...base, choice: 'atLeast', pt: 0.5 }).lineRawTwips).toBe(20)
  })

  it('round-trips a loaded multiple unchanged', () => {
    const attrs = { lineSpacing: 1.15, lineRule: null, lineRawTwips: null }
    expect(lineSpacingAttrs(lineSpacingFromAttrs(attrs))).toEqual(attrs)
  })
})

describe('pickLineSpacing', () => {
  it('opens Multiple at 3 lines from a fixed choice and keeps a custom value', () => {
    const single = lineSpacingFromAttrs({})
    expect(pickLineSpacing(single, 'multiple').multiple).toBe(DEFAULT_MULTIPLE)
    const custom = lineSpacingFromAttrs({ lineSpacing: 1.15 })
    expect(pickLineSpacing(pickLineSpacing(custom, 'double'), 'multiple').multiple).toBe(1.15)
  })
})
