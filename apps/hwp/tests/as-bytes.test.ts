import { describe, expect, it } from 'vitest'
import { asBytes, rawFileBytes } from '../src/shared/as-bytes'

describe('asBytes', () => {
  it('passes through Uint8Array', () => {
    const raw = new Uint8Array([1, 2, 3])
    expect(asBytes(raw)).toBe(raw)
  })

  it('unwraps Node Buffer JSON', () => {
    expect(Array.from(asBytes({ type: 'Buffer', data: [9, 8] }))).toEqual([9, 8])
  })

  it('wraps ArrayBuffer and views', () => {
    expect(Array.from(asBytes(new Uint8Array([4, 5]).buffer))).toEqual([4, 5])
    expect(Array.from(asBytes(new DataView(new Uint8Array([6]).buffer)))).toEqual([6])
  })

  it('rejects junk', () => {
    expect(rawFileBytes('nope')).toBeNull()
    expect(() => asBytes(null)).toThrow('hwp: expected file bytes')
  })
})
