import { describe, expect, it } from 'vitest'
import { asBytes } from '../src/renderer/as-bytes'

describe('asBytes', () => {
  it('passes through Uint8Array', () => {
    const raw = new Uint8Array([1, 2, 3])
    expect(asBytes(raw)).toBe(raw)
  })

  it('unwraps Node Buffer JSON', () => {
    expect(Array.from(asBytes({ type: 'Buffer', data: [9, 8] }))).toEqual([9, 8])
  })
})
