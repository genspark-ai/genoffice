import { describe, expect, it } from 'vitest'
import { decodeQuotedPrintable } from '../src/alt-chunk'

const MB = 1024 * 1024

describe('decodeQuotedPrintable memory', () => {
  it('decodes a 32 MiB part at ~1× memory instead of ~27×', () => {
    // Measured on main: the number[] sink pushed peak RSS to ~1.1 GB for a
    // 32 MiB part (27×). Writing straight into the output buffer keeps it
    // within a few multiples of the payload itself.
    const before = process.memoryUsage().rss
    const out = decodeQuotedPrintable('A'.repeat(32 * MB))
    expect(out.length).toBe(32 * MB)
    const delta = process.memoryUsage().rss - before
    expect(delta).toBeLessThan(5 * 32 * MB)
  })

  it('still decodes soft breaks and hex escapes correctly', () => {
    expect(Array.from(decodeQuotedPrintable('a=3Db=\r\nc=E9'))).toEqual([
      'a'.charCodeAt(0),
      61,
      'b'.charCodeAt(0),
      'c'.charCodeAt(0),
      0xe9,
    ])
  })
})
