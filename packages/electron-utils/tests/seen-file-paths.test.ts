import { describe, expect, it } from 'vitest'
import { seenFilePaths } from '../src/drop-open'

describe('seenFilePaths (grant-ledger feed)', () => {
  it('keeps strings, trims and dedupes', () => {
    expect(seenFilePaths(['/tmp/a.png', ' /tmp/a.png ', '/tmp/b.md'])).toEqual([
      '/tmp/a.png',
      '/tmp/b.md',
    ])
  })

  it('drops junk without dying', () => {
    expect(seenFilePaths([42, null, '', 'x'.repeat(5000), '/tmp/ok.txt'])).toEqual(['/tmp/ok.txt'])
    expect(seenFilePaths('not an array')).toEqual([])
    expect(seenFilePaths(undefined)).toEqual([])
  })

  it('caps the payload like the resolver does', () => {
    const many = Array.from({ length: 300 }, (_, i) => `/tmp/f${i}.txt`)
    expect(seenFilePaths(many).length).toBe(100)
  })
})
