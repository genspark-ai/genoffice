import { describe, expect, it } from 'vitest'

import { readFileBytes } from '../src/renderer/file-bytes'

/** A fake chunk api serving a known byte pattern (i mod 251) with the size
 * riding along, exactly like the main-process handler */
function fakeApi(content: Uint8Array, calls: Array<{ start: number; end: number }> = []) {
  return {
    readFileChunk: async (_path: string, start: number, end: number) => {
      calls.push({ start, end })
      const clampedEnd = Math.min(end, content.length)
      return {
        data: content.slice(start, clampedEnd),
        fileSize: content.length,
      }
    },
  }
}

describe('readFileBytes', () => {
  it('assembles a multi-chunk file through repeated slice invokes', async () => {
    const content = new Uint8Array(2560 * 1024)
    for (let i = 0; i < content.length; i++) content[i] = i % 251
    const calls: Array<{ start: number; end: number }> = []
    const out = await readFileBytes(fakeApi(content, calls), '/docs/book.pdf', 1024 * 1024)
    // deep-equal on a multi-megabyte typed array is pathologically slow in
    // vitest's matcher, so compare length + a rolling checksum
    const checksum = (a: Uint8Array) => {
      let h = 0
      for (let i = 0; i < a.length; i++) h = (h * 31 + a[i]) | 0
      return h
    }
    expect(out.length).toBe(content.length)
    expect(checksum(out)).toBe(checksum(content))
    // three 1 MiB slices, boundaries exactly at the chunk size
    expect(calls.map((c) => c.end - c.start)).toEqual([1024 * 1024, 1024 * 1024, 512 * 1024])
  })

  it('an exactly chunk-sized file needs a single invoke', async () => {
    const content = new Uint8Array(1024 * 1024).fill(7)
    const calls: Array<{ start: number; end: number }> = []
    const out = await readFileBytes(fakeApi(content, calls), '/docs/book.pdf', 1024 * 1024)
    expect(out).toEqual(content)
    expect(calls).toEqual([{ start: 0, end: 1024 * 1024 }])
  })

  it('an empty file resolves to empty bytes without looping forever', async () => {
    const out = await readFileBytes(fakeApi(new Uint8Array(0)), '/docs/empty.pdf')
    expect(out).toHaveLength(0)
  })

  it('a file smaller than two chunks terminates at the reported size', async () => {
    const content = new Uint8Array(6).fill(5)
    const calls: Array<{ start: number; end: number }> = []
    const out = await readFileBytes(fakeApi(content, calls), '/docs/book.pdf', 4)
    expect(out).toEqual(content)
    expect(calls).toEqual([
      { start: 0, end: 4 },
      { start: 4, end: 6 },
    ])
  })
})
