import * as CFB from 'cfb'
import { describe, expect, it } from 'vitest'

import { CompoundFileError, openCompoundStreams, type CompoundFileLimits } from '../src/index'

const LIMITS: CompoundFileLimits = { maxBytes: 1024 * 1024, maxStreamBytes: 1024 * 1024 }

function buildContainer(streams: Array<[string, Uint8Array]>): Uint8Array {
  const cfb = CFB.utils.cfb_new({ root: 'Root Entry' })
  for (const [name, bytes] of streams) CFB.utils.cfb_add(cfb, `/${name}`, Buffer.from(bytes))
  const out = CFB.write(cfb, { type: 'buffer' }) as Buffer
  return new Uint8Array(out.buffer, out.byteOffset, out.byteLength)
}

function codeOf(run: () => unknown): string {
  try {
    run()
  } catch (error) {
    expect(error).toBeInstanceOf(CompoundFileError)
    return (error as CompoundFileError).code
  }
  throw new Error('expected the opener to throw a CompoundFileError')
}

describe('openCompoundStreams', () => {
  it('returns owned stream bytes keyed by root-relative name', () => {
    const bytes = buildContainer([
      ['Workbook', Uint8Array.from([1, 2, 3])],
      ['Extra', Uint8Array.from([4, 5])],
    ])
    const opened = openCompoundStreams(bytes, LIMITS)
    expect(opened.names()).toEqual(['Extra', 'Workbook'])
    expect([...opened.stream('Workbook')!]).toEqual([1, 2, 3])
    expect(opened.stream('Missing')).toBeUndefined()
    // a copy, not a view of the container's mutable storage
    const parsed = CFB.read(bytes, { type: 'buffer' })
    const source = CFB.find(parsed, 'Workbook')!.content as number[]
    source[0] = 9
    expect(opened.stream('Workbook')![0]).toBe(1)
  })

  it('rejects an empty file and a non-CFB file as not_cfb', () => {
    expect(codeOf(() => openCompoundStreams(new Uint8Array(0), LIMITS))).toBe('not_cfb')
    expect(
      codeOf(() => openCompoundStreams(new TextEncoder().encode('not a compound file'), LIMITS)),
    ).toBe('not_cfb')
  })

  it('reports a truncated OLE2 header as truncated', () => {
    const bytes = new Uint8Array(200)
    bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
    expect(codeOf(() => openCompoundStreams(bytes, LIMITS))).toBe('truncated')
  })

  it('reports a damaged CFB body as corrupt', () => {
    const bytes = new Uint8Array(600).fill(0xff)
    bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
    expect(codeOf(() => openCompoundStreams(bytes, LIMITS))).toBe('corrupt')
  })

  it('caps the whole input and each stream', () => {
    const bytes = buildContainer([['Workbook', Uint8Array.from([1, 2, 3, 4])]])
    expect(codeOf(() => openCompoundStreams(bytes, { maxBytes: 600, maxStreamBytes: 1024 }))).toBe(
      'limits_exceeded',
    )
    expect(
      codeOf(() => openCompoundStreams(bytes, { maxBytes: LIMITS.maxBytes, maxStreamBytes: 3 })),
    ).toBe('limits_exceeded')
  })
})
