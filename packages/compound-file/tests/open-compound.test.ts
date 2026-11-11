import * as CFB from 'cfb'
import { describe, expect, it } from 'vitest'

import {
  OLE2_HEADER_BYTES,
  OLE2_MAGIC,
  hasOle2Signature,
  openOle2Container,
  type Ole2Limits,
} from '../src/index'

const LIMITS: Ole2Limits = { maxBytes: 1024 * 1024, maxStreamBytes: 1024 * 1024 }

/** wrap streams in an OLE2 compound file (the same shape cfb itself writes) */
function buildCompound(streams: Array<[string, Uint8Array]>): Uint8Array {
  const container = CFB.utils.cfb_new()
  for (const [name, content] of streams) CFB.utils.cfb_add(container, name, Array.from(content))
  const written = CFB.write(container, { type: 'buffer', fileType: 'cfb' }) as Uint8Array
  return new Uint8Array(written)
}

/** a buffer that carries the OLE2 signature but is not a compound file */
function brokenContainer(): Uint8Array {
  const bytes = new Uint8Array(OLE2_HEADER_BYTES + 64)
  bytes.set(OLE2_MAGIC, 0)
  return bytes
}

function expectFailure(result: ReturnType<typeof openOle2Container>) {
  expect(result.ok).toBe(false)
  if (result.ok) throw new Error('expected the opener to fail')
  return result
}

describe('the OLE2 signature', () => {
  it('accepts a compound file and rejects anything else', () => {
    expect(hasOle2Signature(buildCompound([['WordDocument', new Uint8Array([1, 2, 3])]]))).toBe(
      true,
    )
    expect(hasOle2Signature(new Uint8Array([0x25, 0x50, 0x44, 0x46]))).toBe(false)
    expect(hasOle2Signature(new Uint8Array(0))).toBe(false)
  })
})

describe('opening an OLE2 container', () => {
  it('lists the document streams and hands back their bytes', () => {
    const bytes = buildCompound([
      ['WordDocument', new Uint8Array([1, 2, 3])],
      ['1Table', new Uint8Array([4, 5])],
    ])
    const opened = openOle2Container(bytes, LIMITS)
    expect(opened.ok).toBe(true)
    if (!opened.ok) throw new Error('expected the opener to succeed')
    expect(opened.container.streams()).toEqual(['1Table', 'WordDocument'])
    expect([...(opened.container.stream('WordDocument') ?? [])]).toEqual([1, 2, 3])
    expect(opened.container.stream('Nothing')).toBeUndefined()
  })

  it('leaves cfb’s producer marker out of the stream list', () => {
    const opened = openOle2Container(buildCompound([['WordDocument', new Uint8Array([1])]]), LIMITS)
    if (!opened.ok) throw new Error('expected the opener to succeed')
    expect(opened.container.streams().some((name) => name.startsWith('\u0001'))).toBe(false)
  })

  it('hands back owned bytes: a stream is a Uint8Array copy, not cfb’s raw array', () => {
    const bytes = buildCompound([['WordDocument', new Uint8Array([1, 2, 3])]])
    const opened = openOle2Container(bytes, LIMITS)
    if (!opened.ok) throw new Error('expected the opener to succeed')
    const stream = opened.container.stream('WordDocument')
    // cfb hands back a plain array of byte values; the opener owns the bytes
    // and hands out a real typed array
    expect(stream).toBeInstanceOf(Uint8Array)
    expect(stream?.byteLength).toBe(3)
    expect([...(stream ?? [])]).toEqual([1, 2, 3])
    // and the copy is independent: a later change to the container cannot
    // alter the bytes a caller already holds
    bytes.fill(0)
    expect([...(opened.container.stream('WordDocument') ?? [])]).toEqual([1, 2, 3])
  })

  it('rejects an empty file as not_cfb', () => {
    expect(expectFailure(openOle2Container(new Uint8Array(0), LIMITS)).code).toBe('not_cfb')
  })

  it('rejects a file without the OLE2 signature as not_cfb', () => {
    expect(
      expectFailure(openOle2Container(new Uint8Array([0x25, 0x50, 0x44, 0x46]), LIMITS)).code,
    ).toBe('not_cfb')
  })

  it('rejects a file shorter than the 512-byte header as truncated', () => {
    const bytes = new Uint8Array(OLE2_HEADER_BYTES - 1)
    bytes.set(OLE2_MAGIC, 0)
    expect(expectFailure(openOle2Container(bytes, LIMITS)).code).toBe('truncated')
  })

  it('rejects a container cfb cannot read as corrupt', () => {
    expect(expectFailure(openOle2Container(brokenContainer(), LIMITS)).code).toBe('corrupt')
  })

  it('caps the input size', () => {
    const bytes = buildCompound([['WordDocument', new Uint8Array(16)]])
    expect(
      expectFailure(
        openOle2Container(bytes, { maxBytes: 16, maxStreamBytes: LIMITS.maxStreamBytes }),
      ).code,
    ).toBe('limits_exceeded')
  })

  it('caps a single stream size', () => {
    const bytes = buildCompound([['WordDocument', new Uint8Array(64)]])
    expect(
      expectFailure(openOle2Container(bytes, { maxBytes: LIMITS.maxBytes, maxStreamBytes: 8 }))
        .code,
    ).toBe('limits_exceeded')
  })
})
