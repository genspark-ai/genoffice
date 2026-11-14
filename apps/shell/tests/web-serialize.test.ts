import { describe, expect, it } from 'vitest'
import { decodeFrame, encodeFrame } from '../src/web/serialize'

describe('web bridge serialization', () => {
  it('round-trips plain values untouched', () => {
    const value = { a: 1, b: ['x', { c: null }], d: 'ünï' }
    expect(decodeFrame(encodeFrame(value))).toEqual(value)
  })

  it('delivers ArrayBuffer inputs as Uint8Array like IPC does', () => {
    const bytes = new Uint8Array([1, 2, 3, 250]).buffer
    const out = decodeFrame(encodeFrame({ data: bytes })) as { data: Uint8Array }
    expect(out.data).toBeInstanceOf(Uint8Array)
    expect([...out.data]).toEqual([1, 2, 3, 250])
  })

  it('arrives as Uint8Array like structured clone hands a Buffer over', () => {
    // main-process handlers return Node Buffers; the renderer sees Uint8Array
    const fakeBufferLike = {
      __genofficeBytes: Buffer.from([9, 0, 255]).toString('base64'),
      __genofficeBytesView: true,
    }
    const out = decodeFrame(JSON.stringify(fakeBufferLike)) as Uint8Array
    expect(out).toBeInstanceOf(Uint8Array)
    expect([...out]).toEqual([9, 0, 255])
  })

  it('encodes typed-array views with their live byte range', () => {
    const backing = new Uint8Array([10, 20, 30, 40, 50])
    const view = backing.subarray(1, 4)
    const out = decodeFrame(encodeFrame({ view })) as { view: Uint8Array }
    expect([...out.view]).toEqual([20, 30, 40])
  })

  it('survives large files without corruption', () => {
    const big = new Uint8Array(512 * 1024)
    for (let i = 0; i < big.length; i += 4096) big[i] = i % 251
    const out = decodeFrame(encodeFrame({ big: big.buffer })) as { big: Uint8Array }
    expect([...out.big]).toEqual([...big])
  })

  it('encodes a Buffer return into a Uint8Array on the other side', () => {
    const out = decodeFrame(encodeFrame({ b: Buffer.from([7, 8, 9]) })) as { b: Uint8Array }
    expect(out.b).toBeInstanceOf(Uint8Array)
    expect([...out.b]).toEqual([7, 8, 9])
  })
})
