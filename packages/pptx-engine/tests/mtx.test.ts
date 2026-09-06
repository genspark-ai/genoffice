import { describe, it, expect } from 'vitest'
import { AHuff, BitWriter, unpackLz, unpackMtx } from '../src/vendor/mtx/lzcomp'
import { ctfToSfnt, sanitizeCmap } from '../src/vendor/mtx/ctf'
import { mtxToSfnt } from '../src/vendor/mtx'

/** Encode bytes as an LZCOMP stream of literals (+ DUP2/DUP4/DUP6 where they apply). */
function encodeLz(data: Uint8Array, opts: { dups?: boolean; version?: number } = {}): Uint8Array {
  const w = new BitWriter()
  if ((opts.version ?? 2) !== 1) w.outputBit(false) // no run-length layer
  new AHuff(null, 8) // dist coder (unused: no copy items)
  new AHuff(null, 8) // len coder
  w.writeValue(data.length, 24)
  let numDistRanges = 1
  while (1 + (1 << (3 * numDistRanges)) - 1 < data.length) numDistRanges++
  const DUP2 = 256 + 8 * numDistRanges
  const sym = new AHuff(null, DUP2 + 3)
  for (let i = 0; i < data.length; i++) {
    const c = data[i]!
    if (opts.dups && i >= 2 && data[i - 2] === c) sym.writeSymbol(w, DUP2)
    else if (opts.dups && i >= 4 && data[i - 4] === c) sym.writeSymbol(w, DUP2 + 1)
    else if (opts.dups && i >= 6 && data[i - 6] === c) sym.writeSymbol(w, DUP2 + 2)
    else sym.writeSymbol(w, c)
  }
  return w.bytes()
}

describe('MTX LZCOMP decoder', () => {
  it('round-trips a literal stream through the adaptive Huffman model', () => {
    const src = new Uint8Array(3000)
    for (let i = 0; i < src.length; i++) src[i] = (i * 7919 + (i >> 3)) & 0xff
    expect(unpackLz(encodeLz(src), 2)).toEqual(src)
  })

  it('decodes DUP2/DUP4/DUP6 one-byte copy symbols', () => {
    const src = new Uint8Array([9, 1, 9, 1, 9, 1, 9, 1, 5, 6, 7, 8, 5, 6, 7, 8, 5, 6, 7, 8, 42])
    expect(unpackLz(encodeLz(src, { dups: true }), 2)).toEqual(src)
  })

  it('version 1 streams carry no run-length flag bit', () => {
    const src = new Uint8Array([1, 2, 3, 250, 251, 252])
    expect(unpackLz(encodeLz(src, { version: 1 }), 1)).toEqual(src)
  })

  it('splits the MTX header into three inflated blocks', () => {
    const blocks = [new Uint8Array([1, 2, 3]), new Uint8Array([4, 5]), new Uint8Array([6])].map(
      (b) => encodeLz(b),
    )
    const off1 = 10 + blocks[0]!.length
    const off2 = off1 + blocks[1]!.length
    const be24 = (v: number) => [(v >> 16) & 0xff, (v >> 8) & 0xff, v & 0xff]
    const mtx = new Uint8Array([
      2,
      ...be24(0),
      ...be24(off1),
      ...be24(off2),
      ...blocks[0]!,
      ...blocks[1]!,
      ...blocks[2]!,
    ])
    const [a, b, c] = unpackMtx(mtx)
    expect([...a]).toEqual([1, 2, 3])
    expect([...b]).toEqual([4, 5])
    expect([...c]).toEqual([6])
  })
})

/** Big-endian byte builder for the synthetic CTF container. */
class BE {
  bytes: number[] = []
  u8(v: number) {
    this.bytes.push(v & 0xff)
    return this
  }
  u16(v: number) {
    return this.u8(v >> 8).u8(v)
  }
  u32(v: number) {
    return this.u16(v >>> 16).u16(v & 0xffff)
  }
  tag(t: string) {
    for (const ch of t) this.u8(ch.charCodeAt(0))
    return this
  }
  raw(a: ArrayLike<number>) {
    for (let i = 0; i < a.length; i++) this.bytes.push(a[i]!)
    return this
  }
}

function synthCtf(opts: { glyphCopies?: number } = {}): [Uint8Array, Uint8Array, Uint8Array] {
  // head: 54 bytes, indexToLocFormat (offset 50) = 0 → short loca
  const head = new Uint8Array(54)
  new DataView(head.buffer).setUint32(0, 0x00010000)
  new DataView(head.buffer).setUint32(12, 0x5f0f3cf5)
  new DataView(head.buffer).setUint16(18, 1000)
  // maxp 1.0, numGlyphs = 2
  const copies = opts.glyphCopies ?? 1
  const maxp = new BE()
    .u32(0x00010000)
    .u16(1 + copies)
    .u16(4)
    .u16(1)
    .u16(0)
    .u16(0)
    .u16(2)
    .u16(0)
    .u16(0)
    .u16(0)
    .u16(0)
    .u16(0)
    .u16(0)
    .u16(0)
    .u16(0).bytes
  const hmtxB = new BE().u16(500).u16(0)
  for (let i = 0; i < copies; i++) hmtxB.u16(600).u16(0)
  const hmtx = hmtxB.bytes
  // cvt in CTF delta form: 3 entries → 10, 10+200 (=210 via 238*(248-247)... use 238 code for raw), 210-5
  const cvt = new BE().u16(3).u8(10).u8(238).u16(200).u8(239).u8(5).bytes
  // glyf: glyph 0 empty; glyph 1 = 100×100 square with 4 on-curve points, no instructions
  const square = new BE()
    .u16(1) // one contour, bbox computed
    .u8(3) // 255UShort points-in-contour (first contour: total = 1 + 3)
    .u8(0)
    .u8(11)
    .u8(1)
    .u8(10) // triplet flags: (0,0) (+x) (+y) (−x), all on-curve
    .u8(0)
    .u8(100)
    .u8(100)
    .u8(100) // triplet payload bytes
    .u8(0) // pushCount
    .u8(0).bytes // codeSize
  const glyfB = new BE().u16(0) // glyph 0: numContours 0
  for (let i = 0; i < copies; i++) glyfB.raw(square)
  const glyf = glyfB.bytes
  const tables: Array<[string, number[]]> = [
    ['cvt ', cvt],
    ['glyf', glyf],
    ['head', [...head]],
    ['hmtx', hmtx],
    ['maxp', maxp],
  ]
  const dirEnd = 12 + 16 * tables.length
  const s0 = new BE().u32(0x00010000).u16(tables.length).u16(64).u16(2).u16(16)
  let off = dirEnd
  for (const [tag, data] of tables) {
    s0.tag(tag).u32(0).u32(off).u32(data.length)
    off += data.length
  }
  for (const [, data] of tables) s0.raw(data)
  return [new Uint8Array(s0.bytes), new Uint8Array(0), new Uint8Array(0)]
}

describe('CTF → sfnt', () => {
  it('rebuilds glyf/loca/cvt and writes a checksummed table directory', () => {
    const sfnt = ctfToSfnt(synthCtf())
    const dv = new DataView(sfnt.buffer, sfnt.byteOffset, sfnt.byteLength)
    expect(dv.getUint32(0)).toBe(0x00010000)
    const numTables = dv.getUint16(4)
    const dir = new Map<string, { off: number; len: number }>()
    for (let i = 0; i < numTables; i++) {
      const o = 12 + 16 * i
      const tag = String.fromCharCode(sfnt[o]!, sfnt[o + 1]!, sfnt[o + 2]!, sfnt[o + 3]!)
      dir.set(tag, { off: dv.getUint32(o + 8), len: dv.getUint32(o + 12) })
    }
    // synthesized loca must land in tag order, not last
    expect([...dir.keys()]).toEqual(['cvt ', 'glyf', 'head', 'hmtx', 'loca', 'maxp'])
    const loca = dir.get('loca')!
    expect([...sfnt.subarray(loca.off, loca.off + loca.len)]).toEqual([0, 0, 0, 0, 0, 12])
    const glyf = dir.get('glyf')!
    const g = sfnt.subarray(glyf.off, glyf.off + glyf.len)
    const gv = new DataView(g.buffer, g.byteOffset, g.byteLength)
    expect(gv.getInt16(0)).toBe(1) // numContours
    expect([gv.getInt16(2), gv.getInt16(4), gv.getInt16(6), gv.getInt16(8)]).toEqual([
      0, 0, 100, 100,
    ])
    expect(gv.getUint16(10)).toBe(3) // endPtsOfContours
    expect(gv.getUint16(12)).toBe(0) // instructionLength
    expect([...g.subarray(14, 18)]).toEqual([0x37, 0x33, 0x35, 0x23]) // flags
    expect([...g.subarray(18, 23)]).toEqual([0, 100, 100, 0, 100]) // x: 0,100,100 (neg via flag); y: 0,100
    const cvt = dir.get('cvt ')!
    const cv = new DataView(sfnt.buffer, sfnt.byteOffset + cvt.off, cvt.len)
    expect([cv.getInt16(0), cv.getInt16(2), cv.getInt16(4)]).toEqual([10, 210, 205])
    // whole-font checksum lands on 0xB1B0AFBA
    let sum = 0
    for (let i = 0; i < sfnt.length; i += 4) sum = (sum + dv.getUint32(i)) >>> 0
    expect(sum).toBe(0xb1b0afba)
  })

  it('promotes a short loca to long when the rebuilt glyf outgrows 128 KB', () => {
    const streams = synthCtf({ glyphCopies: 6000 })
    const sfnt = ctfToSfnt(streams)
    const dv = new DataView(sfnt.buffer, sfnt.byteOffset, sfnt.byteLength)
    const n = dv.getUint16(4)
    const dir = new Map<string, { off: number; len: number }>()
    for (let i = 0; i < n; i++) {
      const o = 12 + 16 * i
      dir.set(String.fromCharCode(sfnt[o]!, sfnt[o + 1]!, sfnt[o + 2]!, sfnt[o + 3]!), {
        off: dv.getUint32(o + 8),
        len: dv.getUint32(o + 12),
      })
    }
    expect(dir.get('glyf')!.len).toBeGreaterThan(0x1fffe)
    expect(dir.get('loca')!.len).toBe(4 * 6002)
    expect(dv.getInt16(dir.get('head')!.off + 50)).toBe(1)
    expect(dv.getUint32(dir.get('loca')!.off + 4 * 6001)).toBe(dir.get('glyf')!.len)
  })

  it('mtxToSfnt returns null for garbage instead of throwing', () => {
    expect(mtxToSfnt(new Uint8Array([1, 2, 3]))).toBeNull()
    expect(mtxToSfnt(new Uint8Array(64))).toBeNull()
  })
})

describe('cmap sanitizing', () => {
  // minimal one-segment format 4 (24 bytes); language field at +4
  const fmt4 = (lang: number) =>
    new BE()
      .u16(4)
      .u16(24)
      .u16(lang)
      .u16(2)
      .u16(2)
      .u16(0)
      .u16(0)
      .u16(0xffff)
      .u16(0)
      .u16(1)
      .u16(0)
      .u16(0).bytes
  const fmt14 = new BE().u16(14).u16(0).u32(0x7fffffff).bytes // length runs past the table

  it('drops a record whose subtable runs past the table and keeps the rest byte-identical', () => {
    // header(4) + 3 records(24) = 28; fmt4 at 28, fmt14 at 52, (3,10) shares the fmt4 body
    const table = new BE()
      .u16(0)
      .u16(3)
      .u16(0)
      .u16(5)
      .u32(52)
      .u16(3)
      .u16(1)
      .u32(28)
      .u16(3)
      .u16(10)
      .u32(28)
      .raw(fmt4(0))
      .raw(fmt14).bytes
    const out = sanitizeCmap(new Uint8Array(table))
    const dv = new DataView(out.buffer, out.byteOffset, out.byteLength)
    expect(dv.getUint16(2)).toBe(2)
    expect([dv.getUint16(4), dv.getUint16(6), dv.getUint32(8)]).toEqual([3, 1, 20])
    expect([dv.getUint16(12), dv.getUint16(14), dv.getUint32(16)]).toEqual([3, 10, 20])
    expect([...out.subarray(20)]).toEqual(fmt4(0))
  })

  it('zeroes a non-zero language on Windows-platform subtables (prod_070: language=3 → rejected)', () => {
    const table = new BE().u16(0).u16(1).u16(3).u16(1).u32(12).raw(fmt4(3)).bytes
    const out = sanitizeCmap(new Uint8Array(table))
    expect([...out.subarray(12)]).toEqual(fmt4(0))
    expect(out.length).toBe(table.length)
  })

  it('keeps a well-formed format-14 record (length lives at +2)', () => {
    const fmt14ok = new BE().u16(14).u32(10).u32(0).bytes // 10 bytes, zero variation selectors
    const tbl = new BE()
      .u16(0)
      .u16(2)
      .u16(0)
      .u16(5)
      .u32(20 + 24)
      .u16(3)
      .u16(1)
      .u32(20)
      .raw(fmt4(0))
      .raw(fmt14ok).bytes
    const clean = new Uint8Array(tbl)
    expect(sanitizeCmap(clean)).toBe(clean)
  })

  it('returns a clean table untouched', () => {
    const clean = new Uint8Array(new BE().u16(0).u16(1).u16(3).u16(1).u32(12).raw(fmt4(0)).bytes)
    expect(sanitizeCmap(clean)).toBe(clean)
  })
})
