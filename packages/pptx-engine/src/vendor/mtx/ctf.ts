/*
 * Compact Table Format → sfnt — TypeScript port of libeot's src/ctf (parseCTF.c,
 * parseTTF.c, SFNTContainer.c) and triplet_encodings.c. libeot is MPL-2.0 (see LICENSE).
 * Spec: http://www.w3.org/Submission/MTX/
 */
import { MtxError } from './lzcomp'

/** Big-endian reader with libeot's Stream semantics (byte position + bit position). */
class Reader {
  pos = 0
  bitPos = 0
  constructor(readonly buf: Uint8Array) {}
  get size(): number {
    return this.buf.length
  }
  private need(n: number): void {
    if (this.bitPos !== 0) throw new MtxError('read off byte boundary')
    if (this.pos + n > this.buf.length) throw new MtxError('CTF stream truncated')
  }
  u8(): number {
    this.need(1)
    return this.buf[this.pos++]!
  }
  peekU8(): number {
    this.need(1)
    return this.buf[this.pos]!
  }
  u16(): number {
    this.need(2)
    const v = (this.buf[this.pos]! << 8) | this.buf[this.pos + 1]!
    this.pos += 2
    return v
  }
  s16(): number {
    return (this.u16() << 16) >> 16
  }
  u32(): number {
    this.need(4)
    const b = this.buf
    const v =
      ((b[this.pos]! << 24) | (b[this.pos + 1]! << 16) | (b[this.pos + 2]! << 8) | b[this.pos + 3]!) >>>
      0
    this.pos += 4
    return v
  }
  seek(pos: number): void {
    if (this.bitPos !== 0) throw new MtxError('seek off byte boundary')
    if (pos > this.buf.length) throw new MtxError('seek past end')
    this.pos = pos
  }
  skip(n: number): void {
    this.seek(this.pos + n)
  }
  bytes(n: number): Uint8Array {
    this.need(n)
    const out = this.buf.subarray(this.pos, this.pos + n)
    this.pos += n
    return out
  }
  readNBits(n: number): number {
    let out = 0
    for (let i = 0; i < n; i++) {
      if (this.pos >= this.buf.length) throw new MtxError('bit read past end')
      const bit = (this.buf[this.pos]! >> (7 - this.bitPos)) & 1
      out |= bit << (n - i - 1)
      if (++this.bitPos === 8) {
        this.bitPos = 0
        this.pos++
      }
    }
    return out >>> 0
  }
}

/** Growable big-endian writer; `pos` may be moved back to patch earlier fields. */
class Writer {
  private buf: Uint8Array
  pos = 0
  size = 0
  constructor(capacity = 1024) {
    this.buf = new Uint8Array(Math.max(16, capacity))
  }
  private ensure(n: number): void {
    const need = this.pos + n
    if (need > this.buf.length) {
      const next = new Uint8Array(Math.max(need, this.buf.length * 2))
      next.set(this.buf)
      this.buf = next
    }
  }
  private fix(): void {
    if (this.pos > this.size) this.size = this.pos
  }
  u8(v: number): void {
    this.ensure(1)
    this.buf[this.pos++] = v & 0xff
    this.fix()
  }
  u16(v: number): void {
    this.ensure(2)
    this.buf[this.pos++] = (v >> 8) & 0xff
    this.buf[this.pos++] = v & 0xff
    this.fix()
  }
  s16(v: number): void {
    this.u16(v & 0xffff)
  }
  u32(v: number): void {
    this.ensure(4)
    this.buf[this.pos++] = (v >>> 24) & 0xff
    this.buf[this.pos++] = (v >>> 16) & 0xff
    this.buf[this.pos++] = (v >>> 8) & 0xff
    this.buf[this.pos++] = v & 0xff
    this.fix()
  }
  bytes(b: Uint8Array): void {
    this.ensure(b.length)
    this.buf.set(b, this.pos)
    this.pos += b.length
    this.fix()
  }
  /** Reserve n bytes (zeroed) to be patched later. */
  skip(n: number): void {
    this.ensure(n)
    this.pos += n
    this.fix()
  }
  result(): Uint8Array {
    return this.buf.slice(0, this.size)
  }
}

interface Table {
  tag: string
  offset: number
  length: number
  data: Uint8Array | null
}

// [byteCount, xBits, yBits, deltaX, deltaY, xSign, ySign] per MTX §TripletEncoding
// prettier-ignore
const TRIPLET: ReadonlyArray<readonly [number, number, number, number, number, number, number]> = [[2,0,8,0,0,0,-1],[2,0,8,0,0,0,1],[2,0,8,0,256,0,-1],[2,0,8,0,256,0,1],[2,0,8,0,512,0,-1],[2,0,8,0,512,0,1],[2,0,8,0,768,0,-1],[2,0,8,0,768,0,1],[2,0,8,0,1024,0,-1],[2,0,8,0,1024,0,1],[2,8,0,0,0,-1,0],[2,8,0,0,0,1,0],[2,8,0,256,0,-1,0],[2,8,0,256,0,1,0],[2,8,0,512,0,-1,0],[2,8,0,512,0,1,0],[2,8,0,768,0,-1,0],[2,8,0,768,0,1,0],[2,8,0,1024,0,-1,0],[2,8,0,1024,0,1,0],[2,4,4,1,1,-1,-1],[2,4,4,1,1,1,-1],[2,4,4,1,1,-1,1],[2,4,4,1,1,1,1],[2,4,4,1,17,-1,-1],[2,4,4,1,17,1,-1],[2,4,4,1,17,-1,1],[2,4,4,1,17,1,1],[2,4,4,1,33,-1,-1],[2,4,4,1,33,1,-1],[2,4,4,1,33,-1,1],[2,4,4,1,33,1,1],[2,4,4,1,49,-1,-1],[2,4,4,1,49,1,-1],[2,4,4,1,49,-1,1],[2,4,4,1,49,1,1],[2,4,4,17,1,-1,-1],[2,4,4,17,1,1,-1],[2,4,4,17,1,-1,1],[2,4,4,17,1,1,1],[2,4,4,17,17,-1,-1],[2,4,4,17,17,1,-1],[2,4,4,17,17,-1,1],[2,4,4,17,17,1,1],[2,4,4,17,33,-1,-1],[2,4,4,17,33,1,-1],[2,4,4,17,33,-1,1],[2,4,4,17,33,1,1],[2,4,4,17,49,-1,-1],[2,4,4,17,49,1,-1],[2,4,4,17,49,-1,1],[2,4,4,17,49,1,1],[2,4,4,33,1,-1,-1],[2,4,4,33,1,1,-1],[2,4,4,33,1,-1,1],[2,4,4,33,1,1,1],[2,4,4,33,17,-1,-1],[2,4,4,33,17,1,-1],[2,4,4,33,17,-1,1],[2,4,4,33,17,1,1],[2,4,4,33,33,-1,-1],[2,4,4,33,33,1,-1],[2,4,4,33,33,-1,1],[2,4,4,33,33,1,1],[2,4,4,33,49,-1,-1],[2,4,4,33,49,1,-1],[2,4,4,33,49,-1,1],[2,4,4,33,49,1,1],[2,4,4,49,1,-1,-1],[2,4,4,49,1,1,-1],[2,4,4,49,1,-1,1],[2,4,4,49,1,1,1],[2,4,4,49,17,-1,-1],[2,4,4,49,17,1,-1],[2,4,4,49,17,-1,1],[2,4,4,49,17,1,1],[2,4,4,49,33,-1,-1],[2,4,4,49,33,1,-1],[2,4,4,49,33,-1,1],[2,4,4,49,33,1,1],[2,4,4,49,49,-1,-1],[2,4,4,49,49,1,-1],[2,4,4,49,49,-1,1],[2,4,4,49,49,1,1],[3,8,8,1,1,-1,-1],[3,8,8,1,1,1,-1],[3,8,8,1,1,-1,1],[3,8,8,1,1,1,1],[3,8,8,1,257,-1,-1],[3,8,8,1,257,1,-1],[3,8,8,1,257,-1,1],[3,8,8,1,257,1,1],[3,8,8,1,513,-1,-1],[3,8,8,1,513,1,-1],[3,8,8,1,513,-1,1],[3,8,8,1,513,1,1],[3,8,8,257,1,-1,-1],[3,8,8,257,1,1,-1],[3,8,8,257,1,-1,1],[3,8,8,257,1,1,1],[3,8,8,257,257,-1,-1],[3,8,8,257,257,1,-1],[3,8,8,257,257,-1,1],[3,8,8,257,257,1,1],[3,8,8,257,513,-1,-1],[3,8,8,257,513,1,-1],[3,8,8,257,513,-1,1],[3,8,8,257,513,1,1],[3,8,8,513,1,-1,-1],[3,8,8,513,1,1,-1],[3,8,8,513,1,-1,1],[3,8,8,513,1,1,1],[3,8,8,513,257,-1,-1],[3,8,8,513,257,1,-1],[3,8,8,513,257,-1,1],[3,8,8,513,257,1,1],[3,8,8,513,513,-1,-1],[3,8,8,513,513,1,-1],[3,8,8,513,513,-1,1],[3,8,8,513,513,1,1],[4,12,12,0,0,-1,-1],[4,12,12,0,0,1,-1],[4,12,12,0,0,-1,1],[4,12,12,0,0,1,1],[5,16,16,0,0,-1,-1],[5,16,16,0,0,1,-1],[5,16,16,0,0,-1,1],[5,16,16,0,0,1,1]]

const i16 = (v: number): number => (v << 16) >> 16

function read255UShort(r: Reader): number {
  const code = r.u8()
  switch (code) {
    case 253:
      return r.u16()
    case 255:
      return 253 + r.u8()
    case 254:
      return 506 + r.u8()
    default:
      return code
  }
}

function read255Short(r: Reader): number {
  let code = r.u8()
  if (code === 253) return r.s16()
  let sign = 1
  if (code === 250) {
    sign = -1
    code = r.u8()
  }
  let out: number
  switch (code) {
    case 255:
      out = 250 + r.u8()
      break
    case 254:
      out = 500 + r.u8()
      break
    default:
      out = code
  }
  return out * sign
}

/** Push-instruction stream (§HopCodes): values become PUSHB/PUSHW/NPUSHB/NPUSHW runs. */
function decodePushInstructions(sIn: Reader, out: Writer, pushCount: number): void {
  const NPUSHB = 0x40
  const NPUSHW = 0x41
  const PUSHB = 0xb0
  const PUSHW = 0xb8
  const data: number[] = []
  let lastIsByte = true
  let runCount = 0
  const dump = () => {
    if (runCount === 0) return
    if (runCount < 8) out.u8((lastIsByte ? PUSHB : PUSHW) | (runCount - 1))
    else {
      out.u8(lastIsByte ? NPUSHB : NPUSHW)
      out.u8(runCount)
    }
    for (let i = data.length - runCount; i < data.length; i++) {
      if (lastIsByte) out.u8(data[i]!)
      else out.s16(data[i]!)
    }
  }
  const put = (value: number) => {
    const isByte = value >= 0 && value < 256
    if (isByte !== lastIsByte || runCount === 255) {
      dump()
      lastIsByte = isByte
      runCount = 0
    }
    data.push(value)
    runCount++
  }
  let remaining = pushCount
  while (remaining > 0) {
    const code = sIn.peekU8()
    if (code === 0xfb) {
      // A B 0xFB C -> A B A C A
      if (remaining < 3 || data.length < 2) throw new MtxError('corrupt hop code')
      remaining -= 3
      const prev = data[data.length - 2]!
      sIn.u8()
      put(prev)
      put(read255Short(sIn))
      put(prev)
    } else if (code === 0xfc) {
      // A B 0xFC C D -> A B A C A D A
      if (remaining < 5 || data.length < 2) throw new MtxError('corrupt hop code')
      remaining -= 5
      const prev = data[data.length - 2]!
      sIn.u8()
      put(prev)
      put(read255Short(sIn))
      put(prev)
      put(read255Short(sIn))
      put(prev)
    } else {
      put(read255Short(sIn))
      remaining--
    }
  }
  dump()
}

function makeFlags(x: number, y: number, onCurve: boolean, first: boolean): number {
  let f = onCurve ? 0x01 : 0
  if (!first && x === 0) f |= 0x10
  else if (x > -256 && x < 0) f |= 0x02
  else if (x >= 0 && x < 256) f |= 0x02 | 0x10
  if (!first && y === 0) f |= 0x20
  else if (y > -256 && y < 0) f |= 0x04
  else if (y >= 0 && y < 256) f |= 0x04 | 0x20
  return f
}

function decodeSimpleGlyph(
  numContours: number,
  streams: [Reader, Reader, Reader],
  out: Writer,
  bbox: [number, number, number, number] | null,
): void {
  if (numContours === 0) return
  const sIn = streams[0]
  out.s16(numContours)
  const bboxAt = out.pos
  if (bbox) {
    for (const v of bbox) out.s16(v)
  } else {
    out.skip(8)
  }
  let totalPoints = 0
  for (let i = 0; i < numContours; i++) {
    if (i === 0) totalPoints = 1
    totalPoints += read255UShort(sIn)
    out.s16(totalPoints - 1)
  }
  const flags = new Uint8Array(totalPoints)
  const xs = new Int16Array(totalPoints)
  const ys = new Int16Array(totalPoints)
  for (let i = 0; i < totalPoints; i++) flags[i] = sIn.u8()
  let minX = 32767
  let minY = 32767
  let maxX = -32768
  let maxY = -32768
  let curX = 0
  let curY = 0
  for (let i = 0; i < totalPoints; i++) {
    const enc = TRIPLET[flags[i]! & 0x7f]!
    const more = enc[0] - 1
    const coords = new Reader(sIn.bytes(more))
    const dx = coords.readNBits(enc[1])
    const dy = coords.readNBits(enc[2])
    if (coords.pos !== coords.size || coords.bitPos !== 0) throw new MtxError('triplet width mismatch')
    xs[i] = i16(enc[5] * (dx + enc[3]))
    ys[i] = i16(enc[6] * (dy + enc[4]))
    curX = i16(curX + xs[i]!)
    curY = i16(curY + ys[i]!)
    if (curX < minX) minX = curX
    if (curX > maxX) maxX = curX
    if (curY < minY) minY = curY
    if (curY > maxY) maxY = curY
  }
  const codeSizeAt = out.pos
  out.skip(2)
  const pushCount = read255UShort(sIn)
  decodePushInstructions(streams[1], out, pushCount)
  const codeSize = read255UShort(sIn)
  out.bytes(streams[2].bytes(codeSize))
  const unpackedCodeSize = out.pos - (codeSizeAt + 2)
  for (let i = 0; i < totalPoints; i++) {
    out.u8(makeFlags(xs[i]!, ys[i]!, (flags[i]! & 0x80) === 0, i === 0))
  }
  for (const arr of [xs, ys]) {
    for (let i = 0; i < totalPoints; i++) {
      let v = arr[i]!
      if (i === 0 || v !== 0) {
        if (v > -256 && v < 0) v = -v
        if (v >= 0 && v < 256) out.u8(v)
        else out.s16(v)
      }
    }
  }
  const end = out.pos
  out.pos = codeSizeAt
  out.u16(unpackedCodeSize)
  if (!bbox) {
    out.pos = bboxAt
    out.s16(minX)
    out.s16(minY)
    out.s16(maxX)
    out.s16(maxY)
  }
  out.pos = end
}

function decodeCompositeGlyph(streams: [Reader, Reader, Reader], out: Writer): void {
  const ARGS_WORDS = 0x1
  const HAVE_SCALE = 0x8
  const MORE_COMPONENTS = 0x20
  const HAVE_XY_SCALE = 0x40
  const HAVE_2_BY_2 = 0x80
  const HAVE_INSTR = 0x100
  const sIn = streams[0]
  out.s16(-1)
  for (let i = 0; i < 4; i++) out.s16(sIn.s16())
  let flags: number
  do {
    flags = sIn.u16()
    out.u16(flags)
    out.bytes(sIn.bytes(2)) // glyph index
    out.bytes(sIn.bytes(flags & ARGS_WORDS ? 4 : 2))
    const transform = flags & HAVE_2_BY_2 ? 8 : flags & HAVE_XY_SCALE ? 4 : flags & HAVE_SCALE ? 2 : 0
    if (transform) out.bytes(sIn.bytes(transform))
  } while (flags & MORE_COMPONENTS)
  if (flags & HAVE_INSTR) {
    const numInstrAt = out.pos
    out.skip(2)
    const pushCount = read255UShort(sIn)
    decodePushInstructions(streams[1], out, pushCount)
    const codeSize = read255UShort(sIn)
    out.bytes(streams[2].bytes(codeSize))
    const numInstr = out.pos - (numInstrAt + 2)
    const end = out.pos
    out.pos = numInstrAt
    out.u16(numInstr)
    out.pos = end
  }
}

function decodeGlyph(streams: [Reader, Reader, Reader], out: Writer): void {
  const sIn = streams[0]
  let numContours = sIn.s16()
  if (numContours < 0) {
    decodeCompositeGlyph(streams, out)
    return
  }
  let bbox: [number, number, number, number] | null = null
  if (numContours === 0x7fff) {
    numContours = sIn.s16()
    bbox = [sIn.s16(), sIn.s16(), sIn.s16(), sIn.s16()]
  }
  decodeSimpleGlyph(numContours, streams, out, bbox)
}

function unpackCvt(tbl: Table, sIn: Reader): void {
  sIn.seek(tbl.offset)
  const n = sIn.u16()
  const out = new Writer(n * 2)
  let last = 0
  for (let i = 0; i < n; i++) {
    const code = sIn.u8()
    let val: number
    if (code >= 248) val = 238 * (code - 247) + sIn.u8()
    else if (code >= 239) val = -(238 * (code - 239) + sIn.u8())
    else if (code === 238) val = sIn.s16()
    else val = code
    last = i16(last + val)
    out.s16(last)
  }
  tbl.data = out.result()
}

function populateGlyfAndLoca(
  glyf: Table,
  loca: Table,
  indexToLocFormat: number,
  numGlyphs: number,
  streams: [Reader, Reader, Reader],
): number {
  streams[0].seek(glyf.offset)
  streams[1].seek(0)
  streams[2].seek(0)
  const out = new Writer(Math.max(1024, numGlyphs * 64))
  const offsets = [0]
  for (let i = 0; i < numGlyphs; i++) {
    decodeGlyph(streams, out)
    if (out.pos % 2) out.u8(0)
    offsets.push(out.pos)
  }
  // The rebuilt glyf writes one flag byte per point (no REPEAT runs), so it can outgrow the
  // 128 KB a short loca addresses even when the source used format 0 — promote to long
  const shortLoca = indexToLocFormat === 0 && out.pos <= 0x1fffe
  const locaOut = new Writer((shortLoca ? 2 : 4) * offsets.length)
  for (const o of offsets) {
    if (shortLoca) locaOut.u16(o / 2)
    else locaOut.u32(o)
  }
  glyf.data = out.result()
  loca.data = locaOut.result()
  return shortLoca ? 0 : 1
}

function checksum(data: Uint8Array): number {
  let sum = 0
  for (let i = 0; i < data.length; i += 4) {
    const chunk =
      ((data[i]! << 24) |
        ((data[i + 1] ?? 0) << 16) |
        ((data[i + 2] ?? 0) << 8) |
        (data[i + 3] ?? 0)) >>>
      0
    sum = (sum + chunk) >>> 0
  }
  return sum
}

function dumpContainer(unsorted: Table[]): Uint8Array {
  // the directory must be in ascending tag order (a synthesized loca would otherwise land last)
  const tables = [...unsorted].sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0))
  const n = tables.length
  let searchRange = 1
  let entrySelector = 0
  while (searchRange * 2 <= n) {
    searchRange *= 2
    entrySelector++
  }
  searchRange *= 16
  const out = new Writer(12 + 16 * n + tables.reduce((a, t) => a + ((t.data?.length ?? 0) + 3), 0))
  out.u32(0x00010000)
  out.u16(n)
  out.u16(searchRange)
  out.u16(entrySelector)
  out.u16(n * 16 - searchRange)
  const dirAt = out.pos
  out.skip(16 * n)
  let total = 0
  let headOffset = -1
  const sums: number[] = []
  for (const t of tables) {
    const data = t.data ?? new Uint8Array(0)
    t.offset = out.pos
    t.length = data.length
    if (t.tag === 'head') headOffset = t.offset
    out.bytes(data)
    while (out.pos % 4) out.u8(0)
    const s = checksum(data)
    sums.push(s)
    total = (total + s) >>> 0
  }
  const end = out.pos
  out.pos = dirAt
  tables.forEach((t, i) => {
    for (let k = 0; k < 4; k++) out.u8(t.tag.charCodeAt(k))
    out.u32(sums[i]!)
    out.u32(t.offset)
    out.u32(t.length)
  })
  const header = out.result().subarray(0, dirAt + 16 * n)
  total = (total + checksum(header)) >>> 0
  if (headOffset < 0) throw new MtxError('no head table')
  out.pos = headOffset + 8
  out.u32((0xb1b0afba - total) >>> 0)
  out.pos = end
  return out.result()
}

/**
 * cmap repairs for what Chromium's font sanitizer rejects (fontTools/opentype.js shrug):
 * PowerPoint's embedder leaves a format-14 record whose length runs far past the table, and
 * Windows-platform format-4/12 subtables carrying a non-zero language. Records that do not
 * fit are dropped, the language field is zeroed; a clean table is returned as is.
 */
export function sanitizeCmap(data: Uint8Array): Uint8Array {
  if (data.length < 4) return data
  const r = new Reader(data)
  r.u16()
  const numTables = r.u16()
  if (data.length < 4 + 8 * numTables) return data
  const keep: Array<{ pid: number; eid: number; off: number; len: number; langAt: number }> = []
  let dirty = false
  const u32At = (o: number) =>
    ((data[o]! << 24) | (data[o + 1]! << 16) | (data[o + 2]! << 8) | data[o + 3]!) >>> 0
  for (let i = 0; i < numTables; i++) {
    const pid = r.u16()
    const eid = r.u16()
    const off = r.u32()
    if (off + 4 > data.length) {
      dirty = true
      continue
    }
    const fmt = (data[off]! << 8) | data[off + 1]!
    // formats 8/10/12/13: reserved u16 then u32 length; format 14: u32 length right after the format
    const long = fmt === 8 || fmt === 10 || fmt === 12 || fmt === 13 || fmt === 14
    if (long && off + 8 > data.length) {
      dirty = true
      continue
    }
    const len = long ? u32At(off + (fmt === 14 ? 2 : 4)) : (data[off + 2]! << 8) | data[off + 3]!
    if (len < 4 || off + len > data.length) {
      dirty = true
      continue
    }
    // language: u16 at +4 (format 4), u32 at +8 (format 12); must be 0 on platform 3
    let langAt = -1
    if (pid === 3 && fmt === 4 && ((data[off + 4]! << 8) | data[off + 5]!) !== 0) langAt = off + 4
    if (pid === 3 && fmt === 12 && u32At(off + 8) !== 0) langAt = off + 8
    if (langAt >= 0) dirty = true
    keep.push({ pid, eid, off, len, langAt })
  }
  if (!dirty) return data
  const out = new Writer(data.length)
  out.u16(0)
  out.u16(keep.length)
  const placed = new Map<number, number>()
  let cursor = 4 + 8 * keep.length
  for (const k of keep) {
    if (!placed.has(k.off)) {
      placed.set(k.off, cursor)
      cursor += k.len
    }
  }
  for (const k of keep) {
    out.u16(k.pid)
    out.u16(k.eid)
    out.u32(placed.get(k.off)!)
  }
  for (const k of keep) {
    const dst = placed.get(k.off)!
    out.pos = dst
    out.bytes(data.subarray(k.off, k.off + k.len))
    if (k.langAt >= 0) {
      out.pos = dst + (k.langAt - k.off)
      if (k.langAt === k.off + 4) out.u16(0)
      else out.u32(0)
    }
  }
  out.pos = out.size
  return out.result()
}

/** Rebuild an sfnt from the three inflated CTF streams (parseCTF + dumpContainer). */
export function ctfToSfnt(streams: [Uint8Array, Uint8Array, Uint8Array]): Uint8Array {
  const rs: [Reader, Reader, Reader] = [new Reader(streams[0]), new Reader(streams[1]), new Reader(streams[2])]
  const s0 = rs[0]
  s0.u32() // scaler type
  const numTables = s0.u16()
  s0.u16()
  s0.u16()
  s0.u16()
  const tables: Table[] = []
  for (let i = 0; i < numTables; i++) {
    const tag = String.fromCharCode(s0.u8(), s0.u8(), s0.u8(), s0.u8())
    if (tag === 'hdmx' || tag === 'VDMX') {
      // device metrics are not reconstructed (libeot parity); rasterizers ignore them
      s0.skip(12)
      continue
    }
    s0.skip(4) // checksum
    const offset = s0.u32()
    const length = s0.u32()
    tables.push({ tag, offset, length, data: null })
  }
  let glyf: Table | undefined
  let loca: Table | undefined
  let maxp: Table | undefined
  let head: Table | undefined
  for (const t of tables) {
    if (t.tag === 'loca') loca = t
    else if (t.tag === 'glyf') glyf = t
    else if (t.tag === 'cvt ') unpackCvt(t, s0)
    else {
      s0.seek(t.offset)
      t.data = s0.bytes(t.length).slice()
      if (t.tag === 'cmap') t.data = sanitizeCmap(t.data)
      if (t.tag === 'maxp') maxp = t
      if (t.tag === 'head') {
        head = t
        if (t.data.length < 52) throw new MtxError('malformed head table')
        t.data.fill(0, 8, 12) // checksumAdjustment is recomputed on dump
      }
    }
  }
  if (!head || !maxp) throw new MtxError('CTF lacks head/maxp')
  if (glyf && !loca) {
    loca = { tag: 'loca', offset: 0, length: 0, data: null }
    tables.push(loca)
  }
  if (glyf && loca) {
    const headView = new DataView(head.data!.buffer, head.data!.byteOffset, head.data!.byteLength)
    const indexToLocFormat = headView.getInt16(50, false)
    const maxpView = new DataView(maxp.data!.buffer, maxp.data!.byteOffset, maxp.data!.byteLength)
    const numGlyphs = maxpView.getUint16(4, false)
    const written = populateGlyfAndLoca(glyf, loca, indexToLocFormat, numGlyphs, rs)
    if (written !== indexToLocFormat) headView.setInt16(50, written, false)
  }
  for (const t of tables) if (!t.data) t.data = new Uint8Array(0)
  return dumpContainer(tables)
}
