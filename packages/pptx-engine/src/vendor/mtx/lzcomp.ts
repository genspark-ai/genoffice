/*
 * MicroType Express LZ decoder — TypeScript port of libeot's src/lzcomp
 * (ahuff.c, bitio.c, lzcomp.c, liblzcomp.c). libeot is MPL-2.0 (see LICENSE).
 * Decompression side only; the compressor and the size-limited copy window are not ported.
 */

export class MtxError extends Error {}

/** MSB-first bit writer (BITIO in write mode); only the round-trip tests encode. */
export class BitWriter {
  private out: number[] = []
  private cur = 0
  private count = 0
  outputBit(bit: boolean): void {
    this.cur = (this.cur << 1) | (bit ? 1 : 0)
    if (++this.count === 8) {
      this.out.push(this.cur)
      this.cur = 0
      this.count = 0
    }
  }
  writeValue(value: number, numberOfBits: number): void {
    for (let i = numberOfBits - 1; i >= 0; i--) this.outputBit(((value >>> i) & 1) === 1)
  }
  bytes(): Uint8Array {
    const tail = this.count ? [this.cur << (8 - this.count)] : []
    return new Uint8Array([...this.out, ...tail])
  }
}

/** MSB-first bit reader over a byte buffer (BITIO in read mode). */
export class BitReader {
  private idx = 0
  private bitCount = 0
  private bitBuffer = 0
  constructor(private readonly bytes: Uint8Array) {}

  inputBit(): boolean {
    if (this.bitCount === 0) {
      if (this.idx >= this.bytes.length) throw new MtxError('unexpected end of LZ stream')
      this.bitBuffer = this.bytes[this.idx++]!
      this.bitCount = 8
    }
    this.bitCount--
    this.bitBuffer <<= 1
    return (this.bitBuffer & 0x100) !== 0
  }

  readValue(numberOfBits: number): number {
    let value = 0
    for (let i = 0; i < numberOfBits; i++) {
      value = (value << 1) | (this.inputBit() ? 1 : 0)
    }
    return value >>> 0
  }
}

function bitsUsed(x: number): number {
  let n = 0
  while (x > 0) {
    n++
    x >>>= 1
  }
  return n
}

/** Adaptive Huffman coder (AHUFF): sibling-property tree, leaves carry symbol codes. */
export class AHuff {
  private readonly up: Int32Array
  private readonly left: Int32Array
  private readonly right: Int32Array
  private readonly code: Int32Array
  private readonly weight: Int32Array
  private readonly symbolIndex: Int32Array
  private static readonly ROOT = 1

  constructor(
    private readonly bits: BitReader | null,
    private readonly range: number,
  ) {
    let bitCount2 = 0
    if (range > 256 && range < 512) bitCount2 = bitsUsed(range - 256 - 1) + 1
    const n = 2 * range
    this.up = new Int32Array(n)
    this.left = new Int32Array(n)
    this.right = new Int32Array(n)
    this.code = new Int32Array(n)
    this.weight = new Int32Array(n)
    this.symbolIndex = new Int32Array(range)
    for (let i = 2; i < n; i++) {
      this.up[i] = i >> 1
      this.weight[i] = 1
    }
    for (let i = 1; i < range; i++) {
      this.left[i] = 2 * i
      this.right[i] = 2 * i + 1
    }
    for (let i = 0; i < range; i++) {
      this.code[i] = -1
      this.code[range + i] = i
      this.left[range + i] = -1
      this.right[range + i] = -1
      this.symbolIndex[i] = range + i
    }
    this.initWeight(AHuff.ROOT)
    if (bitCount2 !== 0) {
      this.updateWeight(this.symbolIndex[256]!)
      this.updateWeight(this.symbolIndex[257]!)
      for (let i = 0; i < 12; i++) this.updateWeight(this.symbolIndex[range - 3]!)
      for (let i = 0; i < 6; i++) this.updateWeight(this.symbolIndex[range - 2]!)
    } else {
      for (let j = 0; j < 2; j++) {
        for (let i = 0; i < range; i++) this.updateWeight(this.symbolIndex[i]!)
      }
    }
  }

  private initWeight(a: number): number {
    if (this.code[a]! < 0) {
      this.weight[a] = this.initWeight(this.left[a]!) + this.initWeight(this.right[a]!)
    }
    return this.weight[a]!
  }

  private swapNodes(a: number, b: number): void {
    const { up, left, right, code, weight, symbolIndex } = this
    const upa = up[a]!
    const upb = up[b]!
    const tl = left[a]!
    const tr = right[a]!
    const tc = code[a]!
    const tw = weight[a]!
    left[a] = left[b]!
    right[a] = right[b]!
    code[a] = code[b]!
    weight[a] = weight[b]!
    left[b] = tl
    right[b] = tr
    code[b] = tc
    weight[b] = tw
    up[a] = upa
    up[b] = upb
    for (const x of [a, b]) {
      const c = code[x]!
      if (c < 0) {
        up[left[x]!] = x
        up[right[x]!] = x
      } else {
        symbolIndex[c] = x
      }
    }
  }

  private updateWeight(a: number): void {
    const { up, weight } = this
    for (; a !== AHuff.ROOT; a = up[a]!) {
      const weightA = weight[a]!
      let b = a - 1
      // keep the sibling property: jump over the block of equal weights
      if (weight[b] === weightA) {
        do {
          b--
        } while (weight[b] === weightA)
        b++
        if (b > AHuff.ROOT) {
          this.swapNodes(a, b)
          a = b
        }
      }
      weight[a] = weightA + 1
    }
    weight[AHuff.ROOT]!++
  }

  /** Encoder mirror of readSymbol (MTX_AHUFF_WriteSymbol); used by the round-trip tests. */
  writeSymbol(out: BitWriter, symbol: number): void {
    const { up, right } = this
    const leaf = this.symbolIndex[symbol]!
    const path: boolean[] = []
    for (let a = leaf; a !== AHuff.ROOT; a = up[a]!) path.push(right[up[a]!] === a)
    for (let i = path.length - 1; i >= 0; i--) out.outputBit(path[i]!)
    this.updateWeight(leaf)
  }

  readSymbol(): number {
    const { left, right, code } = this
    let a = AHuff.ROOT
    let symbol: number
    do {
      a = this.bits!.inputBit() ? right[a]! : left[a]!
      symbol = code[a]!
    } while (symbol < 0)
    this.updateWeight(a)
    return symbol
  }
}

/** Run-length layer decoder (RUNLENGTHCOMP::SaveBytes state machine). */
class RunLengthSink {
  private state: 'initial' | 'normal' | 'seenEscape' | 'needByte' = 'initial'
  private escape = 0
  private count = 0
  constructor(private readonly out: ByteSink) {}

  save(value: number): void {
    switch (this.state) {
      case 'normal':
        if (value === this.escape) this.state = 'seenEscape'
        else this.out.push(value)
        break
      case 'seenEscape':
        this.count = value
        if (value === 0) {
          this.out.push(this.escape)
          this.state = 'normal'
        } else {
          this.state = 'needByte'
        }
        break
      case 'needByte':
        for (let i = this.count; i > 0; i--) this.out.push(value)
        this.state = 'normal'
        break
      default:
        this.escape = value
        this.state = 'normal'
    }
  }
}

class ByteSink {
  private buf: Uint8Array
  length = 0
  constructor(capacity: number) {
    this.buf = new Uint8Array(Math.max(capacity, 16))
  }
  push(v: number): void {
    if (this.length >= this.buf.length) {
      const next = new Uint8Array(this.buf.length + (this.buf.length >> 1) + 16)
      next.set(this.buf)
      this.buf = next
    }
    this.buf[this.length++] = v
  }
  bytes(): Uint8Array {
    return this.buf.slice(0, this.length)
  }
}

const PRELOAD_SIZE = 2 * 32 * 96 + 4 * 256
const LEN_WIDTH = 3
const DIST_WIDTH = 3
const BIT_RANGE = LEN_WIDTH - 1
const LEN_MIN = 2
const DIST_MIN = 1
const MAX_2BYTE_DIST = 512

/** Inflate one LZCOMP block (MTX_LZCOMP_UnPackMemory). */
export function unpackLz(data: Uint8Array, version: number): Uint8Array {
  const bits = new BitReader(data)
  const usingRunLength = version === 1 ? false : bits.inputBit()
  const distCoder = new AHuff(bits, 1 << DIST_WIDTH)
  const lenCoder = new AHuff(bits, 1 << LEN_WIDTH)
  const outLen = bits.readValue(24)

  // SetDistRange
  let numDistRanges = 1
  let distMax = DIST_MIN + (1 << (DIST_WIDTH * numDistRanges)) - 1
  while (distMax < outLen) {
    numDistRanges++
    distMax = DIST_MIN + (1 << (DIST_WIDTH * numDistRanges)) - 1
  }
  const DUP2 = 256 + (1 << LEN_WIDTH) * numDistRanges
  const DUP4 = DUP2 + 1
  const DUP6 = DUP4 + 1
  const NUM_SYMS = DUP6 + 1
  const symCoder = new AHuff(bits, NUM_SYMS)

  // InitializeModel: the history window starts with a fixed preload
  const window = new Uint8Array(PRELOAD_SIZE + outLen)
  let i = 0
  for (let k = 0; k < 32; k++) {
    for (let j = 0; j < 96; j++) {
      window[i++] = k
      window[i++] = j
    }
  }
  for (let j = 0; i < PRELOAD_SIZE && j < 256; j++) {
    window[i++] = j
    window[i++] = j
    window[i++] = j
    window[i++] = j
  }

  const sink = new ByteSink(outLen)
  const rl = usingRunLength ? new RunLengthSink(sink) : null
  const emit = (v: number) => (rl ? rl.save(v) : sink.push(v))

  const decodeLength = (symbol: number): { length: number; ranges: number } => {
    const mask = 1 << BIT_RANGE
    let value = 0
    let b = symbol - 256
    const ranges = Math.floor(b / (1 << LEN_WIDTH)) + 1
    b = b % (1 << LEN_WIDTH)
    let done: boolean
    for (;;) {
      done = (b & mask) === 0
      b &= ~mask
      value = (value << BIT_RANGE) | b
      if (done) break
      b = lenCoder.readSymbol()
    }
    return { length: value + LEN_MIN, ranges }
  }
  const decodeDistance = (ranges: number): number => {
    let value = 0
    for (let r = ranges; r > 0; r--) value = (value << DIST_WIDTH) | distCoder.readSymbol()
    return value + DIST_MIN
  }

  const base = PRELOAD_SIZE
  let pos = 0
  while (pos < outLen) {
    const symbol = symCoder.readSymbol()
    let value: number
    if (symbol < 256) {
      value = symbol
    } else if (symbol === DUP2) {
      value = window[base + pos - 2]!
    } else if (symbol === DUP4) {
      value = window[base + pos - 4]!
    } else if (symbol === DUP6) {
      value = window[base + pos - 6]!
    } else {
      const { length: len0, ranges } = decodeLength(symbol)
      const distance = decodeDistance(ranges)
      const length = distance >= MAX_2BYTE_DIST ? len0 + 1 : len0
      const start = pos - distance - length + 1
      if (base + start < 0) throw new MtxError('LZ copy before window start')
      for (let j = 0; j < length; j++) {
        const v = window[base + start + j]!
        window[base + pos++] = v
        emit(v)
      }
      continue
    }
    window[base + pos++] = value
    emit(value)
  }
  return sink.bytes()
}

function be24(b: Uint8Array, o: number): number {
  return (b[o]! << 16) | (b[o + 1]! << 8) | b[o + 2]!
}

/** Split an MTX payload into its three inflated CTF streams (unpackMtx). */
export function unpackMtx(buf: Uint8Array): [Uint8Array, Uint8Array, Uint8Array] {
  if (buf.length < 10) throw new MtxError('MTX header truncated')
  const version = buf[0]!
  const offsets = [10, be24(buf, 4), be24(buf, 7)]
  const sizes = [offsets[1]! - offsets[0]!, offsets[2]! - offsets[1]!, buf.length - offsets[2]!]
  const out: Uint8Array[] = []
  for (let k = 0; k < 3; k++) {
    if (sizes[k]! < 0 || offsets[k]! + sizes[k]! > buf.length) throw new MtxError('MTX block bounds')
    out.push(unpackLz(buf.subarray(offsets[k]!, offsets[k]! + sizes[k]!), version))
  }
  return out as [Uint8Array, Uint8Array, Uint8Array]
}
