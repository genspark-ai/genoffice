import { LegacyXlsError } from './errors'

/**
 * The BIFF8 record walk. A workbook stream is a flat list of records:
 * a 2-byte id, a 2-byte payload length, the payload — with two structural
 * quirks this reader handles explicitly:
 *
 *  - records may be split by CONTINUE records (same stream, new header); a
 *    consumer that asked for `continuable` reads the payload and every
 *    following CONTINUE payload as one logical body,
 *  - inside a split string table, each CONTINUE restarts the character
 *    compression flag, so a segment-aware reader, not plain concatenation,
 *    is what decodes the shared strings.
 */

export const BIFF = {
  bof: 0x0809,
  eof: 0x000a,
  continue: 0x003c,
  boundSheet: 0x0085,
  sst: 0x00fc,
  labelSst: 0x00fd,
  number: 0x0203,
  rk: 0x027e,
  mulRk: 0x00bd,
  mulBlank: 0x00be,
  blank: 0x0201,
  boolErr: 0x0205,
  formula: 0x0006,
  string: 0x0207,
  dimensions: 0x0200,
  mergedCells: 0x00e5,
  colInfo: 0x007d,
  filePass: 0x002f,
  window2: 0x023e,
} as const

export interface BiffRecord {
  id: number
  /** payload length as the header declares */
  length: number
  /** offset of the payload inside the stream */
  offset: number
}

export const RECORD_HEADER = 4

/** Walk every record of the stream without materializing payloads. */
export function* walkRecords(stream: Uint8Array): Generator<BiffRecord> {
  let at = 0
  while (at + RECORD_HEADER <= stream.length) {
    const id = stream[at]! | (stream[at + 1]! << 8)
    const length = stream[at + 2]! | (stream[at + 3]! << 8)
    const offset = at + RECORD_HEADER
    if (offset + length > stream.length) {
      throw new LegacyXlsError('truncated', `record ${id.toString(16)} claims ${length} bytes`)
    }
    yield { id, length, offset }
    at = offset + length
  }
}

/** Payload of one record as a view (no copy). */
export function payloadOf(stream: Uint8Array, record: BiffRecord): Uint8Array {
  return stream.subarray(record.offset, record.offset + record.length)
}

/** The payload of `record` plus every CONTINUE payload that follows it, as
    separate segments: a consumer reads across them and re-checks the flag
    each segment boundary (the shared-string table rule). */
export function recordSegments(stream: Uint8Array, record: BiffRecord): { segments: Uint8Array[] } {
  const segments: Uint8Array[] = [payloadOf(stream, record)]
  let at = record.offset + record.length
  while (at + RECORD_HEADER <= stream.length) {
    const id = stream[at]! | (stream[at + 1]! << 8)
    if (id !== BIFF.continue) break
    const length = stream[at + 2]! | (stream[at + 3]! << 8)
    const offset = at + RECORD_HEADER
    if (offset + length > stream.length) break
    segments.push(stream.subarray(offset, offset + length))
    at = offset + length
  }
  return { segments }
}

/**
 * A byte cursor across the segments of one logical record. String readers
 * need it because a CONTINUE that lands inside a character array restarts
 * the compression flag with its first byte.
 */
export class SegmentedReader {
  private segment = 0
  private at = 0
  /** the compression flag the current characters read with: set by the
      string's own header, replaced by the first byte of a CONTINUE that
      lands mid-string (MS-XLS 2.1.9) */
  private highByte = false
  /** true while a character array is being read across segment boundaries */
  private midChars = false

  constructor(private readonly segments: Uint8Array[]) {
    if (segments.length === 0) throw new LegacyXlsError('corrupt', 'empty record')
  }

  /** enter the next segment. Mid-string, its first byte is the new
      compression flag; at a string boundary it is not, and the next string's
      own header governs. */
  private advanceSegment(): void {
    if (this.segment + 1 >= this.segments.length) {
      throw new LegacyXlsError('bad_sst', 'string runs past the record')
    }
    this.segment += 1
    this.at = 0
    const seg = this.segments[this.segment]!
    if (seg.length === 0) {
      this.advanceSegment()
      return
    }
    if (this.midChars) {
      this.highByte = (seg[0]! & 1) === 1
      this.at = 1
    }
  }

  u8(): number {
    this.ensure(1)
    const seg = this.segments[this.segment]!
    return seg[this.at++]!
  }

  u16(): number {
    return this.u8() | (this.u8() << 8)
  }

  /** read `count` characters compressed per `high`, crossing segments */
  chars(count: number, high: boolean): string {
    this.highByte = high
    this.midChars = true
    let out = ''
    let left = count
    while (left > 0) {
      if (this.at >= this.segments[this.segment]!.length) {
        this.advanceSegment()
        continue
      }
      const seg = this.segments[this.segment]!
      // in the high-byte mode every character costs two segment bytes
      const take = this.highByte
        ? Math.min(left, (seg.length - this.at) >> 1)
        : Math.min(left, seg.length - this.at)
      if (this.highByte) {
        out += Buffer.from(seg.buffer, seg.byteOffset + this.at, take * 2).toString('utf16le')
        this.at += take * 2
      } else {
        out += Buffer.from(seg.buffer, seg.byteOffset + this.at, take).toString('latin1')
        this.at += take
      }
      left -= take
    }
    this.midChars = false
    return out
  }

  rawBytes(count: number): Uint8Array {
    const out = new Uint8Array(count)
    let filled = 0
    while (filled < count) {
      if (this.at >= this.segments[this.segment]!.length) {
        this.advanceSegment()
        continue
      }
      const take = Math.min(count - filled, this.segments[this.segment]!.length - this.at)
      out.set(this.segments[this.segment]!.subarray(this.at, this.at + take), filled)
      this.at += take
      filled += take
    }
    return out
  }

  private ensure(n: number): void {
    while (this.at + n > this.segments[this.segment]!.length) {
      if (this.at >= this.segments[this.segment]!.length) {
        // the cursor sits exactly at a segment boundary: a boundary outside a
        // character array carries no flag byte, the next field just continues
        this.advanceSegment()
        continue
      }
      throw new LegacyXlsError('bad_sst', 'scalar crosses a segment boundary')
    }
  }
}

/** RK value decode (MS-XLS 2.5.122): 30 bits of payload, two flags. */
export function decodeRk(rk: number): number {
  const div100 = (rk & 1) === 1
  const isInt = (rk & 2) === 2
  let value: number
  if (isInt) {
    value = rk >> 2 // arithmetic shift keeps the sign
  } else {
    const buffer = new ArrayBuffer(8)
    const view = new DataView(buffer)
    // the upper 30 bits are the most significant ones of a double
    view.setUint32(0, 0)
    view.setUint32(4, rk & 0xfffffffc)
    value = view.getFloat64(0)
  }
  return div100 ? value / 100 : value
}
