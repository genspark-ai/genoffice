import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

// advanceWidths resolves a face through the system font index, so the index is
// swapped for a map pointing at the synthetic faces built below. Everything else
// (readTableDir/readTable/norm/styleScore) stays real, so the whole
// open -> table dir -> parseCmap -> glyphOf path runs unmocked.
const index = vi.hoisted(() => new Map<string, unknown[]>())
vi.mock('../src/sfnt', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/sfnt')>()
  return {
    ...actual,
    getFontIndex: () => ({ byPs: new Map(), byFamily: index }),
  }
})

const { advanceWidths } = await import('../src/advance')

const u16 = (n: number): Buffer => {
  const b = Buffer.alloc(2)
  b.writeUInt16BE(n & 0xffff)
  return b
}
const u32 = (n: number): Buffer => {
  const b = Buffer.alloc(4)
  b.writeUInt32BE(n >>> 0)
  return b
}

const UNITS_PER_EM = 1000
const ADVANCES = [500, 600, 700, 800]
const FAMILY = 'crafted cmap family'

/** cmap holding a single (3,1) format-4 subtable of exactly `sub` bytes */
function cmapTable(sub: Buffer): Buffer {
  return Buffer.concat([
    u16(0), // version
    u16(1), // numTables
    u16(3), // platformID: Windows
    u16(1), // encodingID: BMP -> the rank parseCmap selects
    u32(12), // subtable offset
    sub,
  ])
}

/**
 * cmap format 4 whose declared segCountX2 claims far more segment arrays than
 * the bytes actually present, so the per-segment reads run off the subtable.
 * `size` picks which of the three unguarded reads is the one that escapes:
 * the endCode read is bounds-checked, the startCode/idDelta/idRangeOffset
 * reads that follow it are not.
 */
function truncatedFormat4(size: number): Buffer {
  const SEG_COUNT_X2 = 0x0100 // 128 segments -> 1040 bytes if fully present
  const sub = Buffer.alloc(Math.max(size, 16))
  sub.writeUInt16BE(4, 0) // format
  sub.writeUInt16BE(sub.length, 2) // length
  sub.writeUInt16BE(0, 4) // language
  sub.writeUInt16BE(SEG_COUNT_X2, 6) // segCountX2
  // endCode[0..127] at byte 14: every segment present spans the probe
  // codepoint, so the binary search never narrows away from the first mid it
  // visits. Only the entries that physically fit get written, so a subtable
  // truncated inside the endCode array stays well-formed up to its own end.
  for (let k = 0; 14 + 2 * k + 2 <= sub.length && k < SEG_COUNT_X2 / 2; k++) {
    sub.writeUInt16BE(0xffff, 14 + 2 * k)
  }
  return cmapTable(sub)
}

/** format-4 subtable too short to even carry its own segCountX2 */
function stubFormat4(): Buffer {
  const sub = Buffer.alloc(6) // parseCmap accepts anything down to 4 bytes
  sub.writeUInt16BE(4, 0) // format
  sub.writeUInt16BE(6, 2) // length
  return cmapTable(sub)
}

/** minimal well-formed 2-segment format 4 mapping U+0078 to glyph 2 */
function wellFormedFormat4(): Buffer {
  const sub = Buffer.alloc(32)
  sub.writeUInt16BE(4, 0) // format
  sub.writeUInt16BE(32, 2) // length
  sub.writeUInt16BE(0, 4) // language
  sub.writeUInt16BE(4, 6) // segCountX2: 2 segments
  sub.writeUInt16BE(4, 8) // searchRange
  sub.writeUInt16BE(1, 10) // entrySelector
  sub.writeUInt16BE(0, 12) // rangeShift
  sub.writeUInt16BE(0x0078, 14) // endCode[0]
  sub.writeUInt16BE(0xffff, 16) // endCode[1]: mandatory terminator
  sub.writeUInt16BE(0, 18) // reservedPad
  sub.writeUInt16BE(0x0078, 20) // startCode[0]
  sub.writeUInt16BE(0xffff, 22) // startCode[1]
  sub.writeInt16BE(-0x78 + 2, 24) // idDelta[0]: 0x78 + (-118) = glyph 2
  sub.writeInt16BE(1, 26) // idDelta[1]
  sub.writeUInt16BE(0, 28) // idRangeOffset[0]: direct delta mapping
  sub.writeUInt16BE(0, 30) // idRangeOffset[1]
  return cmapTable(sub)
}

/** sfnt with head/hhea/hmtx/cmap, enough for parseFaceAdvances to accept it */
function font(cmap: Buffer): string {
  const head = Buffer.alloc(54)
  head.writeUInt16BE(UNITS_PER_EM, 18)
  const hhea = Buffer.alloc(36)
  hhea.writeUInt16BE(ADVANCES.length, 34) // numberOfHMetrics
  const hmtx = Buffer.alloc(4 * ADVANCES.length)
  ADVANCES.forEach((a, i) => hmtx.writeUInt16BE(a, 4 * i))

  const tables: [string, Buffer][] = [
    ['cmap', cmap],
    ['hhea', hhea],
    ['hmtx', hmtx],
    ['head', head],
  ]
  const dir = Buffer.alloc(16 * tables.length)
  let cursor = 12 + dir.length
  const body: Buffer[] = []
  tables.forEach(([tag, data], i) => {
    dir.write(tag, 16 * i, 4, 'latin1')
    dir.writeUInt32BE(0, 16 * i + 4) // checksum: unread
    dir.writeUInt32BE(cursor, 16 * i + 8)
    dir.writeUInt32BE(data.length, 16 * i + 12)
    const pad = Buffer.alloc((4 - (data.length % 4)) % 4)
    body.push(data, pad)
    cursor += data.length + pad.length
  })
  const header = Buffer.concat([u32(0x00010000), u16(tables.length), u16(0), u16(0), u16(0)])

  const path = join(mkdtempSync(join(tmpdir(), 'font-metrics-cmap-')), 'crafted.ttf')
  writeFileSync(path, Buffer.concat([header, dir, ...body]))
  index.set('craftedcmapfamily', [{ path, offset: 0, style: 'regular' }])
  return path
}

describe('advanceWidths with a truncated cmap format-4 subtable', () => {
  // 300/500/700 bytes are each past what the checked endCode read needs (143)
  // but short of the next array: startCode (401), idDelta (657), idRangeOffset
  // (913). 40 bytes is the deep-truncation case the existing endCode guard
  // already rejects, kept so the boundary stays covered.
  const sizes = [40, 300, 500, 700]

  for (const size of sizes) {
    it(`does not throw on a ${size}-byte format-4 subtable declaring 128 segments`, () => {
      font(truncatedFormat4(size))
      const widths = advanceWidths(FAMILY, 'xyz', 12)
      expect(widths).toHaveLength(3)
      // unmapped codepoints degrade to NaN (summed as unusable) rather than
      // throwing out of the caller's frame
      for (const w of widths!) expect(Number.isNaN(w) || w > 0).toBe(true)
    })
  }

  it('does not throw on a 6-byte format-4 subtable with no segment arrays', () => {
    font(stubFormat4())
    const widths = advanceWidths(FAMILY, 'xyz', 12)
    expect(widths).toHaveLength(3)
    for (const w of widths!) expect(Number.isNaN(w)).toBe(true)
  })

  it('still resolves a codepoint mapped inside a well-formed subtable', () => {
    font(wellFormedFormat4())
    expect(advanceWidths(FAMILY, 'x', 12)).toEqual([ADVANCES[2]! * ((12 * 20) / UNITS_PER_EM)])
  })
})
