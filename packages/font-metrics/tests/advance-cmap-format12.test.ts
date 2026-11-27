import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

import { buildSfnt } from './helpers/synthetic-font'

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
const FAMILY = 'crafted cmap12 family'

/** cmap holding a single (3,10) format-12 subtable of exactly `sub` bytes */
function cmapTable(sub: Buffer): Buffer {
  return Buffer.concat([
    u16(0), // version
    u16(1), // numTables
    u16(3), // platformID: Windows
    u16(10), // encodingID: full Unicode -> the rank parseCmap selects for format 12
    u32(12), // subtable offset
    sub,
  ])
}

/**
 * format-12 header whose declared nGroups is `nGroups` (a raw uint32, trusted
 * as the binary search's initial hi) and whose language field is `language`.
 * The 16-byte header carries the search state before any group record: with
 * the old (lo + hi) >> 1 a huge nGroups wraps ToInt32 negative and mid = -1
 * re-reads length@4 / language@8 / nGroups@12 as if they were a group record.
 */
function format12Header(nGroups: number, language = 0): Buffer {
  const sub = Buffer.alloc(16)
  sub.writeUInt16BE(12, 0) // format
  sub.writeUInt16BE(0, 2) // reserved
  sub.writeUInt32BE(16, 4) // length: header only, no group bytes follow
  sub.writeUInt32BE(language, 8)
  sub.writeUInt32BE(nGroups >>> 0, 12)
  return sub
}

/** one sorted sequential-group record (startCharCode, endCharCode, startGlyphID) */
function group(start: number, end: number, startGlyph: number): Buffer {
  return Buffer.concat([u32(start), u32(end), u32(startGlyph)])
}

/** minimal well-formed 3-group format 12 mapping a/x/U+0100 to glyphs 1/2/3 */
function wellFormedFormat12(): Buffer {
  return Buffer.concat([
    format12Header(3),
    group(0x0061, 0x0062, 1), // a b
    group(0x0078, 0x0078, 2), // x
    group(0x0100, 0x0101, 3), // Ā ā
  ])
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

  const path = join(mkdtempSync(join(tmpdir(), 'font-metrics-cmap12-')), 'crafted.ttf')
  writeFileSync(path, buildSfnt(tables))
  index.set('craftedcmap12family', [{ path, offset: 0, style: 'regular' }])
  return path
}

describe('advanceWidths with a crafted cmap format-12 subtable', () => {
  it('still resolves codepoints mapped by a well-formed 3-group subtable', () => {
    font(cmapTable(wellFormedFormat12()))
    // 'x' hits the first mid, 'a' narrows hi, 'Ā' narrows lo — all three
    // directions of the search; 'q' falls between the groups and stays unmapped
    expect(advanceWidths(FAMILY, 'axqĀ', 12)).toEqual([
      ADVANCES[1]! * ((12 * 20) / UNITS_PER_EM),
      ADVANCES[2]! * ((12 * 20) / UNITS_PER_EM),
      NaN,
      ADVANCES[3]! * ((12 * 20) / UNITS_PER_EM),
    ])
  })

  it('returns NaN, not a header byte pattern read as a group, when nGroups overflows the int32 mid', () => {
    // nGroups = 0xFFFFFFFF puts hi (4294967294) past the int32 range, so the
    // old (lo + hi) >> 1 yields mid = -1 and reads the header as a group whose
    // start/end are the length/language fields; with language >= cp that
    // "maps" every codepoint to a garbage glyph id (clamped to the last
    // advance) instead of reporting the face unusable. The language field
    // doubles as that fake group's endCharCode for exactly that reason.
    font(cmapTable(format12Header(0xffffffff, 0xffff)))
    const widths = advanceWidths(FAMILY, 'xyz', 12)
    expect(widths).toHaveLength(3)
    for (const w of widths!) expect(Number.isNaN(w)).toBe(true)
  })

  it('does not throw when a huge nGroups drives the mid arithmetic negative', () => {
    // nGroups = 2^32 - 3: the old mid = -2 puts the group read at byte -8,
    // a RangeError escaping advanceWidths' caller (glyphOf runs outside the
    // parse try/catch). The fixed search lands past the buffer and degrades
    // to the unmapped NaN path.
    font(cmapTable(format12Header(4294967293)))
    const widths = advanceWidths(FAMILY, 'xyz', 12)
    expect(widths).toHaveLength(3)
    for (const w of widths!) expect(Number.isNaN(w)).toBe(true)
  })

  it('does not throw on a 6-byte format-12 subtable too short for its own header', () => {
    // parseCmap accepts anything down to 4 bytes, but nGroups lives at offset
    // 12 of the 16-byte header — reading it from a stub would throw out of
    // the caller, like the format-4 stub case
    const sub = Buffer.alloc(6)
    sub.writeUInt16BE(12, 0) // format
    font(cmapTable(sub))
    const widths = advanceWidths(FAMILY, 'xyz', 12)
    expect(widths).toHaveLength(3)
    for (const w of widths!) expect(Number.isNaN(w)).toBe(true)
  })
})
