/**
 * Synthetic sfnt bytes for the font-metrics tests: a font file is built from
 * the tables a test cares about, so no binary fixture blob is needed.
 */

/** one name-table record */
export interface NameRecordSpec {
  /** the record's platform (0 = Unicode, 1 = Macintosh, 3 = Windows) */
  platform: number
  /** the record's encoding id */
  encoding: number
  /** the record's language id (a Windows LCID, or the packed Unicode language) */
  language: number
  /** the record's name id (1 = family, 2 = subfamily, 6 = PostScript name) */
  nameId: number
  /** the name text */
  text: string
}

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

/**
 * A name table (format 0) carrying the given records. Unicode and Windows
 * records are UTF-16BE, Macintosh records Latin-1.
 */
export function buildNameTable(records: readonly NameRecordSpec[]): Buffer {
  const encoded = records.map((r) =>
    r.platform === 0 || r.platform === 3
      ? Buffer.from(r.text, 'utf16le').swap16()
      : Buffer.from(r.text, 'latin1'),
  )
  const header = Buffer.alloc(6)
  header.writeUInt16BE(0, 0) // format
  header.writeUInt16BE(records.length, 2)
  const stringBase = 6 + 12 * records.length
  header.writeUInt16BE(stringBase, 4)
  const dir = Buffer.alloc(12 * records.length)
  let cursor = 0
  records.forEach((r, i) => {
    const at = 12 * i
    dir.writeUInt16BE(r.platform, at)
    dir.writeUInt16BE(r.encoding, at + 2)
    dir.writeUInt16BE(r.language, at + 4)
    dir.writeUInt16BE(r.nameId, at + 6)
    dir.writeUInt16BE(encoded[i]!.length, at + 8)
    dir.writeUInt16BE(cursor, at + 10)
    cursor += encoded[i]!.length
  })
  return Buffer.concat([header, dir, ...encoded])
}

/** A TrueType-outline font file carrying the given tables. */
export function buildSfnt(tables: readonly (readonly [string, Buffer])[]): Buffer {
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
  return Buffer.concat([header, dir, ...body])
}
