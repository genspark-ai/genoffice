/**
 * Minimal WOFF2 reader for advance-width assertions on the bundled KR subsets.
 *
 * Only supports what those files use: CFF flavor (no glyf/loca transform),
 * untransformed hmtx, cmap formats 4/12. Avoids adding a woff2 dependency.
 */
import { readFileSync } from 'node:fs'
import { brotliDecompressSync } from 'node:zlib'

const KNOWN_TAGS = [
  'cmap',
  'head',
  'hhea',
  'hmtx',
  'maxp',
  'name',
  'OS/2',
  'post',
  'cvt ',
  'fpgm',
  'glyf',
  'loca',
  'prep',
  'CFF ',
  'VORG',
  'EBDT',
  'EBLC',
  'gasp',
  'hdmx',
  'kern',
  'LTSH',
  'PCLT',
  'VDMX',
  'vhea',
  'vmtx',
  'BASE',
  'GDEF',
  'GPOS',
  'GSUB',
  'EBSC',
  'JSTF',
  'MATH',
  'CBDT',
  'CBLC',
  'COLR',
  'CPAL',
  'SVG ',
  'sbix',
  'acnt',
  'avar',
  'bdat',
  'bloc',
  'bsln',
  'cvar',
  'fdsc',
  'feat',
  'fmtx',
  'fvar',
  'gvar',
  'hsty',
  'just',
  'lcar',
  'mort',
  'morx',
  'opbd',
  'prop',
  'trak',
  'Zapf',
  'Silf',
  'Glat',
  'Gloc',
  'Feat',
  'Sill',
]

function uintBase128(buf: Buffer, off: number): [number, number] {
  let v = 0
  for (let i = 0; i < 5; i++) {
    const b = buf[off++]
    v = v * 128 + (b & 0x7f)
    if (!(b & 0x80)) return [v, off]
  }
  throw new Error('bad UIntBase128')
}

interface Woff2Font {
  unitsPerEm: number
  tables: Map<string, Buffer>
  cmap: Map<number, number>
  advances: number[]
}

export function readWoff2(path: string): Woff2Font {
  const buf = readFileSync(path)
  if (buf.toString('latin1', 0, 4) !== 'wOF2') throw new Error('not woff2')
  const numTables = buf.readUInt16BE(12)
  const totalCompressedSize = buf.readUInt32BE(20)
  let off = 48
  const entries: { tag: string; len: number }[] = []
  for (let i = 0; i < numTables; i++) {
    const flags = buf[off++]
    let tag: string
    if ((flags & 0x3f) === 0x3f) {
      tag = buf.toString('latin1', off, off + 4)
      off += 4
    } else {
      tag = KNOWN_TAGS[flags & 0x3f]
    }
    const transform = (flags >> 6) & 0x3
    let len: number
    ;[len, off] = uintBase128(buf, off)
    const transformed = tag === 'glyf' || tag === 'loca' ? transform !== 3 : transform !== 0
    if (transformed) throw new Error(`transformed table ${tag} unsupported`)
    entries.push({ tag, len })
  }
  const raw = brotliDecompressSync(buf.subarray(off, off + totalCompressedSize))
  const tables = new Map<string, Buffer>()
  let pos = 0
  for (const { tag, len } of entries) {
    tables.set(tag, raw.subarray(pos, pos + len))
    pos += len
  }
  const head = tables.get('head')!
  const hhea = tables.get('hhea')!
  const maxp = tables.get('maxp')!
  const hmtxBuf = tables.get('hmtx')!
  const numGlyphs = maxp.readUInt16BE(4)
  const numHMetrics = hhea.readUInt16BE(34)
  const advances: number[] = []
  for (let g = 0; g < numGlyphs; g++) {
    advances.push(
      g < numHMetrics ? hmtxBuf.readUInt16BE(g * 4) : hmtxBuf.readUInt16BE((numHMetrics - 1) * 4),
    )
  }
  return {
    unitsPerEm: head.readUInt16BE(18),
    tables,
    cmap: parseCmap(tables.get('cmap')!),
    advances,
  }
}

function parseCmap(cmap: Buffer): Map<number, number> {
  const n = cmap.readUInt16BE(2)
  let best = -1
  let bestScore = -1
  for (let i = 0; i < n; i++) {
    const platform = cmap.readUInt16BE(4 + i * 8)
    const encoding = cmap.readUInt16BE(6 + i * 8)
    const offset = cmap.readUInt32BE(8 + i * 8)
    const format = cmap.readUInt16BE(offset)
    const score =
      platform === 3 && encoding === 10 && format === 12
        ? 3
        : platform === 3 && encoding === 1 && format === 4
          ? 2
          : format === 12 || format === 4
            ? 1
            : 0
    if (score > bestScore) {
      bestScore = score
      best = offset
    }
  }
  const out = new Map<number, number>()
  const format = cmap.readUInt16BE(best)
  if (format === 12) {
    const nGroups = cmap.readUInt32BE(best + 12)
    for (let i = 0; i < nGroups; i++) {
      const p = best + 16 + i * 12
      const start = cmap.readUInt32BE(p)
      const end = cmap.readUInt32BE(p + 4)
      const gid = cmap.readUInt32BE(p + 8)
      for (let c = start; c <= end; c++) out.set(c, gid + (c - start))
    }
  } else if (format === 4) {
    const segCountX2 = cmap.readUInt16BE(best + 6)
    const segCount = segCountX2 / 2
    const endP = best + 14
    const startP = endP + segCountX2 + 2
    const deltaP = startP + segCountX2
    const rangeP = deltaP + segCountX2
    for (let s = 0; s < segCount; s++) {
      const end = cmap.readUInt16BE(endP + s * 2)
      const start = cmap.readUInt16BE(startP + s * 2)
      const delta = cmap.readInt16BE(deltaP + s * 2)
      const rangeOff = cmap.readUInt16BE(rangeP + s * 2)
      for (let c = start; c <= end && c !== 0xffff; c++) {
        let gid: number
        if (rangeOff === 0) {
          gid = (c + delta) & 0xffff
        } else {
          const gi = rangeP + s * 2 + rangeOff + (c - start) * 2
          gid = cmap.readUInt16BE(gi)
          if (gid !== 0) gid = (gid + delta) & 0xffff
        }
        if (gid !== 0) out.set(c, gid)
      }
    }
  } else {
    throw new Error(`cmap format ${format} unsupported`)
  }
  return out
}

/** advance of a codepoint in em (throws if unmapped) */
export function advanceEm(font: Woff2Font, cp: number): number {
  const gid = font.cmap.get(cp)
  if (gid === undefined) throw new Error(`U+${cp.toString(16)} not in cmap`)
  return font.advances[gid] / font.unitsPerEm
}
