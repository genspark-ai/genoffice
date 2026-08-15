/**
 * Shared sfnt/ttc plumbing: system font directory scan, per-face name-table
 * index, and targeted table reads. Everything works from small fd reads of the
 * tables it needs — CJK collections run 100MB+ (PingFang.ttc ~180MB), so
 * whole-file reads at scan time are not an option.
 */
import { closeSync, openSync, readdirSync, readSync, statSync } from 'node:fs'
import type { Dirent } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'

export const norm = (s: string) =>
  s
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[\s\-_]/g, '')

function fontDirs(): string[] {
  switch (process.platform) {
    case 'darwin':
      return [
        '/System/Library/Fonts',
        '/System/Library/Fonts/Supplemental',
        '/Library/Fonts',
        join(homedir(), 'Library/Fonts'),
        // On-demand system font assets: PingFang and other CJK faces live here on
        // recent macOS ( <hash>.asset/AssetData/*.ttc ), not under /System/Library/Fonts
        '/System/Library/AssetsV2/com_apple_MobileAsset_Font7',
        '/System/Library/AssetsV2/com_apple_MobileAsset_Font6',
      ]
    case 'win32':
      return ['C:\\Windows\\Fonts', join(homedir(), 'AppData/Local/Microsoft/Windows/Fonts')]
    default:
      return ['/usr/share/fonts', '/usr/local/share/fonts', join(homedir(), '.fonts')]
  }
}

export const TTC_TAG = 0x74746366 // 'ttcf'
export const OTTO_TAG = 0x4f54544f // CFF-flavored sfnt (e.g. PingFang on recent macOS)
const SFNT_FLAVORS = new Set([0x00010000, 0x74727565, OTTO_TAG]) // 1.0 / 'true' / 'OTTO'

export function readAt(fd: number, pos: number, len: number): Buffer {
  const buf = Buffer.alloc(len)
  if (readSync(fd, buf, 0, len, pos) !== len) throw new Error('short read')
  return buf
}

export interface FaceNames {
  ps: string[]
  families: string[]
  subfamilies: string[]
}

function parseNames(buf: Buffer): FaceNames {
  const out: FaceNames = { ps: [], families: [], subfamilies: [] }
  const count = buf.readUInt16BE(2)
  const strBase = buf.readUInt16BE(4)
  for (let i = 0; i < count; i++) {
    const r = 6 + 12 * i
    if (r + 12 > buf.length) break
    const platform = buf.readUInt16BE(r)
    const encoding = buf.readUInt16BE(r + 2)
    const nameId = buf.readUInt16BE(r + 6)
    const list =
      nameId === 6
        ? out.ps
        : nameId === 1 || nameId === 16
          ? out.families
          : nameId === 2 || nameId === 17
            ? out.subfamilies
            : null
    if (!list) continue
    const len = buf.readUInt16BE(r + 8)
    const off = strBase + buf.readUInt16BE(r + 10)
    if (off + len > buf.length) continue
    let s: string
    if (platform === 0 || platform === 3) {
      s = Buffer.from(buf.subarray(off, off + len))
        .swap16()
        .toString('utf16le')
    } else if (platform === 1 && encoding === 0) {
      s = buf.toString('latin1', off, off + len)
    } else {
      continue // Mac-platform legacy CJK codepages: not decodable here
    }
    if (s && !list.includes(s)) list.push(s)
  }
  return out
}

/** Table directory of the face whose offset table starts at `offset`; null for non-sfnt faces */
export function readTableDir(
  fd: number,
  offset: number,
): Map<string, { off: number; len: number }> | null {
  const head = readAt(fd, offset, 12)
  if (!SFNT_FLAVORS.has(head.readUInt32BE(0))) return null
  const numTables = head.readUInt16BE(4)
  if (numTables === 0 || numTables > 64) return null
  const dir = readAt(fd, offset + 12, 16 * numTables)
  const tables = new Map<string, { off: number; len: number }>()
  for (let t = 0; t < numTables; t++) {
    tables.set(dir.toString('latin1', 16 * t, 16 * t + 4), {
      off: dir.readUInt32BE(16 * t + 8),
      len: dir.readUInt32BE(16 * t + 12),
    })
  }
  return tables
}

/** Bytes of one table, bounded by `maxLen`; null when absent or implausibly sized */
export function readTable(
  fd: number,
  tables: Map<string, { off: number; len: number }>,
  tag: string,
  maxLen: number,
): Buffer | null {
  const entry = tables.get(tag)
  if (!entry || entry.len < 2 || entry.len > maxLen) return null
  return readAt(fd, entry.off, entry.len)
}

/** Names of the face whose offset table starts at `offset`; null for non-sfnt faces */
export function readFaceNames(fd: number, offset: number): FaceNames | null {
  const tables = readTableDir(fd, offset)
  if (!tables) return null
  const name = readTable(fd, tables, 'name', 1 << 20)
  return name && name.length >= 6 ? parseNames(name) : null
}

export interface FaceRef {
  path: string
  /** Position of the face's offset table within the file (0 unless .ttc) */
  offset: number
  /** Normalized family + subfamily text, for style ranking on family matches */
  style: string
}

export interface FontIndex {
  byPs: Map<string, FaceRef>
  byFamily: Map<string, FaceRef[]>
}

function* fontFiles(dir: string, depth: number): Generator<string> {
  let entries: Dirent[]
  try {
    entries = readdirSync(dir, { withFileTypes: true })
  } catch {
    return
  }
  for (const e of entries) {
    const full = join(dir, e.name)
    let isDir = e.isDirectory()
    let isFile = e.isFile()
    if (e.isSymbolicLink()) {
      try {
        const st = statSync(full)
        isDir = st.isDirectory()
        isFile = st.isFile()
      } catch {
        continue
      }
    }
    if (isDir && depth > 0) yield* fontFiles(full, depth - 1)
    else if (isFile && /\.(ttf|otf|ttc|otc)$/i.test(e.name)) yield full
  }
}

/** Offsets of every face's offset table within the file (many for .ttc, [0] otherwise) */
export function faceOffsets(fd: number): number[] {
  if (readAt(fd, 0, 4).readUInt32BE(0) !== TTC_TAG) return [0]
  const n = readAt(fd, 8, 4).readUInt32BE(0)
  if (n > 64) return []
  const o = readAt(fd, 12, 4 * n)
  return Array.from({ length: n }, (_, i) => o.readUInt32BE(4 * i))
}

function indexFile(path: string, index: FontIndex): void {
  let fd: number
  try {
    fd = openSync(path, 'r')
  } catch {
    return
  }
  try {
    for (const offset of faceOffsets(fd)) {
      let names: FaceNames | null
      try {
        names = readFaceNames(fd, offset)
      } catch {
        names = null
      }
      if (!names) continue
      const face: FaceRef = {
        path,
        offset,
        style: norm([...names.families, ...names.subfamilies].join(' ')),
      }
      for (const p of names.ps) if (!index.byPs.has(norm(p))) index.byPs.set(norm(p), face)
      for (const f of names.families) {
        const k = norm(f)
        index.byFamily.set(k, [...(index.byFamily.get(k) ?? []), face])
      }
    }
  } catch {
    /* unreadable or malformed file: skip */
  } finally {
    closeSync(fd)
  }
}

let index: FontIndex | null = null

export function getFontIndex(): FontIndex {
  if (!index) {
    index = { byPs: new Map(), byFamily: new Map() }
    for (const dir of fontDirs()) for (const path of fontFiles(dir, 2)) indexFile(path, index)
  }
  return index
}

// "bolditalic" deliberately has no atomic alternative: a BoldItalic face must yield
// ['bold', 'italic'] so it can score against each want separately (a keep-original
// bold toggle on an italic run wants both tokens; an atomic token would score zero).
// "oblique" folds into "italic": families like Helvetica name their slanted faces
// Oblique, and an italic want must still count them as hits.
export const styleTokens = (ps: string): string[] =>
  (
    norm(ps).match(
      /semibold|extrabold|bold|italic|oblique|light|thin|medium|heavy|black|regular|w\d/g,
    ) ?? []
  ).map((t) => (t === 'oblique' ? 'italic' : t))

const REGULARISH = ['regular', 'medium', 'w3', 'w4']

/**
 * Token-level style score of a face against wanted tokens, not substring: a
 * wanted "bold" must not also credit Semibold/ExtraBold faces (the tokenizer
 * consumes the longest alternative first), and unmatched face tokens push
 * style variants below the exact face instead of tying with it (BoldItalic
 * loses to Bold on a bold-only want via the extras penalty, but beats both
 * single-style faces when both are wanted). With no wanted tokens, Regular-ish
 * tokens count as wanted so the Regular face outranks untagged subfamilies.
 */
export function styleScore(face: FaceRef, want: string[]): number {
  const have = styleTokens(face.style)
  const wanted = want.length > 0 ? want : REGULARISH
  const hits = want.filter((t) => have.includes(t)).length
  const extras = have.filter((t) => !wanted.includes(t)).length
  const regularBonus = want.length === 0 && have.some((t) => REGULARISH.includes(t)) ? 1 : 0
  return hits * 4 - extras + regularBonus
}
