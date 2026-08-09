/**
 * Locate an installed font matching a PDF font's PostScript name (nameID 6) or family
 * (nameID 1/16), for text-edit rebuilds that keep the original typeface. The index is
 * built from targeted fd reads of each face's name table only — CJK collections run
 * 100MB+ (PingFang.ttc ~180MB), so whole-file reads at scan time are not an option.
 */
import { closeSync, openSync, readFileSync, readdirSync, readSync, statSync } from 'node:fs'
import type { Dirent } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'

const norm = (s: string) =>
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

const TTC_TAG = 0x74746366 // 'ttcf'
const OTTO_TAG = 0x4f54544f // CFF-flavored sfnt (e.g. PingFang on recent macOS)
const SFNT_FLAVORS = new Set([0x00010000, 0x74727565, OTTO_TAG]) // 1.0 / 'true' / 'OTTO'

/** Picks FPDFText_LoadFont's font_type: glyf faces embed as FontFile2 (TRUETYPE),
    CFF faces as FontFile3 (TYPE1) */
export const isTruetype = (b: Buffer): boolean => b.length >= 4 && b.readUInt32BE(0) !== OTTO_TAG

function readAt(fd: number, pos: number, len: number): Buffer {
  const buf = Buffer.alloc(len)
  if (readSync(fd, buf, 0, len, pos) !== len) throw new Error('short read')
  return buf
}

interface FaceNames {
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

/** Names of the face whose offset table starts at `offset`; null for non-sfnt faces */
function readFaceNames(fd: number, offset: number): FaceNames | null {
  const head = readAt(fd, offset, 12)
  if (!SFNT_FLAVORS.has(head.readUInt32BE(0))) return null
  const numTables = head.readUInt16BE(4)
  if (numTables === 0 || numTables > 64) return null
  const dir = readAt(fd, offset + 12, 16 * numTables)
  for (let t = 0; t < numTables; t++) {
    if (dir.toString('latin1', 16 * t, 16 * t + 4) !== 'name') continue
    const tblOff = dir.readUInt32BE(16 * t + 8)
    const tblLen = dir.readUInt32BE(16 * t + 12)
    if (tblLen < 6 || tblLen > 1 << 20) return null
    return parseNames(readAt(fd, tblOff, tblLen))
  }
  return null
}

interface FaceRef {
  path: string
  /** Position of the face's offset table within the file (0 unless .ttc) */
  offset: number
  /** Normalized family + subfamily text, for style ranking on family matches */
  style: string
}

interface FontIndex {
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

function indexFile(path: string, index: FontIndex): void {
  let fd: number
  try {
    fd = openSync(path, 'r')
  } catch {
    return
  }
  try {
    const offsets =
      readAt(fd, 0, 4).readUInt32BE(0) === TTC_TAG
        ? ((): number[] => {
            const n = readAt(fd, 8, 4).readUInt32BE(0)
            if (n > 64) return []
            const o = readAt(fd, 12, 4 * n)
            return Array.from({ length: n }, (_, i) => o.readUInt32BE(4 * i))
          })()
        : [0]
    for (const offset of offsets) {
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

function buildIndex(): FontIndex {
  const built: FontIndex = { byPs: new Map(), byFamily: new Map() }
  for (const dir of fontDirs()) for (const path of fontFiles(dir, 2)) indexFile(path, built)
  return built
}

/** Extract a single face from a ttc into a standalone sfnt (rewrite the table directory,
    copy table data; passthrough for plain ttf) */
function extractFace(buf: Buffer, offset: number): Buffer {
  if (buf.readUInt32BE(0) !== TTC_TAG) return buf
  const numTables = buf.readUInt16BE(offset + 4)
  let total = 12 + 16 * numTables
  const entries: Array<{ dirPos: number; tOff: number; tLen: number; newOff: number }> = []
  for (let t = 0; t < numTables; t++) {
    const e = offset + 12 + 16 * t
    const tLen = buf.readUInt32BE(e + 12)
    entries.push({ dirPos: e, tOff: buf.readUInt32BE(e + 8), tLen, newOff: total })
    total += (tLen + 3) & ~3
  }
  const out = Buffer.alloc(total)
  buf.copy(out, 0, offset, offset + 12)
  for (let t = 0; t < numTables; t++) {
    const e = entries[t]!
    buf.copy(out, 12 + 16 * t, e.dirPos, e.dirPos + 8)
    out.writeUInt32BE(e.newOff, 12 + 16 * t + 8)
    out.writeUInt32BE(e.tLen, 12 + 16 * t + 12)
    buf.copy(out, e.newOff, e.tOff, e.tOff + e.tLen)
  }
  return out
}

const faceCache = new Map<string, Buffer>()

// "bolditalic" deliberately has no atomic alternative: a BoldItalic face must yield
// ['bold', 'italic'] so it can score against each want separately (a keep-original
// bold toggle on an italic run wants both tokens; an atomic token would score zero).
// "oblique" folds into "italic": families like Helvetica name their slanted faces
// Oblique, and an italic want must still count them as hits.
const styleTokens = (ps: string): string[] =>
  (
    norm(ps).match(
      /semibold|extrabold|bold|italic|oblique|light|thin|medium|heavy|black|regular|w\d/g,
    ) ?? []
  ).map((t) => (t === 'oblique' ? 'italic' : t))

const REGULARISH = ['regular', 'medium', 'w3', 'w4']

/**
 * Standalone sfnt bytes of the installed font best matching (psName, family):
 * exact PostScript name first, then family with the PS name's style tokens
 * (Regular preferred when there are none). Null when nothing matches.
 */
export function findSystemFont(psName: string, family: string): Buffer | null {
  index ??= buildIndex()
  let face = psName ? index.byPs.get(norm(psName)) : undefined
  if (!face && family) {
    const candidates = index.byFamily.get(norm(family))
    if (candidates?.length) {
      const want = styleTokens(psName)
      // Token-level comparison, not substring: a wanted "bold" must not also credit
      // Semibold/ExtraBold faces (the tokenizer consumes the longest alternative
      // first), and unmatched face tokens push style variants below the exact face
      // instead of tying with it (BoldItalic loses to Bold on a bold-only want via
      // the extras penalty, but beats both single-style faces when both are wanted).
      // With no wanted tokens, Regular-ish tokens count as wanted so the Regular
      // face outranks untagged subfamilies.
      const score = (f: FaceRef) => {
        const have = styleTokens(f.style)
        const wanted = want.length > 0 ? want : REGULARISH
        const hits = want.filter((t) => have.includes(t)).length
        const extras = have.filter((t) => !wanted.includes(t)).length
        const regularBonus = want.length === 0 && have.some((t) => REGULARISH.includes(t)) ? 1 : 0
        return hits * 4 - extras + regularBonus
      }
      face = [...candidates].sort((a, b) => score(b) - score(a))[0]
    }
  }
  if (!face) return null
  const key = `${face.path}#${face.offset}`
  let bytes = faceCache.get(key)
  if (!bytes) {
    try {
      bytes = extractFace(readFileSync(face.path), face.offset)
    } catch {
      return null
    }
    if (faceCache.size >= 4) faceCache.clear()
    faceCache.set(key, bytes)
  }
  return bytes
}
