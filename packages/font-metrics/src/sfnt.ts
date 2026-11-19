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

/**
 * Font asset catalogs under AssetsV2 (downloaded and OS-preinstalled), newest catalog
 * version first. The version suffix moves with macOS releases (macOS 26 downloads into
 * com_apple_MobileAsset_Font8 and leaves Font7 empty), so catalogs are discovered rather
 * than hardcoded.
 */
export function appleFontAssetRoots(assetsV2 = '/System/Library/AssetsV2'): string[] {
  const roots: Array<{ path: string; ver: number }> = []
  for (const base of [assetsV2, join(assetsV2, 'PreinstalledAssetsV2/InstallWithOs')]) {
    let names: string[]
    try {
      names = readdirSync(base)
    } catch {
      continue
    }
    for (const d of names) {
      const m = /^com_apple_MobileAsset_Font(\d+)$/.exec(d)
      if (m) roots.push({ path: join(base, d), ver: Number(m[1]) })
    }
  }
  // Newer catalogs first; at equal version the downloaded catalog beats the preinstalled one
  return roots.sort((a, b) => b.ver - a.ver).map((r) => r.path)
}

/**
 * Per-user font directory on Linux (fontconfig's XDG location). Desktop "install font for
 * this user" actions (KDE, GNOME, Dolphin, copy-to-~/.local/share/fonts by hand) land here,
 * so leaving it out of the scan makes those families invisible to metrics while Chromium —
 * which resolves through fontconfig — still draws them. Layout then falls back to heuristic
 * advances: every run after a Latin one starts too early, by roughly half an em, so a CJK
 * glyph covers the last digit of the run before it.
 */
export function xdgUserFontDir(dataHome = process.env.XDG_DATA_HOME): string {
  return join(dataHome || join(homedir(), '.local', 'share'), 'fonts')
}

export function fontDirs(): string[] {
  switch (process.platform) {
    case 'darwin':
      return [
        '/System/Library/Fonts',
        '/System/Library/Fonts/Supplemental',
        '/Library/Fonts',
        join(homedir(), 'Library/Fonts'),
        // On-demand system font assets: PingFang and other CJK faces live here on
        // recent macOS ( <hash>.asset/AssetData/*.ttc ), not under /System/Library/Fonts
        ...appleFontAssetRoots(),
      ]
    case 'win32':
      return ['C:\\Windows\\Fonts', join(homedir(), 'AppData/Local/Microsoft/Windows/Fonts')]
    default:
      return [
        '/usr/share/fonts',
        '/usr/local/share/fonts',
        xdgUserFontDir(),
        join(homedir(), '.fonts'),
      ]
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

/** One family-name record of a face's name table, with the language it was
 *  written for: a face spells its family in several languages, and only the
 *  English (or Unicode-default) spelling is the canonical one. */
export interface FamilyNameRecord {
  /** the name text */
  name: string
  /** the record's name id (1 = family, 16 = typographic family) */
  nameId: number
  /** the record's platform (0 = Unicode, 1 = Macintosh, 3 = Windows) */
  platform: number
  /** the record's language id (a Windows LCID, or the packed Unicode language) */
  language: number
}

export interface FaceNames {
  ps: string[]
  families: string[]
  subfamilies: string[]
  /** every family-name record, with the language id it was written for */
  familyRecords: FamilyNameRecord[]
}

/**
 * A family-name record that spells the face's family in English (or in the
 * Unicode default): the en-US Windows record, the Unicode record with
 * language 0, the legacy Macintosh record. Every other record is a localized
 * spelling of the same face.
 */
function isCanonicalRecord(platform: number, language: number): boolean {
  if (platform === 0) return language === 0 || language === 0x0409
  if (platform === 3) return language === 0 || language === 0x0409
  if (platform === 1) return language === 0
  return false
}

/** The face's canonical (en-US) family spelling; the first family name when
 *  the face carries no English record at all. */
function canonicalFamily(names: FaceNames): string {
  for (const record of names.familyRecords) {
    if (isCanonicalRecord(record.platform, record.language)) return record.name
  }
  return names.families[0] ?? ''
}

function parseNames(buf: Buffer): FaceNames {
  const out: FaceNames = { ps: [], families: [], subfamilies: [], familyRecords: [] }
  const count = buf.readUInt16BE(2)
  const strBase = buf.readUInt16BE(4)
  for (let i = 0; i < count; i++) {
    const r = 6 + 12 * i
    if (r + 12 > buf.length) break
    const platform = buf.readUInt16BE(r)
    const encoding = buf.readUInt16BE(r + 2)
    const language = buf.readUInt16BE(r + 4)
    const nameId = buf.readUInt16BE(r + 6)
    const isFamily = nameId === 1 || nameId === 16
    const list =
      nameId === 6
        ? out.ps
        : isFamily
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
    if (!s) continue
    if (!list.includes(s)) list.push(s)
    if (isFamily) out.familyRecords.push({ name: s, nameId, platform, language })
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
  /** Every family spelling the face carries, the localized ones included */
  families: readonly string[]
  /** The face's canonical (en-US) family spelling */
  canonical: string
  /** Every family record of the face, with the name id it was written for */
  familyRecords: readonly FamilyNameRecord[]
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
        families: [...names.families],
        canonical: canonicalFamily(names),
        familyRecords: [...names.familyRecords],
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

/** Index every font file under `dir` into `index` (the system font
 *  directories, or a directory of one's own). */
export function indexFontDirectory(dir: string, index: FontIndex): void {
  for (const path of fontFiles(dir, 2)) indexFile(path, index)
}

export function getFontIndex(): FontIndex {
  if (!index) {
    index = { byPs: new Map(), byFamily: new Map() }
    for (const dir of fontDirs()) indexFontDirectory(dir, index)
  }
  return index
}

/**
 * The other family spellings of the installed face a declared name belongs to,
 * the canonical (en-US) one first. A document that declares a CJK font by its
 * localized name (\u65b9\u6b63\u4eff\u5b8b_GBK) gets the English spelling the browser
 * resolves (FZFangSong-Z02), and the other way round; a face with a single
 * spelling yields nothing (genoffice#2449).
 *
 * Only the spellings of the same family slot count as aliases: a face writes
 * its family under two name ids, and the legacy one (1) carries the weight
 * with it ("Hiragino Sans GB W3") while the typographic one (16) stays
 * family-wide ("Hiragino Sans GB"). A localized name is matched under the
 * name id the face writes it under, and only the records under that same id
 * are aliases of it — a sibling weight face spells its own legacy family
 * with a weight attached, which is a different family name, not another
 * spelling of this one. Without that, a family of seven weights would offer
 * "<Family> W0" as an alias and body text would render in the thinnest
 * weight whenever the canonical spelling does not resolve.
 *
 * The match is exact after normalization (norm), never fuzzy: a name no
 * installed face carries has no alias.
 */
export function familyAliases(name: string, index: FontIndex = getFontIndex()): readonly string[] {
  const key = norm(name)
  if (!key) return []
  const faces = index.byFamily.get(key)
  if (!faces?.length) return []
  const canonical: string[] = []
  const localized: string[] = []
  const seen = new Set<string>([key])
  for (const face of faces) {
    const records = face.familyRecords ?? []
    // the (platform, name id) slots this face writes the matched name under.
    // The platform is part of the key: an Apple face writes the family-wide
    // spelling under a Macintosh record (platform 1, name id 1) next to its
    // Windows weight-attached records (platform 3, name id 1, "… W0"…) —
    // matching on the name id alone would let the family-wide slot pull every
    // weight face's Windows name back in (genoffice#2469 round 4).
    const matchedSlots = new Set<string>()
    for (const record of records) {
      if (norm(record.name) === key) matchedSlots.add(`${record.platform}:${record.nameId}`)
    }
    if (matchedSlots.size === 0) continue
    for (const record of records) {
      if (!matchedSlots.has(`${record.platform}:${record.nameId}`)) continue
      const recordKey = norm(record.name)
      if (seen.has(recordKey)) continue
      seen.add(recordKey)
      if (isCanonicalRecord(record.platform, record.language)) canonical.push(record.name)
      else localized.push(record.name)
    }
  }
  return [...canonical, ...localized]
}

// "bolditalic" deliberately has no atomic alternative: a BoldItalic face must yield
// ['bold', 'italic'] so it can score against each want separately (a keep-original
// bold toggle on an italic run wants both tokens; an atomic token would score zero).
// "oblique" folds into "italic": families like Helvetica name their slanted faces
// Oblique, and an italic want must still count them as hits.
// "demibold"/"demi" fold the same way and must precede "bold": otherwise 'demibold'
// contains 'bold' and a Demibold face tokenizes as Bold, tying it on a bold want.
export const styleTokens = (ps: string): string[] =>
  (
    norm(ps).match(
      /semibold|extrabold|demibold|demi|bold|italic|oblique|light|thin|medium|heavy|black|regular|w\d/g,
    ) ?? []
  ).map((t) => (t === 'oblique' ? 'italic' : t === 'demi' ? 'demibold' : t))

const REGULARISH = ['regular', 'medium', 'w3', 'w4']
const BOLDISH = ['bold', 'semibold', 'demibold', 'extrabold']

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
  // A heavier-than-regular want with no exact face still prefers a neighbouring
  // weight (Demibold want -> Bold face) over falling all the way back to Regular.
  const nearWeight =
    want.some((t) => BOLDISH.includes(t) && !have.includes(t)) &&
    have.some((t) => BOLDISH.includes(t) && !want.includes(t))
      ? 1
      : 0
  const extras = have.filter(
    (t) => !wanted.includes(t) && !(nearWeight && BOLDISH.includes(t)),
  ).length
  const regularBonus = want.length === 0 && have.some((t) => REGULARISH.includes(t)) ? 1 : 0
  return hits * 4 + nearWeight * 2 - extras + regularBonus
}
