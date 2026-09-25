/**
 * Downloadable/installable font store, shared by every app's main process.
 *
 * A curated catalog of OFL-licensed families (./font-catalog) is mirrored on the
 * GenOffice CDN (versioned paths, sha256-pinned). Downloads and user-installed
 * font files land in <userData>/fonts; each app wires that dir into its own
 * rendering path (slides' FontRegistry measures and registers for canvas, docs
 * and sheets serve the bytes to the renderer as FontFaces). Pure TypeScript with
 * no Electron import: the app injects its store dir, CDN URL and fetch impl.
 */
import { createHash } from 'node:crypto'
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  writeFileSync,
} from 'node:fs'
import { basename, join, resolve, sep } from 'node:path'
import { FONT_CATALOG, type FontScript } from './font-catalog'

/** Normalize: NFKC (full-width MS -> MS), lowercase, strip spaces/hyphens/underscores. */
function norm(s: string): string {
  return s
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[\s\-_]/g, '')
}

/** Metadata for one face in a ttc/ttf (name table), used to pick a face by requested family/style. */
export interface FaceInfo {
  /** Position of the offset table within the file (0 for non-ttc) */
  offset: number
  /** Family name for drawing: prefer the ASCII English name (resolvable by CSS by name) */
  display: string
  /** Normalized set of family names (name 1/16, including localized names) */
  famKeys: string[]
  /** Family + subfamily concatenation (normalized), used for style picks like bold/W6 */
  styleText: string
  /** OS/2 usWeightClass (name-independent weight evidence for ranking) */
  weight?: number
}

function readNameStrings(
  buf: Buffer,
  nameOff: number,
): { families: string[]; subfamilies: string[] } {
  const families: string[] = []
  const subfamilies: string[] = []
  const count = buf.readUInt16BE(nameOff + 2)
  const strBase = nameOff + buf.readUInt16BE(nameOff + 4)
  for (let i = 0; i < count; i++) {
    const r = nameOff + 6 + 12 * i
    const platform = buf.readUInt16BE(r)
    const encoding = buf.readUInt16BE(r + 2)
    const nameId = buf.readUInt16BE(r + 6)
    if (nameId !== 1 && nameId !== 2 && nameId !== 16 && nameId !== 17) continue
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
      continue // Mac-platform non-Roman encodings (legacy Korean/Chinese codepages) cannot be decoded; skip
    }
    if (!s) continue
    const list = nameId === 1 || nameId === 16 ? families : subfamilies
    if (!list.includes(s)) list.push(s)
  }
  return { families, subfamilies }
}

/** Parse every face of a sfnt/ttc buffer: name-table families plus OS/2 weight. */
export function readFaceDir(buf: Buffer): FaceInfo[] {
  const offsets =
    buf.toString('ascii', 0, 4) === 'ttcf'
      ? Array.from({ length: buf.readUInt32BE(8) }, (_, i) => buf.readUInt32BE(12 + 4 * i))
      : [0]
  return offsets.map((offset) => {
    let families: string[] = []
    let subfamilies: string[] = []
    let weight: number | undefined
    try {
      const numTables = buf.readUInt16BE(offset + 4)
      for (let t = 0; t < numTables; t++) {
        const e = offset + 12 + 16 * t
        const tag = buf.toString('ascii', e, e + 4)
        if (tag === 'name') {
          ;({ families, subfamilies } = readNameStrings(buf, buf.readUInt32BE(e + 8)))
        } else if (tag === 'OS/2') {
          // usWeightClass: name records can be missing/undecodable (cloud numeric files),
          // so face ranking needs the weight straight from the table
          const w = buf.readUInt16BE(buf.readUInt32BE(e + 8) + 4)
          if (w >= 1 && w <= 1000) weight = w
        }
      }
    } catch {
      /* Even if the name table is unreadable, the first face can still be parsed */
    }
    const ascii = families.find((f) => /^[\x20-\x7e]+$/.test(f))
    return {
      offset,
      display: ascii ?? families[0] ?? '',
      famKeys: families.map(norm),
      styleText: norm([...families, ...subfamilies].join(' ')),
      ...(weight != null ? { weight } : {}),
    }
  })
}

/** Family names (name 1/16) declared by a local font file; [] when unreadable. */
export function fontFileFamilies(path: string): string[] {
  try {
    const faces = readFaceDir(readFileSync(path))
    return [...new Set(faces.map((f) => f.display).filter(Boolean))]
  } catch {
    return []
  }
}

/** Extract a single face from a ttc into a standalone sfnt (rewrite the table directory,
 *  copy table data by original offset); passthrough for plain sfnt files. `drop` lists layout
 *  tables (GSUB/GPOS/GDEF) to omit for metrics-only parsing when opentype.js rejects them. */
export function extractFace(buf: Buffer, offset: number, drop?: ReadonlySet<string>): ArrayBuffer {
  const isTtc = buf.toString('ascii', 0, 4) === 'ttcf'
  if (!isTtc && !drop) {
    return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer
  }
  const faceOff = isTtc ? offset : 0
  const numTables = buf.readUInt16BE(faceOff + 4)
  const entries: Array<{ dirPos: number; tOff: number; tLen: number; newOff: number }> = []
  for (let t = 0; t < numTables; t++) {
    const e = faceOff + 12 + 16 * t
    if (drop?.has(buf.toString('ascii', e, e + 4))) continue
    entries.push({
      dirPos: e,
      tOff: buf.readUInt32BE(e + 8),
      tLen: buf.readUInt32BE(e + 12),
      newOff: 0,
    })
  }
  let total = 12 + 16 * entries.length
  for (const e of entries) {
    e.newOff = total
    total += (e.tLen + 3) & ~3
  }
  const out = Buffer.alloc(total)
  buf.copy(out, 0, faceOff, faceOff + 4)
  out.writeUInt16BE(entries.length, 4)
  const pow = 2 ** Math.floor(Math.log2(entries.length || 1))
  out.writeUInt16BE(pow * 16, 6)
  out.writeUInt16BE(Math.log2(pow), 8)
  out.writeUInt16BE(entries.length * 16 - pow * 16, 10)
  for (let t = 0; t < entries.length; t++) {
    const e = entries[t]!
    buf.copy(out, 12 + 16 * t, e.dirPos, e.dirPos + 8)
    out.writeUInt32BE(e.newOff, 12 + 16 * t + 8)
    out.writeUInt32BE(e.tLen, 12 + 16 * t + 12)
    buf.copy(out, e.newOff, e.tOff, e.tOff + e.tLen)
  }
  return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer
}

function normalizeFontCdnBaseUrl(value: unknown): string | null {
  if (typeof value !== 'string' || !value.trim()) return null
  try {
    const url = new URL(value.trim())
    if (url.protocol !== 'https:' || url.username || url.password || url.search || url.hash) {
      return null
    }
    const path = url.pathname.replace(/\/+$/, '')
    return `${url.origin}${path}`
  } catch {
    return null
  }
}

/** Read the build-injected font CDN URL from packaged app metadata. */
export function extractFontCdnBaseUrl(pkg: unknown): string | null {
  if (!pkg || typeof pkg !== 'object') return null
  const raw = (pkg as Record<string, unknown>).genofficeFontCdn
  if (!raw || typeof raw !== 'object') return null
  return normalizeFontCdnBaseUrl((raw as Record<string, unknown>).baseUrl)
}

/**
 * Official packages receive the URL through electron-builder extraMetadata;
 * source/dev builds may opt in with an environment variable. Without either,
 * all downloadable-font UI stays disabled while local font installation works.
 */
export function resolveFontCdnBaseUrl(opts: {
  isPackaged: boolean
  appPath: string
  envUrl?: string
}): string | null {
  if (!opts.isPackaged) return normalizeFontCdnBaseUrl(opts.envUrl)
  try {
    return extractFontCdnBaseUrl(
      JSON.parse(readFileSync(join(opts.appPath, 'package.json'), 'utf8')),
    )
  } catch {
    return null
  }
}

export interface CatalogEntry {
  family: string
  script: FontScript
  installed: boolean
  downloading: boolean
}

const downloading = new Map<string, Promise<void>>()

export function listFontCatalog(opts: {
  baseUrl: string | null
  installed: (family: string) => boolean
}): CatalogEntry[] {
  if (!opts.baseUrl) return []
  return FONT_CATALOG.map((f) => ({
    family: f.family,
    script: f.script,
    installed: opts.installed(f.family),
    downloading: downloading.has(f.family),
  }))
}

async function fetchVerified(
  fetchImpl: (input: string) => Promise<Response>,
  url: string,
  sha256: string,
): Promise<Buffer> {
  const res = await fetchImpl(url)
  if (!res.ok) throw new Error(`download failed: HTTP ${res.status}`)
  const buf = Buffer.from(await res.arrayBuffer())
  const got = createHash('sha256').update(buf).digest('hex')
  if (got !== sha256) throw new Error('download failed: checksum mismatch')
  return buf
}

/** Download every style file of a catalog family into the store. Throws on any failure;
 *  a concurrent call for the same family joins the in-flight download. */
export function downloadFontFamily(
  opts: {
    baseUrl: string | null
    storeDir: string
    fetchImpl: (input: string) => Promise<Response>
  },
  family: string,
): Promise<void> {
  const entry = FONT_CATALOG.find((f) => f.family === family)
  if (!entry) return Promise.reject(new Error(`not in catalog: ${family}`))
  if (!opts.baseUrl) return Promise.reject(new Error('font downloads are unavailable'))
  const inFlight = downloading.get(family)
  if (inFlight) return inFlight
  const run = (async () => {
    const dir = opts.storeDir
    mkdirSync(dir, { recursive: true })
    for (const file of entry.files) {
      const dest = join(dir, file.file)
      if (existsSync(dest)) continue
      const url = new URL(encodeURIComponent(file.file), `${opts.baseUrl}/`).toString()
      const buf = await fetchVerified(opts.fetchImpl, url, file.sha256)
      writeFileSync(dest, buf)
    }
  })().finally(() => downloading.delete(family))
  downloading.set(family, run)
  return run
}

const SFNT_MAGIC = new Set(['00010000', '4f54544f', '74746366', '74727565']) // sfnt / OTTO / ttcf / true

/**
 * Copy user-picked font files into the store, renamed to their primary family
 * name so the filename-keyed registry index can find them. Returns the family
 * names that were installed.
 */
export function installLocalFontFiles(storeDir: string, paths: string[]): string[] {
  const dir = storeDir
  mkdirSync(dir, { recursive: true })
  const installed: string[] = []
  for (const p of paths) {
    let head: string
    try {
      head = readFileSync(p).subarray(0, 4).toString('hex')
    } catch {
      continue
    }
    if (!SFNT_MAGIC.has(head)) continue
    const families = fontFileFamilies(p)
    const primary = families[0]
    if (!primary) continue
    const ext =
      basename(p)
        .match(/\.(ttc|otc|otf)$/i)?.[1]
        ?.toLowerCase() ?? 'ttf'
    // Family-derived name = registry index key; suffix keeps distinct style files apart
    const styleTag = /bold\s*italic/i.test(basename(p))
      ? '-BoldItalic'
      : /bold/i.test(basename(p))
        ? '-Bold'
        : /italic|oblique/i.test(basename(p))
          ? '-Italic'
          : ''
    const dest = join(dir, `${primary.replace(/[\\/:]/g, '')}${styleTag}.${ext}`)
    try {
      copyFileSync(p, dest)
      installed.push(...families)
    } catch {
      /* unreadable/locked source: skip */
    }
  }
  return [...new Set(installed)]
}

/** One registrable face of a store font file: family plus FontFace descriptors. */
export interface StoreFontFace {
  file: string
  faceOffset: number
  family: string
  weight: number
  italic: boolean
}

/** Every registrable face in the store dir (per-face for ttc collections). */
export function storeFontFaces(storeDir: string): StoreFontFace[] {
  let files: string[]
  try {
    files = readdirSync(storeDir)
  } catch {
    return []
  }
  const faces: StoreFontFace[] = []
  for (const file of files) {
    if (!/\.(ttf|otf|ttc|otc)$/i.test(file)) continue
    let buf: Buffer
    try {
      buf = readFileSync(join(storeDir, file))
    } catch {
      continue
    }
    if (!SFNT_MAGIC.has(buf.subarray(0, 4).toString('hex'))) continue
    for (const face of readFaceDir(buf)) {
      if (!face.display) continue
      faces.push({
        file,
        faceOffset: face.offset,
        family: face.display,
        weight: face.weight ?? 400,
        italic: /italic|oblique/.test(face.styleText),
      })
    }
  }
  return faces
}

/** Standalone sfnt bytes of one store font face, for renderer FontFace registration.
 *  Rejects anything that is not a bare store filename (path traversal). */
export function readStoreFontFace(storeDir: string, file: string, faceOffset: number): ArrayBuffer {
  if (file !== basename(file) || file.includes('/') || file.includes('\\')) {
    throw new Error(`not a store font: ${file}`)
  }
  const dest = resolve(storeDir, file)
  if (!dest.startsWith(resolve(storeDir) + sep)) throw new Error(`not a store font: ${file}`)
  return extractFace(readFileSync(dest), faceOffset)
}
