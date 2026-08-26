/**
 * Downloadable/installable font store (main process).
 *
 * A curated catalog of OFL-licensed families is mirrored on the GenOffice CDN
 * (versioned paths, sha256-pinned). Downloads and user-installed font files both
 * land in <userData>/fonts, which the FontRegistry scans as a private dir — the
 * same measure-and-register pipeline as Office DFonts, so a newly installed font
 * immediately drives both layout metrics and canvas drawing.
 */
import { createHash } from 'node:crypto'
import { mkdirSync, readFileSync, writeFileSync, existsSync, copyFileSync } from 'node:fs'
import { basename, join } from 'node:path'
import { app, net } from 'electron'
import type { OpenedPptx } from '@genoffice/pptx-engine'
import { familyAvailable, fontFileFamilies, setUserFontDir } from './fonts'
import { FONT_CATALOG, type CatalogFamily } from './font-catalog'

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
 * Official packages receive the URL through electron-builder extraMetadata.
 * Source/dev builds may opt in with an environment variable; without either,
 * all downloadable-font UI stays disabled while local font installation works.
 */
export function fontCdnBaseUrl(): string | null {
  if (!app.isPackaged) return normalizeFontCdnBaseUrl(process.env.GENOFFICE_FONT_CDN_URL)
  try {
    const pkg = JSON.parse(readFileSync(join(app.getAppPath(), 'package.json'), 'utf8')) as unknown
    return extractFontCdnBaseUrl(pkg)
  } catch {
    return null
  }
}

export function fontStoreDir(): string {
  return join(app.getPath('userData'), 'fonts')
}

/** Wire the store dir into the font registry; call once at startup. */
export function initFontStore(): void {
  setUserFontDir(fontStoreDir())
}

const downloading = new Map<string, Promise<void>>()

export interface FontCatalogEntry {
  family: string
  script: CatalogFamily['script']
  installed: boolean
  downloading: boolean
}

export function listFontCatalog(): FontCatalogEntry[] {
  if (!fontCdnBaseUrl()) return []
  return FONT_CATALOG.map((f) => ({
    family: f.family,
    script: f.script,
    installed: familyAvailable(f.family),
    downloading: downloading.has(f.family),
  }))
}

async function fetchVerified(url: string, sha256: string): Promise<Buffer> {
  const res = await net.fetch(url)
  if (!res.ok) throw new Error(`download failed: HTTP ${res.status}`)
  const buf = Buffer.from(await res.arrayBuffer())
  const got = createHash('sha256').update(buf).digest('hex')
  if (got !== sha256) throw new Error('download failed: checksum mismatch')
  return buf
}

/** Download every style file of a catalog family into the store. Throws on any failure;
 *  a concurrent call for the same family joins the in-flight download. */
export function downloadFontFamily(family: string): Promise<void> {
  const entry = FONT_CATALOG.find((f) => f.family === family)
  if (!entry) return Promise.reject(new Error(`not in catalog: ${family}`))
  const baseUrl = fontCdnBaseUrl()
  if (!baseUrl) return Promise.reject(new Error('font downloads are unavailable'))
  const inFlight = downloading.get(family)
  if (inFlight) return inFlight
  const run = (async () => {
    const dir = fontStoreDir()
    mkdirSync(dir, { recursive: true })
    for (const file of entry.files) {
      const dest = join(dir, file.file)
      if (existsSync(dest)) continue
      const url = new URL(encodeURIComponent(file.file), `${baseUrl}/`).toString()
      const buf = await fetchVerified(url, file.sha256)
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
export function installLocalFontFiles(paths: string[]): string[] {
  const dir = fontStoreDir()
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

/** Deck-referenced families that are missing locally but present in the catalog. */
export function missingCatalogFonts(opened: OpenedPptx): string[] {
  if (!fontCdnBaseUrl()) return []
  const wanted = new Set<string>()
  type TextLike = { paragraphs?: Array<{ runs?: Array<{ fontFamily?: string }> }> }
  const collectText = (text: TextLike | undefined): void => {
    for (const p of text?.paragraphs ?? [])
      for (const r of p.runs ?? []) if (r.fontFamily) wanted.add(r.fontFamily)
  }
  const walk = (els: unknown[]): void => {
    for (const el of els as Array<{
      type?: string
      children?: unknown[]
      text?: TextLike
      rows?: Array<Array<{ text?: TextLike }>>
    }>) {
      if (!el || typeof el !== 'object') continue
      if (el.children) walk(el.children)
      collectText(el.text)
      for (const row of el.rows ?? []) for (const cell of row) collectText(cell.text)
    }
  }
  for (const s of opened.deck.slides) walk(s.elements as unknown[])
  const inCatalog = new Set(FONT_CATALOG.map((f) => f.family))
  return [...wanted].filter((f) => inCatalog.has(f) && !familyAvailable(f)).sort()
}
