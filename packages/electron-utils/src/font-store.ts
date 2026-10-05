/**
 * The downloadable font store, shared by every app that can fetch a family.
 *
 * ## Why this is not in apps/slides
 *
 * The catalog is data and the store is the plumbing around it; neither belongs
 * to the app that happened to need them first. Keeping them here is what lets
 * docs and sheets offer the same families without each re-deriving the CDN
 * rules, the checksum discipline and the licence bookkeeping.
 *
 * ## Why the environment arrives as an argument
 *
 * `electron-utils` carries "no Electron dependency, pure TS" in its
 * description, and the store needs `app.getPath` and `net.fetch`. Passing them
 * in keeps that true — and makes the whole thing testable without spawning an
 * Electron process, which is the point of the extraction.
 *
 * ## The rule the callers must honour
 *
 * A family is tens of megabytes — Noto Serif SC alone is 28 MiB. Nothing here
 * decides *when* to download; `listCatalog` hands the caller the byte count so
 * a UI can ask first. Downloading on a bare selection is this store's caller's
 * decision to get wrong, not the store's to make.
 */
import { createHash } from 'node:crypto'
import { existsSync, readFileSync, writeFileSync, mkdirSync, copyFileSync } from 'node:fs'
import { basename, join } from 'node:path'

import { FONT_CATALOG, type CatalogFamily } from './font-catalog'

export type { CatalogFamily, CatalogFile, FontLicense } from './font-catalog'

/** Everything the store needs from its host, so none of it imports electron. */
export interface FontStoreEnv {
  /** where downloaded and user-installed files live */
  readonly dir: string
  /** the mirror's base URL, or null when this build ships none */
  readonly cdnBaseUrl: string | null
  /** host fetch: resolve with the body, or `ok: false` for an HTTP error */
  readonly fetchBytes: (url: string) => Promise<{ ok: boolean; status: number; bytes: Uint8Array }>
  /** the host's own view of whether a family is already usable */
  readonly isFamilyAvailable: (family: string) => boolean
  /**
   * The family names inside a font file, read from its `name` table.
   *
   * Injected rather than imported: the implementation lives in the slides
   * font registry, which is tied to the pptx engine, and a store that guessed
   * from the filename would register the wrong index key.
   */
  readonly fontFileFamilies: (path: string) => readonly string[]
}

export interface CatalogEntry {
  readonly family: string
  readonly script: CatalogFamily['script']
  readonly license: CatalogFamily['license']
  /** true when the host already has this family */
  readonly installed: boolean
  /** total download size, so a UI can say what it is about to fetch */
  readonly bytes: number
}

/** The sfnt magic a real font file starts with; anything else is not one. */
const SFNT_MAGIC = new Set(['00010000', '4f54544f', '74746366', '74727565'])

function isPublished(family: CatalogFamily): boolean {
  return family.published !== false
}

/** What fetching this family costs, in bytes. */
export function familyDownloadBytes(family: CatalogFamily): number {
  return family.files.reduce((total, file) => total + file.bytes, 0)
}

/**
 * The rows an app may offer, and what each would cost.
 *
 * Empty when the build ships no mirror: an app showing a download affordance
 * that cannot work is worse than showing none.
 */
export function listCatalog(env: FontStoreEnv): CatalogEntry[] {
  if (!env.cdnBaseUrl) return []
  return FONT_CATALOG.filter(isPublished).map((family) => ({
    family: family.family,
    script: family.script,
    license: family.license,
    installed: env.isFamilyAvailable(family.family),
    bytes: familyDownloadBytes(family),
  }))
}

/** A download already running, so two callers asking join one request. */
const inFlight = new Map<string, Promise<void>>()

async function fetchVerified(env: FontStoreEnv, url: string, sha256: string): Promise<Uint8Array> {
  const response = await env.fetchBytes(url)
  if (!response.ok) throw new Error(`download failed: HTTP ${response.status}`)
  const got = createHash('sha256').update(response.bytes).digest('hex')
  if (got !== sha256) throw new Error('download failed: checksum mismatch')
  return response.bytes
}

/**
 * Fetch every style of a family into the store.
 *
 * Verifies every file against the pinned sha256 before it is written, so a
 * truncated or swapped artifact cannot end up registered as a font.
 */
export function downloadFontFamily(env: FontStoreEnv, family: string): Promise<void> {
  const entry = FONT_CATALOG.find((f) => f.family === family)
  if (!entry || !isPublished(entry)) {
    return Promise.reject(new Error(`not in catalog: ${family}`))
  }
  if (!env.cdnBaseUrl) return Promise.reject(new Error('font downloads are unavailable'))
  const existing = inFlight.get(family)
  if (existing) return existing
  const run = (async () => {
    mkdirSync(env.dir, { recursive: true })
    for (const file of entry.files) {
      const dest = join(env.dir, file.file)
      if (existsSync(dest)) continue
      const url = new URL(encodeURIComponent(file.file), `${env.cdnBaseUrl}/`).toString()
      const bytes = await fetchVerified(env, url, file.sha256)
      writeFileSync(dest, bytes)
    }
  })().finally(() => inFlight.delete(family))
  inFlight.set(family, run)
  return run
}

/** True when the file is already fetched, whatever the host calls "installed". */
export function familyDownloaded(env: FontStoreEnv, family: string): boolean {
  const entry = FONT_CATALOG.find((f) => f.family === family)
  if (!entry) return false
  return entry.files.every((file) => existsSync(join(env.dir, file.file)))
}

/**
 * Copy user-picked font files into the store, renamed to their family so the
 * filename-keyed index can find them. Returns the families that landed.
 */
export function installLocalFontFiles(env: FontStoreEnv, paths: readonly string[]): string[] {
  mkdirSync(env.dir, { recursive: true })
  const installed: string[] = []
  for (const path of paths) {
    let head: string
    try {
      head = readFileSync(path).subarray(0, 4).toString('hex')
    } catch {
      continue
    }
    if (!SFNT_MAGIC.has(head)) continue
    const families = env.fontFileFamilies(path)
    const primary = families[0]
    if (!primary) continue
    const ext =
      basename(path)
        .match(/\.(ttc|otc|otf)$/i)?.[1]
        ?.toLowerCase() ?? 'ttf'
    // Family-derived name = registry index key; the suffix keeps distinct
    // style files of one family apart.
    const styleTag = /bold\s*italic/i.test(basename(path))
      ? '-BoldItalic'
      : /bold/i.test(basename(path))
        ? '-Bold'
        : /italic|oblique/i.test(basename(path))
          ? '-Italic'
          : ''
    const dest = join(env.dir, `${primary.replace(/[\\/:]/g, '')}${styleTag}.${ext}`)
    try {
      copyFileSync(path, dest)
      installed.push(...families)
    } catch {
      /* unreadable/locked source: skip */
    }
  }
  return [...new Set(installed)]
}

/** Normalise a mirror URL the way the store has always accepted one. */
export function normalizeCdnBaseUrl(value: unknown): string | null {
  if (typeof value !== 'string' || !value.trim()) return null
  try {
    const url = new URL(value.trim())
    if (url.protocol !== 'https:' || url.username || url.password || url.search || url.hash) {
      return null
    }
    return `${url.origin}${url.pathname.replace(/\/+$/, '')}`
  } catch {
    return null
  }
}
