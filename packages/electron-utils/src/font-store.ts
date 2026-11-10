/** The downloadable font store, shared by every app that can fetch a family. */
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
  /** Injected, not imported: the name-table reader lives in the pptx-tied slides registry, and only hosts offering a local install have one. */
  readonly fontFileFamilies?: (path: string) => readonly string[]
}

export interface CatalogEntry {
  readonly family: string
  readonly script: CatalogFamily['script']
  readonly license: CatalogFamily['license']
  readonly installed: boolean
  /** total download size, so a UI can say what it is about to fetch */
  readonly bytes: number
}

/** The sfnt magic a real font file starts with; anything else is not one. */
const SFNT_MAGIC = new Set(['00010000', '4f54544f', '74746366', '74727565'])

function isPublished(family: CatalogFamily): boolean {
  return family.published !== false
}

export function familyDownloadBytes(family: CatalogFamily): number {
  return family.files.reduce((total, file) => total + file.bytes, 0)
}

/** Empty when the build ships no mirror: an affordance that cannot work is worse than none. */
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

/** Takes the catalog entry so a test can drive a synthetic family rather than re-pinning generated hashes. */
export function downloadFontFamily(env: FontStoreEnv, family: string): Promise<void> {
  const entry = FONT_CATALOG.find((f) => f.family === family)
  if (!entry || !isPublished(entry)) {
    return Promise.reject(new Error(`not in catalog: ${family}`))
  }
  return downloadCatalogEntry(env, entry)
}

/** Verifies every file against its pinned sha256 before writing, so a swapped artifact cannot be registered as a font. */
export function downloadCatalogEntry(env: FontStoreEnv, entry: CatalogFamily): Promise<void> {
  if (!env.cdnBaseUrl) return Promise.reject(new Error('font downloads are unavailable'))
  const existing = inFlight.get(entry.family)
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
  })().finally(() => inFlight.delete(entry.family))
  inFlight.set(entry.family, run)
  return run
}

/** Returns [] when the host injects no name-table reader, because it has no local install to offer. */
export function installLocalFontFiles(env: FontStoreEnv, paths: readonly string[]): string[] {
  const readFamilies = env.fontFileFamilies
  if (!readFamilies) return []
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
    const families = readFamilies(path)
    const primary = families[0]
    if (!primary) continue
    const ext =
      basename(path)
        .match(/\.(ttc|otc|otf)$/i)?.[1]
        ?.toLowerCase() ?? 'ttf'
    // Family-derived name is the registry index key; the suffix keeps one family's style files apart.
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
