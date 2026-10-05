/**
 * The download/install half of the font store, wired to Electron.
 *
 * The catalog and the store logic live in `@genoffice/electron-utils/font-store`
 * so docs and sheets can offer the same families; this file is only the glue —
 * where the store dir is, which mirror this build ships, and how each app
 * answers "is this family already usable".
 */
import { join } from 'node:path'
import { readFileSync } from 'node:fs'
import { app, net } from 'electron'
import type { OpenedPptx } from '@genoffice/pptx-engine'
import {
  downloadFontFamily as downloadFromStore,
  familyDownloaded,
  installLocalFontFiles as installIntoStore,
  listCatalog,
  normalizeCdnBaseUrl,
  type FontStoreEnv,
} from '@genoffice/electron-utils/font-store'
import { FONT_CATALOG, type CatalogFamily } from '@genoffice/electron-utils/font-catalog'
import { familyAvailable, fontFileFamilies, setUserFontDir } from './fonts'

export { FONT_CATALOG, normalizeCdnBaseUrl }
export type { CatalogFamily }

/** Read the build-injected font CDN URL from packaged app metadata. */
export function extractFontCdnBaseUrl(pkg: unknown): string | null {
  if (!pkg || typeof pkg !== 'object') return null
  const raw = (pkg as Record<string, unknown>).genofficeFontCdn
  if (!raw || typeof raw !== 'object') return null
  return normalizeCdnBaseUrl((raw as Record<string, unknown>).baseUrl)
}

/**
 * Official packages receive the URL through electron-builder extraMetadata.
 * Source/dev builds may opt in with an environment variable; without either,
 * all downloadable-font UI stays disabled while local font installation works.
 */
export function fontCdnBaseUrl(): string | null {
  if (!app.isPackaged) return normalizeCdnBaseUrl(process.env.GENOFFICE_FONT_CDN_URL)
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

/** The store, described to the shared module in this app's terms. */
function storeEnv(): FontStoreEnv {
  return {
    dir: fontStoreDir(),
    cdnBaseUrl: fontCdnBaseUrl(),
    fetchBytes: async (url) => {
      const response = await net.fetch(url)
      if (!response.ok) return { ok: false, status: response.status, bytes: new Uint8Array() }
      return {
        ok: true,
        status: response.status,
        bytes: new Uint8Array(await response.arrayBuffer()),
      }
    },
    isFamilyAvailable: familyAvailable,
    fontFileFamilies,
  }
}

export interface FontCatalogEntry {
  family: string
  script: CatalogFamily['script']
  license: CatalogFamily['license']
  installed: boolean
  /** total download size; the ribbon shows it so a pick never hides a 28 MiB fetch */
  bytes: number
}

/** Rows whose files are live on the CDN: the only ones the pickers may offer. */
export function listFontCatalog(): FontCatalogEntry[] {
  return listCatalog(storeEnv())
}

export function downloadFontFamily(family: string): Promise<void> {
  return downloadFromStore(storeEnv(), family)
}

export function installLocalFontFiles(paths: string[]): string[] {
  return installIntoStore(storeEnv(), paths)
}

/** True when every file of a catalog family is already in the store. */
export function isFamilyDownloaded(family: string): boolean {
  return familyDownloaded(storeEnv(), family)
}

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
