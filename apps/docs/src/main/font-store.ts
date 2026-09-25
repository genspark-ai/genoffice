/**
 * Downloadable/installable font store for the docs main process. Thin Electron
 * wiring over the shared pure core (@genoffice/electron-utils/font-store):
 * downloads and user-picked font files land in <userData>/fonts; the renderer
 * fetches per-face sfnt bytes over IPC and registers them as FontFaces, which
 * also makes the fonts visible to the open-time missing-font check.
 */
import { app, dialog, net } from 'electron'
import type { BrowserWindow } from 'electron'
import { join } from 'node:path'
import {
  downloadFontFamily as downloadCatalogFont,
  installLocalFontFiles as installFontFiles,
  listFontCatalog as listCatalog,
  readStoreFontFace,
  resolveFontCdnBaseUrl,
  storeFontFaces,
  type CatalogEntry,
  type StoreFontFace,
} from '@genoffice/electron-utils/font-store'
import { showOpenDialogWithMemory } from '@genoffice/electron-utils'
import { FONT_CATALOG } from '@genoffice/electron-utils/font-catalog'

/** Resolve the build-injected font CDN URL for this app. */
export function docsFontCdnBaseUrl(): string | null {
  return resolveFontCdnBaseUrl({
    isPackaged: app.isPackaged,
    appPath: app.getAppPath(),
    envUrl: process.env.GENOFFICE_FONT_CDN_URL,
  })
}

export function docsFontStoreDir(): string {
  return join(app.getPath('userData'), 'fonts')
}

/** Families that exist in the store dir (per-face names), cached per read. */
let familyCache: Set<string> | null = null
function storeFamilies(): Set<string> {
  familyCache ??= new Set(storeFontFaces(docsFontStoreDir()).map((f) => f.family))
  return familyCache
}
function invalidateFamilyCache(): void {
  familyCache = null
}

export function listDocsFontCatalog(): CatalogEntry[] {
  return listCatalog({ baseUrl: docsFontCdnBaseUrl(), installed: (f) => storeFamilies().has(f) })
}

/** Families in the catalog that the store does not have yet. */
export function missingCatalogFamilies(declared: Iterable<string>): string[] {
  if (!docsFontCdnBaseUrl()) return []
  const inCatalog = new Set(FONT_CATALOG.map((f) => f.family))
  const missing = new Set<string>()
  for (const name of declared) {
    if (inCatalog.has(name) && !storeFamilies().has(name)) missing.add(name)
  }
  return [...missing].sort()
}

export async function downloadDocsFont(family: string): Promise<void> {
  await downloadCatalogFont(
    {
      baseUrl: docsFontCdnBaseUrl(),
      storeDir: docsFontStoreDir(),
      fetchImpl: (input) => net.fetch(input),
    },
    family,
  )
  invalidateFamilyCache()
}

export async function installLocalDocsFonts(
  parent: BrowserWindow | null | undefined,
): Promise<string[]> {
  const r = await showOpenDialogWithMemory(dialog, parent, {
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Fonts', extensions: ['ttf', 'otf', 'ttc', 'otc'] }],
  })
  if (r.canceled || !r.filePaths.length) return []
  const families = installFontFiles(docsFontStoreDir(), r.filePaths)
  invalidateFamilyCache()
  return families
}

export function docsFontStoreFaces(): StoreFontFace[] {
  return storeFontFaces(docsFontStoreDir())
}

export function docsFontData(file: string, faceOffset: number): ArrayBuffer {
  return readStoreFontFace(docsFontStoreDir(), file, faceOffset)
}
