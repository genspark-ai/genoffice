/**
 * Downloadable/installable font store for the sheets main process. Thin Electron
 * wiring over the shared pure core (@genoffice/electron-utils/font-store):
 * downloads and user-picked font files land in <userData>/fonts; the renderer
 * fetches per-face sfnt bytes over IPC and registers them as FontFaces, which
 * canvas cell/shape drawing and the DOM dialogs both pick up automatically.
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

/** Resolve the build-injected font CDN URL for this app. */
export function sheetsFontCdnBaseUrl(): string | null {
  return resolveFontCdnBaseUrl({
    isPackaged: app.isPackaged,
    appPath: app.getAppPath(),
    envUrl: process.env.GENOFFICE_FONT_CDN_URL,
  })
}

export function sheetsFontStoreDir(): string {
  return join(app.getPath('userData'), 'fonts')
}

/** Families that exist in the store dir (per-face names), cached per read. */
let familyCache: Set<string> | null = null
function storeFamilies(): Set<string> {
  familyCache ??= new Set(storeFontFaces(sheetsFontStoreDir()).map((f) => f.family))
  return familyCache
}
function invalidateFamilyCache(): void {
  familyCache = null
}

export function listSheetsFontCatalog(): CatalogEntry[] {
  return listCatalog({ baseUrl: sheetsFontCdnBaseUrl(), installed: (f) => storeFamilies().has(f) })
}

export async function downloadSheetsFont(family: string): Promise<void> {
  await downloadCatalogFont(
    {
      baseUrl: sheetsFontCdnBaseUrl(),
      storeDir: sheetsFontStoreDir(),
      fetchImpl: (input) => net.fetch(input),
    },
    family,
  )
  invalidateFamilyCache()
}

export async function installLocalSheetsFonts(
  parent: BrowserWindow | null | undefined,
): Promise<string[]> {
  const r = await showOpenDialogWithMemory(dialog, parent, {
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Fonts', extensions: ['ttf', 'otf', 'ttc', 'otc'] }],
  })
  if (r.canceled || !r.filePaths.length) return []
  const families = installFontFiles(sheetsFontStoreDir(), r.filePaths)
  invalidateFamilyCache()
  return families
}

export function sheetsFontStoreFaces(): StoreFontFace[] {
  return storeFontFaces(sheetsFontStoreDir())
}

export function sheetsFontData(file: string, faceOffset: number): ArrayBuffer {
  return readStoreFontFace(sheetsFontStoreDir(), file, faceOffset)
}
