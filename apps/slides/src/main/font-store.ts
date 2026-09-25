/**
 * Downloadable/installable font store (main process). Thin Electron wiring over
 * the shared pure core (@genoffice/electron-utils/font-store): the CDN URL comes
 * from packaged metadata or an env var, downloads land in <userData>/fonts, and
 * the FontRegistry scans that dir as a private one — the same measure-and-register
 * pipeline as Office DFonts, so a newly installed font immediately drives both
 * layout metrics and canvas drawing.
 */
import { join } from 'node:path'
import { app, net } from 'electron'
import type { OpenedPptx } from '@genoffice/pptx-engine'
import { familyAvailable, setUserFontDir } from './fonts'
import {
  downloadFontFamily as downloadCatalogFont,
  extractFontCdnBaseUrl,
  installLocalFontFiles as installFontFiles,
  listFontCatalog as listCatalog,
  resolveFontCdnBaseUrl,
  type CatalogEntry,
} from '@genoffice/electron-utils/font-store'
import { FONT_CATALOG } from './font-catalog'

export { extractFontCdnBaseUrl }

/** Resolve the build-injected font CDN URL for this app. */
export function fontCdnBaseUrl(): string | null {
  return resolveFontCdnBaseUrl({
    isPackaged: app.isPackaged,
    appPath: app.getAppPath(),
    envUrl: process.env.GENOFFICE_FONT_CDN_URL,
  })
}

export function fontStoreDir(): string {
  return join(app.getPath('userData'), 'fonts')
}

/** Wire the store dir into the font registry; call once at startup. */
export function initFontStore(): void {
  setUserFontDir(fontStoreDir())
}

export function listFontCatalog(): CatalogEntry[] {
  return listCatalog({ baseUrl: fontCdnBaseUrl(), installed: familyAvailable })
}

/** Download every style file of a catalog family into the store. */
export function downloadFontFamily(family: string): Promise<void> {
  return downloadCatalogFont(
    { baseUrl: fontCdnBaseUrl(), storeDir: fontStoreDir(), fetchImpl: (input) => net.fetch(input) },
    family,
  )
}

/** Copy user-picked font files into the store; returns the installed family names. */
export function installLocalFontFiles(paths: string[]): string[] {
  return installFontFiles(fontStoreDir(), paths)
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
