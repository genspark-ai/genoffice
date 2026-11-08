import { readFileSync } from 'node:fs'
import {
  DEFAULT_PDF_ZOOM,
  PDF_DEFAULT_ZOOM_KEY,
  normalizePdfDefaultZoom,
} from '../shared/default-zoom'
import type { PdfDefaultZoom } from '../shared/default-zoom'

/** Read on each open so a settings change takes effect in the next PDF. */
export function readPdfDefaultZoom(settingsPath: string): PdfDefaultZoom {
  try {
    const settings: unknown = JSON.parse(readFileSync(settingsPath, 'utf8'))
    if (settings && typeof settings === 'object' && !Array.isArray(settings)) {
      return normalizePdfDefaultZoom((settings as Record<string, unknown>)[PDF_DEFAULT_ZOOM_KEY])
    }
  } catch {
    // A missing or corrupt settings file keeps the existing opening behavior.
  }
  return DEFAULT_PDF_ZOOM
}
