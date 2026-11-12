/// Settings → General "Default PDF zoom": how a PDF without a remembered
/// reading position opens. Stored under `pdfDefaultZoom` in
/// userData/app-settings.json; the PDF renderer reads it on every open.
import { readFileSync } from 'node:fs'
import { writeJsonAtomic } from './atomic-write'

export const PDF_DEFAULT_ZOOM_KEY = 'pdfDefaultZoom'

export const PDF_DEFAULT_ZOOMS = ['width', 'page', 'last'] as const
export type PdfDefaultZoom = (typeof PDF_DEFAULT_ZOOMS)[number]

function readSettings(settingsPath: string): Record<string, unknown> {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath, 'utf8'))
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) return raw as Record<string, unknown>
  } catch {
    /* missing or corrupt: default */
  }
  return {}
}

export function sanitizePdfDefaultZoom(value: unknown): PdfDefaultZoom {
  return (PDF_DEFAULT_ZOOMS as readonly unknown[]).includes(value)
    ? (value as PdfDefaultZoom)
    : 'width'
}

export function readPdfDefaultZoom(settingsPath: string): PdfDefaultZoom {
  return sanitizePdfDefaultZoom(readSettings(settingsPath)[PDF_DEFAULT_ZOOM_KEY])
}

export function writePdfDefaultZoom(settingsPath: string, zoom: PdfDefaultZoom): void {
  const settings = readSettings(settingsPath)
  if (zoom === 'width') delete settings[PDF_DEFAULT_ZOOM_KEY]
  else settings[PDF_DEFAULT_ZOOM_KEY] = zoom
  writeJsonAtomic(settingsPath, settings)
}
