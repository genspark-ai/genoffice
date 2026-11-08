/** Opening preference; last-used preserves the existing per-file view history. */
export type PdfDefaultZoom = 'width' | 'page' | 'last-used'

export const PDF_DEFAULT_ZOOM_KEY = 'pdfDefaultZoom'
export const DEFAULT_PDF_ZOOM: PdfDefaultZoom = 'last-used'

export function isPdfDefaultZoom(value: unknown): value is PdfDefaultZoom {
  return value === 'width' || value === 'page' || value === 'last-used'
}

export function normalizePdfDefaultZoom(value: unknown): PdfDefaultZoom {
  return isPdfDefaultZoom(value) ? value : DEFAULT_PDF_ZOOM
}
