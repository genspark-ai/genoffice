/**
 * Chromium lays print jobs out at device scale 1, but pagination measured the
 * canvas at the screen's scale, and Blink rounds font ascent/descent to device
 * pixels: a line whose height comes from a taller secondary-font run is up to
 * half a pixel shorter in print than on a Retina screen. Across a document the
 * drift reaches several pixels and the fixed page clips cut through glyphs at
 * the seams. Zooming the sheets by the screen ratio and printing at the inverse
 * scale reproduces the screen rounding, so print geometry equals measurement.
 */

// Chromium accepts printToPDF scales in 0.1..2
const MAX_ZOOM = 10
const MIN_ZOOM = 0.5

export function printZoom(dpr: number = window.devicePixelRatio): number {
  if (!(dpr > 0) || Math.abs(dpr - 1) < 1e-3) return 1
  return Math.min(Math.max(dpr, MIN_ZOOM), MAX_ZOOM)
}

const previewRoot = () => document.querySelector<HTMLElement>('.pagination-preview')

/** Zooms the preview sheets for print (print media only) and returns the job's print scale. */
export function setPrintZoom(): number {
  const zoom = printZoom()
  if (zoom !== 1) previewRoot()?.style.setProperty('--pv-print-zoom', String(zoom))
  return 1 / zoom
}

export function clearPrintZoom(): void {
  previewRoot()?.style.removeProperty('--pv-print-zoom')
}
