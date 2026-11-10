/**
 * Pure validators for the PDF print/export IPC args (docs:print-pdf-buffer,
 * docs:print, docs:save-merged-pdf). Renderer-supplied numbers reach Chromium
 * printToPDF verbatim, so they are range-checked here; the handlers stay thin.
 * Electron-free so unit tests can import this module directly.
 */

/** printable page dimension in twips: 0.1in (Word's floor) .. 50in */
export function validPrintDim(v: unknown): v is number {
  return typeof v === 'number' && Number.isFinite(v) && v >= 144 && v <= 72000
}

/** print scale factor: 0.1 .. 5 (undefined = Chromium default) */
export function validPrintScale(scale: unknown): boolean {
  return (
    scale === undefined ||
    (typeof scale === 'number' && Number.isFinite(scale) && scale >= 0.1 && scale <= 5)
  )
}

/** combined geometry guard shared by the printToPDF IPC handlers (docs:export-pdf,
 * docs:print-pdf-buffer): page width/height in twips plus optional scale */
export function validPrintGeometry(
  pageWidthTwips: unknown,
  pageHeightTwips: unknown,
  scale?: unknown,
): boolean {
  return validPrintDim(pageWidthTwips) && validPrintDim(pageHeightTwips) && validPrintScale(scale)
}

/** finite positive scale for the printToPDF/print option objects (Infinity fails `> 0` checks) */
export function printScaleOption(scale: unknown): { scale: number } | Record<string, never> {
  return typeof scale === 'number' && Number.isFinite(scale) && scale > 0 && scale !== 1
    ? { scale }
    : {}
}

/** the document's own paper size for the CUPS print route: the renderer reads
 * it off the pagination preview's first sheet (twips), and printToPDF bakes
 * the pages at that size instead of Chromium's Letter default. Absent or
 * malformed = no pageSize, never a wrong size. */
export function validPrintPaper(
  paper: unknown,
): { widthTwips: number; heightTwips: number } | null {
  if (typeof paper !== 'object' || paper === null) return null
  const p = paper as Record<string, unknown>
  if (!validPrintDim(p.widthTwips) || !validPrintDim(p.heightTwips)) return null
  return { widthTwips: p.widthTwips, heightTwips: p.heightTwips }
}

export { printJobOptions } from '@genoffice/electron-utils/print-job'
export type { PrintDuplex, PrintJobOptions } from '@genoffice/electron-utils/print-job'
