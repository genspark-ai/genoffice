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

/**
 * The same paper for the webContents.print() route, in the unit Electron
 * takes (microns).
 *
 * Chromium's own default paper is Letter and webContents.print only falls
 * back to the printer's default when asked to (usePrinterDefaultPageSize
 * defaults to false), so a document with its own sheet has to state it: an A4
 * page printed on Letter is taller than the paper, so Chromium splits the
 * sheet in two — the printed page loses its bottom and the page order goes
 * with it (genoffice#2563). Absent or malformed = no pageSize, never a wrong
 * size.
 */
export function printPaperOption(
  paper: unknown,
): { pageSize: { width: number; height: number } } | Record<string, never> {
  const geometry = validPrintPaper(paper)
  if (!geometry) return {}
  return {
    pageSize: {
      // 1440 twips to the inch, 25400 microns to the inch
      width: Math.round((geometry.widthTwips * 25_400) / 1440),
      height: Math.round((geometry.heightTwips * 25_400) / 1440),
    },
  }
}

export { printJobOptions } from '@genoffice/electron-utils/print-job'
export type { PrintDuplex, PrintJobOptions } from '@genoffice/electron-utils/print-job'
