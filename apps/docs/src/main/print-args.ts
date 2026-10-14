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

export type PrintDuplex = 'simplex' | 'longEdge' | 'shortEdge'

/** per-job options chosen in the in-app print dialog; deviceName null = system dialog */
export interface PrintJobOptions {
  deviceName?: string | null
  copies?: number
  collate?: boolean
  duplex?: PrintDuplex
}

const DUPLEX_MODES: ReadonlySet<string> = new Set(['simplex', 'longEdge', 'shortEdge'])

/** webContents.print fields for the dialog's job options. A named printer ('' =
 * OS default) prints silently; null/omitted keeps the system dialog, which then
 * owns copies/duplex itself. Out-of-range values are dropped, not clamped. */
export function printJobOptions(raw: unknown): {
  silent: boolean
  deviceName?: string
  copies?: number
  collate?: boolean
  duplexMode?: PrintDuplex
} {
  if (typeof raw !== 'object' || raw === null) return { silent: false }
  const job = raw as Record<string, unknown>
  const silent = typeof job.deviceName === 'string'
  const out: ReturnType<typeof printJobOptions> = { silent }
  if (!silent) return out
  if (job.deviceName !== '') out.deviceName = job.deviceName as string
  if (
    typeof job.copies === 'number' &&
    Number.isInteger(job.copies) &&
    job.copies >= 1 &&
    job.copies <= 999
  ) {
    out.copies = job.copies
  }
  if (typeof job.collate === 'boolean') out.collate = job.collate
  if (typeof job.duplex === 'string' && DUPLEX_MODES.has(job.duplex)) {
    out.duplexMode = job.duplex as PrintDuplex
  }
  return out
}
