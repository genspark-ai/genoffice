import { isValidOptionName, isValidOptionValue } from './printer-options'

export type PrintDuplex = 'simplex' | 'longEdge' | 'shortEdge'

/** per-job options chosen in an in-app print dialog; deviceName null = system dialog */
export interface PrintJobOptions {
  deviceName?: string | null
  copies?: number
  collate?: boolean
  duplex?: PrintDuplex
  /** the document's own paper in inches (the pdf renderer's measured first
   * page): the CUPS print route sizes printToPDF's pages with it instead of
   * Chromium's Letter default. Out of range = dropped, never clamped. */
  pageSizeIn?: { readonly width: number; readonly height: number }
  /** driver options picked from the queue's advertised groups (quality,
   * color, media type) — charset-guarded before they become lp arguments */
  driverOptions?: readonly { readonly name: string; readonly value: string }[]
}

export interface PrinterInfo {
  name: string
  displayName: string
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
  pageSizeIn?: { readonly width: number; readonly height: number }
  driverOptions?: { readonly name: string; readonly value: string }[]
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
  const size = job.pageSizeIn as Record<string, unknown> | undefined
  if (
    typeof size === 'object' &&
    size !== null &&
    typeof size.width === 'number' &&
    Number.isFinite(size.width) &&
    size.width > 0 &&
    size.width <= 200 &&
    typeof size.height === 'number' &&
    Number.isFinite(size.height) &&
    size.height > 0 &&
    size.height <= 200
  ) {
    out.pageSizeIn = { width: size.width, height: size.height }
  }
  const driver = job.driverOptions
  if (Array.isArray(driver) && driver.length > 0) {
    const clean = driver
      .slice(0, 16)
      .filter(
        (o): o is { name: string; value: string } =>
          typeof o === 'object' &&
          o !== null &&
          isValidOptionName((o as Record<string, unknown>).name) &&
          isValidOptionValue((o as Record<string, unknown>).value),
      )
    if (clean.length > 0) out.driverOptions = clean
  }
  return out
}

/** The IPC answer for one `webContents.print` callback: success passes, and
 * a cancel the user chose in the interactive dialog is not an error (the
 * caller keeps the dialog open). Anything else the print system reports is
 * an error — including "Print job canceled" on a silent job, where no
 * dialog exists to have canceled anything, so a driver-level failure must
 * surface instead of vanishing. */
export function printOutcome(
  success: boolean,
  failureReason: string | undefined,
  silent: boolean,
): { ok: boolean; error?: string } {
  if (success) return { ok: true }
  if (!silent && failureReason && /cancel/i.test(failureReason)) return { ok: false }
  return { ok: false, ...(failureReason ? { error: failureReason } : {}) }
}
