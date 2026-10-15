export type PrintDuplex = 'simplex' | 'longEdge' | 'shortEdge'

/** per-job options chosen in an in-app print dialog; deviceName null = system dialog */
export interface PrintJobOptions {
  deviceName?: string | null
  copies?: number
  collate?: boolean
  duplex?: PrintDuplex
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
