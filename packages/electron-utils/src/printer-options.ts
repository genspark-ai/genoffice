/// Driver option groups as CUPS reports them: `lpoptions -p <queue> -l` lines
/// read `Name/Translation: value1 value2 *default`. The print dialog shows
/// the groups the driver actually advertises (quality, color, media type —
/// the job options the dialogs already own — copies, sides — are not asked
/// for again), and the chosen values ride the job to lp as `-o name=value`.
/// Names and values come back as spawn arguments, so both are charset-guarded
/// here and in every consumer.

import { execFile } from 'node:child_process'

export interface PrinterOption {
  name: string
  /** the driver's own display name (locale data, not a UI string) */
  translation: string
  values: string[]
  defaultValue?: string
}

const OPTION_NAME_RE = /^[A-Za-z0-9_.+-]{1,64}$/
const OPTION_VALUE_RE = /^[A-Za-z0-9_.+-]{1,64}$/

/** the groups the print dialog knows how to present, in presentation order */
export const DRIVER_OPTION_NAMES: readonly string[] = [
  'cupsPrintQuality',
  'ColorModel',
  'MediaType',
]

export function isValidOptionName(value: unknown): value is string {
  return typeof value === 'string' && OPTION_NAME_RE.test(value)
}

export function isValidOptionValue(value: unknown): value is string {
  return typeof value === 'string' && OPTION_VALUE_RE.test(value)
}

export function parseLpOptions(raw: string): PrinterOption[] {
  const out: PrinterOption[] = []
  for (const line of raw.split('\n')) {
    const m = /^([A-Za-z0-9_.+-]+)\/([^:\n]{1,128}):\s*(.+)$/.exec(line.trim())
    if (!m || m[1] === undefined || m[2] === undefined || m[3] === undefined) continue
    const name = m[1]
    const translation = m[2]
    const values: string[] = []
    let defaultValue: string | undefined
    for (const token of m[3].split(/\s+/)) {
      const isDefault = token.startsWith('*')
      const value = isDefault ? token.slice(1) : token
      if (!OPTION_VALUE_RE.test(value)) continue
      if (isDefault) defaultValue = value
      if (!values.includes(value)) values.push(value)
    }
    if (!OPTION_NAME_RE.test(name) || values.length === 0) continue
    out.push({
      name,
      translation: translation.trim(),
      values,
      ...(defaultValue ? { defaultValue } : {}),
    })
  }
  return out
}

/** Queries one queue's option groups; any failure (unknown queue, no
 * lpoptions) resolves empty — the dialog then simply shows nothing extra. */
export async function queryPrinterOptions(
  queue: string,
  execFileFn: typeof execFile = execFile,
): Promise<PrinterOption[]> {
  return new Promise((resolve) => {
    execFileFn('/usr/bin/lpoptions', ['-p', queue, '-l'], (err, stdout) => {
      if (err) resolve([])
      else resolve(parseLpOptions(String(stdout)))
    })
  })
}
