/**
 * Error model for the legacy .xls reader. Every failure the parser can detect
 * is a LegacyXlsError with a stable code, so callers (the import UI in later
 * stages) can react without parsing messages, and a malformed file can never
 * surface as a raw RangeError/TypeError.
 */
export const LEGACY_XLS_ERROR_CODES = [
  'not_cfb',
  'truncated',
  'missing_stream',
  'unsupported_version',
  'encrypted',
  'limits_exceeded',
  'corrupt',
  'bad_sst',
  'bad_boundsheet',
] as const

export type LegacyXlsErrorCode = (typeof LEGACY_XLS_ERROR_CODES)[number]

/** What the user can do about each failure; appended to every message. */
const ADVICE: Record<LegacyXlsErrorCode, string> = {
  not_cfb: 'The file does not look like an Excel 97-2003 workbook.',
  truncated: 'The file is shorter than its own structure claims.',
  missing_stream: 'The workbook stream is missing from the file.',
  unsupported_version: 'Only Excel 97-2003 (.xls) workbooks are supported.',
  encrypted: 'The workbook is password-protected; remove the password first.',
  limits_exceeded: 'The workbook is too large to convert.',
  corrupt: 'The workbook structure is damaged beyond parsing.',
  bad_sst: 'The shared-string table is damaged.',
  bad_boundsheet: 'The sheet directory is damaged.',
}

export class LegacyXlsError extends Error {
  readonly code: LegacyXlsErrorCode

  constructor(code: LegacyXlsErrorCode, detail?: string) {
    const advice = ADVICE[code]
    super(detail ? `${advice} (${detail})` : advice)
    this.name = 'LegacyXlsError'
    this.code = code
  }
}
