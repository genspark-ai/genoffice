/**
 * Error model for the legacy .ppt reader. Every failure the parser can detect
 * is a LegacyPptError with a stable code, so callers (the import UI in later
 * stages) can react without parsing messages, and a malformed file can never
 * surface as a raw RangeError/TypeError.
 */
export const LEGACY_PPT_ERROR_CODES = [
  'not_cfb',
  'truncated',
  'missing_stream',
  'unsupported_version',
  'encrypted',
  'limits_exceeded',
  'corrupt',
  'bad_persist',
  'bad_slide',
] as const

export type LegacyPptErrorCode = (typeof LEGACY_PPT_ERROR_CODES)[number]

/** What the user can do about each failure; appended to every message. */
const ADVICE: Record<LegacyPptErrorCode, string> = {
  not_cfb: 'The file does not look like a PowerPoint 97-2003 presentation.',
  truncated: 'The file is shorter than its own structure claims.',
  missing_stream: 'The presentation stream is missing from the file.',
  unsupported_version: 'Only PowerPoint 97-2003 (.ppt) files are supported.',
  encrypted:
    'The presentation is password-protected; remove the password in PowerPoint before converting it.',
  limits_exceeded: 'The presentation is too large to convert.',
  corrupt: 'The presentation structure is damaged beyond parsing.',
  bad_persist: 'The presentation change directory is damaged.',
  bad_slide: 'A slide container is damaged.',
}

export class LegacyPptError extends Error {
  readonly code: LegacyPptErrorCode

  constructor(code: LegacyPptErrorCode, detail?: string) {
    const advice = ADVICE[code]
    super(detail ? `${advice} (${detail})` : advice)
    this.name = 'LegacyPptError'
    this.code = code
  }
}
