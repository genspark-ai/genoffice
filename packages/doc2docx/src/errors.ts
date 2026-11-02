/**
 * Error model for the legacy .doc reader. Every failure the parser can detect
 * is a LegacyDocError with a stable code, so callers (the import UI in later
 * stages) can react without parsing messages, and a malformed file can never
 * surface as a raw RangeError/TypeError.
 */
export const LEGACY_DOC_ERROR_CODES = [
  'not_cfb',
  'truncated',
  'bad_fib',
  'unsupported_version',
  'encrypted',
  'missing_stream',
  'bad_clx',
  'bad_piece_table',
  'limits_exceeded',
  'corrupt',
  'bad_fkp',
  'bad_sprm',
  'bad_stsh',
  'bad_section',
  'bad_hdd',
  'bad_table',
] as const

export type LegacyDocErrorCode = (typeof LEGACY_DOC_ERROR_CODES)[number]

/** What the user can do about each failure; appended to every message. */
const HINTS: Record<LegacyDocErrorCode, string> = {
  not_cfb:
    'the file is not an OLE2 compound document; a .docx renamed to .doc is imported by the Office Open XML reader instead',
  truncated:
    'the file is incomplete (truncated download or copy); get an intact copy of the .doc file',
  bad_fib: 'the Word header is invalid; open the file in Word or LibreOffice and save it again',
  unsupported_version:
    'only Word 97-2003 documents are supported; open the file in Word or LibreOffice and save it as .docx',
  encrypted:
    'the document is password-protected; remove the password in Word or LibreOffice before importing it',
  missing_stream:
    'the document is missing a stream every Word file needs; open it in Word or LibreOffice and save it again',
  bad_clx:
    'the document text table is invalid; open the file in Word or LibreOffice and save it again',
  bad_piece_table:
    'the document text table points outside the file; open the file in Word or LibreOffice and save it again',
  limits_exceeded:
    'the document is larger than the legacy import limit; split it or convert it to .docx in Word or LibreOffice',
  corrupt: 'the document is corrupt; open it in Word or LibreOffice and save it again',
  bad_fkp:
    'the document formatting table is invalid; open the file in Word or LibreOffice and save it again',
  bad_sprm:
    'the document formatting records are invalid; open the file in Word or LibreOffice and save it again',
  bad_stsh:
    'the document stylesheet is invalid; open the file in Word or LibreOffice and save it again',
  bad_section:
    'the document section table is invalid; open the file in Word or LibreOffice and save it again',
  bad_hdd:
    'the document header/footer table is invalid; open the file in Word or LibreOffice and save it again',
  bad_table:
    'the document table properties are invalid; open the file in Word or LibreOffice and save it again',
}

/** A legacy .doc that cannot be read, with a machine-readable reason code. */
export class LegacyDocError extends Error {
  readonly code: LegacyDocErrorCode

  constructor(code: LegacyDocErrorCode, detail: string) {
    super(`Cannot import this .doc file (${code}): ${detail}. ${HINTS[code]}.`)
    this.name = 'LegacyDocError'
    this.code = code
  }
}

export function isLegacyDocError(error: unknown): error is LegacyDocError {
  return error instanceof LegacyDocError
}
