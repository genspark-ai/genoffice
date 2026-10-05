/**
 * Byte-faithful round-trip for source-mode files.
 *
 * A `.txt` or `.json` file has no document model behind it, so the only thing
 * standing between the user and a corrupted file is not reinterpreting what was
 * read. Three things a text editor normally normalises away are therefore kept
 * exactly as found:
 *
 * - a UTF-8 BOM, which some Windows tools still require
 * - CRLF line endings, which re-saving as LF is a whole-file diff in git
 * - a missing trailing newline, which a well-behaved writer would append
 *
 * Everything else (tabs vs spaces, indentation, key order) is the user's own
 * formatting and is never touched.
 */

/** The parts of a source file that must survive an edit unchanged. */
export interface SourceTextFormat {
  /** the file opened with a UTF-8 BOM */
  bom: boolean
  /** the dominant line ending found in the file */
  eol: '\n' | '\r\n'
  /** the file ended with a line terminator */
  trailingNewline: boolean
}

const BOM = '\uFEFF'

/** Read the format markers out of raw file text, and return the text without them. */
export function readSourceText(raw: string): { text: string; format: SourceTextFormat } {
  const bom = raw.startsWith(BOM)
  const body = bom ? raw.slice(1) : raw
  const eol: '\n' | '\r\n' = body.includes('\r\n') ? '\r\n' : '\n'
  // a lone \r is a line terminator too (classic Mac files), so it normalises
  // to \n here and is rewritten in the file's own style on save
  const text = body.replace(/\r\n?/g, '\n')
  return {
    text,
    format: {
      bom,
      eol,
      trailingNewline: body === '' || body.endsWith('\n') || body.endsWith('\r'),
    },
  }
}

/** Put the format markers back around edited text so the saved file matches the original. */
export function writeSourceText(text: string, format: SourceTextFormat): string {
  const normalized = text.replace(/\r\n/g, '\n')
  const body = format.eol === '\r\n' ? normalized.replace(/\n/g, '\r\n') : normalized
  // a file that ended without a terminator must still end without one: adding
  // the newline is a whole-file diff, and some tools treat it as a real edit
  const withEol =
    body === '' || body.endsWith('\n')
      ? body
      : format.trailingNewline
        ? `${body}${format.eol}`
        : body
  return format.bom ? BOM + withEol : withEol
}
