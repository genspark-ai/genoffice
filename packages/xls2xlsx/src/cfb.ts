import { CompoundFileError, openCompoundStreams } from '@genoffice/compound-file'
import { LegacyXlsError } from './errors'

/** Caps mirror the doc2docx reader: a malformed container must fail fast, not
    hand the parser gigabytes to walk. */
const MAX_INPUT_BYTES = 256 * 1024 * 1024
const MAX_STREAM_BYTES = 128 * 1024 * 1024

/** The streams a BIFF8 workbook keeps its records in. Excel 97 wrote "Workbook";
    some producers ("Book") and Mac writers use the older spelling. */
const WORKBOOK_STREAMS = ['Workbook', 'Book']

export interface XlsContainer {
  /** the Workbook/Book stream bytes: one record stream, globals first */
  workbook: Uint8Array
}

/**
 * Open a legacy .xls container. The OLE2 rules — signature, 512-byte header,
 * input/stream caps and owned stream bytes — come from @genoffice/compound-file,
 * the same opener the .doc reader uses, instead of a second weaker copy.
 */
export function openXlsContainer(bytes: Uint8Array): XlsContainer {
  let opened
  try {
    opened = openCompoundStreams(bytes, {
      maxBytes: MAX_INPUT_BYTES,
      maxStreamBytes: MAX_STREAM_BYTES,
    })
  } catch (error) {
    if (error instanceof CompoundFileError) {
      throw new LegacyXlsError(error.code, error.detail)
    }
    throw error
  }
  for (const name of WORKBOOK_STREAMS) {
    const content = opened.stream(name)
    if (content && content.length > 0) return { workbook: content }
  }
  throw new LegacyXlsError('missing_stream')
}
