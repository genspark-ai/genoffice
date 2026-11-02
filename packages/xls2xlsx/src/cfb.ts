import * as CFB from 'cfb'
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

/** Open a legacy .xls container: OLE2/CFB validated, the workbook stream out. */
export function openXlsContainer(bytes: Uint8Array): XlsContainer {
  if (bytes.length > MAX_INPUT_BYTES) {
    throw new LegacyXlsError('limits_exceeded', `input ${bytes.length} bytes`)
  }
  if (bytes.length < 8 || bytes[0] !== 0xd0 || bytes[1] !== 0xcf) {
    throw new LegacyXlsError('not_cfb')
  }
  let cfb: CFB.CFB$Container
  try {
    cfb = CFB.read(bytes as unknown as Buffer, { type: 'buffer' })
  } catch (err) {
    throw new LegacyXlsError('not_cfb', err instanceof Error ? err.message : undefined)
  }
  for (const name of WORKBOOK_STREAMS) {
    const entry = CFB.find(cfb, name)
    if (!entry) continue
    // cfb hands the stream back as a Buffer for big streams and a plain
    // number array for small ones — both are array-likes of bytes
    const content = entry.content as ArrayLike<number> | undefined
    const size = content?.length ?? 0
    if (size === 0) continue
    if (size > MAX_STREAM_BYTES) {
      throw new LegacyXlsError('limits_exceeded', `stream ${name} ${size} bytes`)
    }
    const workbook = new Uint8Array(size)
    for (let i = 0; i < size; i++) workbook[i] = content![i]!
    return { workbook }
  }
  throw new LegacyXlsError('missing_stream')
}
