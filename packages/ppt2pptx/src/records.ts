import { LegacyPptError } from './errors'

/**
 * The MS-PPT record walk. Every record opens with a 16-bit (version:4 |
 * instance:12), a 16-bit type, and a 32-byte length; container records carry
 * version 0xF and hold further records, atoms hold their own bytes.
 */

export const RECORD_HEADER = 8

export interface PptRecord {
  /** the 4-bit version */
  version: number
  /** the 12-bit instance */
  instance: number
  type: number
  /** payload length as the header declares */
  length: number
  /** offset of the payload inside the stream */
  offset: number
}

export const PPT_TYPES = {
  documentContainer: 0x03e8,
  environment: 0x03f2,
  ppDrawingGroup: 0x040b,
  ppDrawing: 0x040c,
  fontCollection: 0x07d5,
  colorSchemeAtom: 0x07f0,
  oePlaceholderAtom: 0x0bc3,
  outlineTextRefAtom: 0x0f9e,
  styleTextPropAtom: 0x0fa1,
  txMasterStyleAtom: 0x0fa3,
  textRulerAtom: 0x0fa6,
  fontEntityAtom: 0x0fb7,
  documentAtom: 0x03e9,
  endDocument: 0x03ea,
  slideContainer: 0x03ee,
  slideAtom: 0x03ef,
  notesContainer: 0x03f0,
  slidePersistAtom: 0x03f3,
  mainMasterContainer: 0x03f8,
  textHeaderAtom: 0x0f9f,
  textCharsAtom: 0x0fa0,
  textBytesAtom: 0x0fa8,
  slideListWithText: 0x0ff0,
  cString: 0x0fba,
  styleTextProp9Atom: 0x0fac,
  outlineTextProps9: 0x0fae,
  outlineTextPropsHeader9Atom: 0x0faf,
  list: 0x07d0,
  progTags: 0x1388,
  progBinaryTag: 0x138a,
  binaryTagExtension: 0x138b,
  userEditAtom: 0x0ff5,
  currentUserAtom: 0x0ff6,
  persistPtrIncrementalBlock: 0x1772,
  /** Escher OfficeArtClientTextbox: an atom whose payload holds MS-PPT text records */
  escherClientTextbox: 0xf00d,
} as const

/** Walk the records of one stream (flat; consumers recurse into containers). */
export function* walkRecords(
  stream: Uint8Array,
  from = 0,
  to = stream.length,
): Generator<PptRecord> {
  let at = from
  while (at + RECORD_HEADER <= to) {
    const head = stream[at]! | (stream[at + 1]! << 8)
    const type = stream[at + 2]! | (stream[at + 3]! << 8)
    const v = new DataView(stream.buffer, stream.byteOffset + at + 4, 4)
    const length = v.getUint32(0, true)
    const offset = at + RECORD_HEADER
    if (offset + length > to) {
      throw new LegacyPptError('truncated', `record ${type.toString(16)} claims ${length} bytes`)
    }
    yield { version: head & 0x000f, instance: head >> 4, type, length, offset }
    at = offset + length
  }
}

/**
 * The same walk for the drawing / style layers: a damaged record ends the
 * walk instead of failing the deck, so one bad master or shape costs only
 * itself.
 */
export function* walkRecordsLenient(
  stream: Uint8Array,
  from = 0,
  to = stream.length,
): Generator<PptRecord> {
  let at = from
  while (at + RECORD_HEADER <= to) {
    const head = stream[at]! | (stream[at + 1]! << 8)
    const type = stream[at + 2]! | (stream[at + 3]! << 8)
    const length = new DataView(stream.buffer, stream.byteOffset + at + 4, 4).getUint32(0, true)
    const offset = at + RECORD_HEADER
    if (offset + length > to) return
    yield { version: head & 0x000f, instance: head >> 4, type, length, offset }
    at = offset + length
  }
}

export function payloadOf(stream: Uint8Array, record: PptRecord): Uint8Array {
  return stream.subarray(record.offset, record.offset + record.length)
}

/**
 * True when the payload is itself a record stream: MS-PPT containers carry
 * version 0xF, and the Escher OfficeArtClientTextbox atom (version 0) holds
 * the shape's TextHeaderAtom / TextCharsAtom pair instead of raw bytes.
 */
export function isRecordContainer(record: PptRecord): boolean {
  return record.version === 0xf || record.type === PPT_TYPES.escherClientTextbox
}

/** decode a TextCharsAtom payload: UTF-16LE with PowerPoint's control runes */
export function decodeTextChars(payload: Uint8Array): string {
  return normalizeRunes(rawTextChars(payload))
}

/** decode a TextBytesAtom payload: one byte per char (Windows codepage 1252
    at heart; latin1 covers the bytes, the smart-quote runes get named) */
export function decodeTextBytes(payload: Uint8Array): string {
  return normalizeRunes(rawTextBytes(payload))
}

/** the text as PowerPoint stores it: '\r' ends a paragraph, '\u000b' breaks a line */
export function rawTextChars(payload: Uint8Array): string {
  const text = Buffer.from(payload.buffer, payload.byteOffset, payload.length).toString('utf16le')
  return nameSpecials(text)
}

export function rawTextBytes(payload: Uint8Array): string {
  const text = Buffer.from(payload.buffer, payload.byteOffset, payload.length).toString('latin1')
  return nameSpecials(text)
}

/** names the cp1252 runes and blanks the bullet markers, one char for one so run counts still line up */
function nameSpecials(text: string): string {
  return [...text]
    .map((ch) => (ch === '\u001e' || ch === '\u001f' ? ' ' : (CP1252_SPECIALS[ch] ?? ch)))
    .join('')
}

const CP1252_SPECIALS: Record<string, string> = {
  '\u0091': '\u2018',
  '\u0092': '\u2019',
  '\u0093': '\u201c',
  '\u0094': '\u201d',
  '\u0095': '\u2022',
  '\u0096': '\u2013',
  '\u0097': '\u2014',
}

/** PowerPoint text carries its own paragraph/format runes (\u000B vertical tab
    = soft break, \u001E bullet-in-paragraph, ...); they become newlines or
    disappear, so a converted slide reads like the original */
function normalizeRunes(text: string): string {
  return [...text]
    .map((ch) => {
      if (ch === '\u000b') return '\n'
      // the bullet-in-paragraph rune separates visually where a glyph follows
      if (ch === '\u001e' || ch === '\u001f') return ' '
      return CP1252_SPECIALS[ch] ?? ch
    })
    .join('')
    .replace(/\r\n?/g, '\n')
}
