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
  const text = Buffer.from(payload.buffer, payload.byteOffset, payload.length).toString('utf16le')
  return normalizeRunes(text)
}

/** decode a TextBytesAtom payload: one byte per char (ANSI, i.e. Windows-1252;
    latin1 covers the bytes outside the 0x80–0x9F block) */
export function decodeTextBytes(payload: Uint8Array): string {
  const text = Buffer.from(payload.buffer, payload.byteOffset, payload.length).toString('latin1')
  return normalizeRunes(text)
}

/**
 * Windows-1252's 0x80–0x9F block: latin1 reads those bytes as C1 controls, but
 * in the ANSI codepage a slide's bytes atom comes from they are punctuation.
 * The five bytes cp1252 leaves undefined (0x81, 0x8D, 0x8F, 0x90, 0x9D) keep
 * their latin1 value.
 */
const CP1252_SPECIALS: Record<string, string> = {
  '\u0080': '\u20ac', // €
  '\u0082': '\u201a', // ‚
  '\u0083': '\u0192', // ƒ
  '\u0084': '\u201e', // „
  '\u0085': '\u2026', // …
  '\u0086': '\u2020', // †
  '\u0087': '\u2021', // ‡
  '\u0088': '\u02c6', // ˆ
  '\u0089': '\u2030', // ‰
  '\u008a': '\u0160', // Š
  '\u008b': '\u2039', // ‹
  '\u008c': '\u0152', // Œ
  '\u008e': '\u017d', // Ž
  '\u0091': '\u2018', // ‘
  '\u0092': '\u2019', // ’
  '\u0093': '\u201c', // “
  '\u0094': '\u201d', // ”
  '\u0095': '\u2022', // •
  '\u0096': '\u2013', // –
  '\u0097': '\u2014', // —
  '\u0098': '\u02dc', // ˜
  '\u0099': '\u2122', // ™
  '\u009a': '\u0161', // š
  '\u009b': '\u203a', // ›
  '\u009c': '\u0153', // œ
  '\u009e': '\u017e', // ž
  '\u009f': '\u0178', // Ÿ
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
