import { LegacyDocError } from './errors'
import type { PieceTable } from './clx'

export interface LegacyDocText {
  /** piece text in CP order with Word's control marks still in place (\r paragraph marks, \x07 cell marks, ...) */
  rawText: string
  /** the mapped body text: paragraphs separated by newlines */
  text: string
  /** the mapped paragraphs; the final paragraph mark leaves a trailing empty paragraph */
  paragraphs: string[]
}

/**
 * Words bytes 0x80-0x9F of Windows-1252 (the complement of Latin-1): curly
 * quotes, dashes, ellipsis, trademark, ... Word's compressed pieces use this
 * codepage, not ISO-8859-1.
 */
const CP1252_HIGH: Record<number, string> = {
  0x80: '\u20ac',
  0x82: '\u201a',
  0x83: '\u0192',
  0x84: '\u201e',
  0x85: '\u2026',
  0x86: '\u2020',
  0x87: '\u2021',
  0x88: '\u02c6',
  0x89: '\u2030',
  0x8a: '\u0160',
  0x8b: '\u2039',
  0x8c: '\u0152',
  0x8e: '\u017d',
  0x91: '\u2018',
  0x92: '\u2019',
  0x93: '\u201c',
  0x94: '\u201d',
  0x95: '\u2022',
  0x96: '\u2013',
  0x97: '\u2014',
  0x98: '\u02dc',
  0x99: '\u2122',
  0x9a: '\u0161',
  0x9b: '\u203a',
  0x9c: '\u0153',
  0x9e: '\u017e',
  0x9f: '\u0178',
}

const REPLACEMENT = '\ufffd'

/** Windows-1252 piece bytes -> one character each (undefined slots are undefined in the codepage too) */
function decodeCompressed(bytes: Uint8Array): string {
  const chunks: string[] = []
  for (const byte of bytes) {
    if (byte < 0x80) chunks.push(String.fromCharCode(byte))
    else if (byte < 0xa0) chunks.push(CP1252_HIGH[byte] ?? REPLACEMENT)
    else chunks.push(String.fromCharCode(byte))
  }
  return chunks.join('')
}

/** UTF-16LE piece bytes -> one UTF-16 code unit per two bytes (surrogate pairs stay pairs) */
function decodeUtf16Le(bytes: Uint8Array): string {
  const units: number[] = []
  for (let offset = 0; offset + 1 < bytes.length; offset += 2)
    units.push(bytes[offset]! | (bytes[offset + 1]! << 8))
  const chunks: string[] = []
  for (let offset = 0; offset < units.length; offset += 4096)
    chunks.push(String.fromCharCode(...units.slice(offset, offset + 4096)))
  return chunks.join('')
}

const FIELD_INSTRUCTION_START = '\u0013'
const FIELD_SEPARATOR = '\u0014'
const FIELD_END = '\u0015'

export interface WordControlMapper {
  /** map the next slice of one paragraph; field state carries across calls */
  map(text: string): string
}

/**
 * Map Word text to display text (MS-DOC control marks): cell mark -> tab,
 * vertical tab / form feed -> newline, non-breaking hyphen -> '-', optional
 * hyphen and friends -> '', other C0 controls dropped. Field markup keeps the
 * field result and drops the instruction:
 * \x13 instruction \x14 result \x15 -> result.
 *
 * Text is passed in consecutive slices (one per CHPX run); field state carries
 * across them so a field split by a formatting run still maps correctly.
 */
export function createWordControlMapper(): WordControlMapper {
  const fields: Array<'instruction' | 'result'> = []
  return {
    map(text: string): string {
      let out = ''
      for (const char of text) {
        if (char === FIELD_INSTRUCTION_START) {
          fields.push('instruction')
          continue
        }
        if (char === FIELD_SEPARATOR) {
          if (fields.length > 0) fields[fields.length - 1] = 'result'
          continue
        }
        if (char === FIELD_END) {
          fields.pop()
          continue
        }
        if (fields.includes('instruction')) continue
        switch (char) {
          case '\t': // a literal tab inside the text is kept
          case '\n':
            out += char
            break
          case '\u0007': // cell mark / table column separator
            out += '\t'
            break
          case '\u000b': // vertical tab / line break
          case '\u000c': // form feed / page break
            out += '\n'
            break
          case '\u001e': // non-breaking hyphen
            out += '-'
            break
          case '\u001f': // optional hyphen
          case '\u0001':
          case '\u0002':
          case '\u0005':
          case '\u0008':
            break
          default:
            if (char.charCodeAt(0) < 0x20) break
            out += char
        }
      }
      return out
    },
  }
}

/** Map one paragraph's Word control marks (a fresh field state for each paragraph). */
export function mapWordControlChars(paragraph: string): string {
  return createWordControlMapper().map(paragraph)
}

/**
 * Read the main document story: the piece text for CPs [0, ccpText), mapped
 * to paragraphs. Footnote/header/textbox stories are later stages.
 */
export function extractBodyText(
  wordDocument: Uint8Array,
  pieceTable: PieceTable,
  ccpText: number,
): LegacyDocText {
  const chunks: string[] = []
  let remaining = ccpText
  for (const piece of pieceTable.pieces) {
    if (remaining <= 0) break
    const chars = Math.min(piece.cpEnd - piece.cpStart, remaining)
    if (chars <= 0) continue
    if (piece.compressed) {
      chunks.push(decodeCompressed(wordDocument.subarray(piece.fc, piece.fc + chars)))
    } else {
      chunks.push(decodeUtf16Le(wordDocument.subarray(piece.fc, piece.fc + chars * 2)))
    }
    remaining -= chars
  }
  if (remaining > 0)
    throw new LegacyDocError(
      'bad_piece_table',
      `the piece table covers ${ccpText - remaining} of the ${ccpText} body characters`,
    )
  const rawText = chunks.join('')
  const paragraphs = rawText.split('\r').map(mapWordControlChars)
  return { rawText, text: paragraphs.join('\n'), paragraphs }
}
