/**
 * Comments (annotations) of a Word 97-2003 document.
 *
 * PlcfandRef (FIB pair 4) anchors each comment at a main-document CP (the
 * \x05 annotation reference character) and carries an ATRDPre10 per comment:
 * author initials, the author's index into SttbfRMark (pair 51) and a bookmark
 * tag. The tag names an ATNBE of SttbfAtnBkmk (pair 37) whose position selects
 * the PlcfAtnBkf (pair 42) entry; its ibkl selects the PlcfAtnBkl (pair 43)
 * CP, so [bkf CP, bkl CP) is the commented text. A tag of -1 is a point
 * comment. The comment text lives in the comment document (CPs after the
 * header document), split by PlcfandTxt (pair 5) like PlcfHdd splits headers.
 * Word 2002+ adds ATRDPost10 (pair 112) with the comment date.
 */
import { LegacyDocError, isLegacyDocError } from './errors'
import type { FibInfo } from './fib'
import type { PieceTable } from './clx'
import type { FormattedParagraph } from './formatting'
import type { StoryFormatter } from './headers'
import { decodeWordDocumentRange, mapWordControlChars } from './text'

export interface LegacyComment {
  /** position in PlcfandRef; the w:id of the written comment */
  id: number
  author: string
  initials: string
  /** ISO timestamp (ATRDPost10 dttm), absent in Word 97-2000 files */
  date?: string
  /** CP of the annotation reference character in the main document */
  referenceCp: number
  /** commented main-document range; absent for a point comment */
  rangeStartCp?: number
  rangeEndCp?: number
  /** comment text paragraphs (control characters mapped) */
  paragraphs: string[]
  /** the same paragraphs with resolved formatting, when a StoryFormatter was given */
  formatted?: FormattedParagraph[]
}

export interface CommentsResult {
  comments: LegacyComment[]
  warnings: string[]
}

export interface CommentsSource {
  fib: FibInfo
  wordDocument: Uint8Array
  table: Uint8Array
  pieces: PieceTable
}

const ATRD_PRE10_BYTES = 30
const ATRD_POST10_BYTES = 18
const XST_INITIALS_CHARS = 9
const BKF_BYTES = 4
const NO_BOOKMARK = -1

function badAtn(detail: string): never {
  throw new LegacyDocError('bad_atn', detail)
}

function viewOf(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

function checkRange(table: Uint8Array, fc: number, lcb: number, what: string): void {
  if (fc > table.length || lcb > table.length - fc)
    badAtn(`the ${what} range (${fc}+${lcb}) lies outside the ${table.length}-byte table stream`)
}

/** CP-only PLC (PlcfandTxt, PlcfAtnBkl): every 4 bytes one CP */
function readCps(table: Uint8Array, fc: number, count: number): number[] {
  const view = viewOf(table)
  const cps: number[] = []
  for (let index = 0; index < count; index++) cps.push(view.getUint32(fc + index * 4, true))
  return cps
}

/**
 * An STTB (MS-DOC 2.9.270): extended (0xFFFF, UTF-16 strings) or 8-bit,
 * cData entries of cch + text + cbExtra bytes each. Short or malformed
 * tables yield the entries that could be read.
 */
function readSttb(
  table: Uint8Array,
  fc: number,
  lcb: number,
): { strings: string[]; extra: Uint8Array[] } {
  const strings: string[] = []
  const extra: Uint8Array[] = []
  if (lcb < 4) return { strings, extra }
  const view = viewOf(table)
  const end = fc + lcb
  let offset = fc
  const extended = view.getUint16(offset, true) === 0xffff
  if (extended) offset += 2
  if (offset + 4 > end) return { strings, extra }
  const count = view.getUint16(offset, true)
  const cbExtra = view.getUint16(offset + 2, true)
  offset += 4
  for (let index = 0; index < count; index++) {
    const cch = extended ? view.getUint16(offset, true) : table[offset]!
    offset += extended ? 2 : 1
    const byteLength = extended ? cch * 2 : cch
    if (offset + byteLength + cbExtra > end) break
    let text = ''
    if (extended) {
      for (let char = 0; char < cch; char++)
        text += String.fromCharCode(view.getUint16(offset + char * 2, true))
    } else {
      for (let char = 0; char < cch; char++) text += String.fromCharCode(table[offset + char]!)
    }
    offset += byteLength
    strings.push(text)
    extra.push(table.subarray(offset, offset + cbExtra))
    offset += cbExtra
  }
  return { strings, extra }
}

/** DTTM (MS-DOC 2.9.65) -> ISO timestamp; undefined for the zero / unset value */
export function dttmToIso(dttm: number): string | undefined {
  if (dttm === 0) return undefined
  const minute = dttm & 0x3f
  const hour = (dttm >>> 6) & 0x1f
  const day = (dttm >>> 11) & 0x1f
  const month = (dttm >>> 16) & 0x0f
  const year = 1900 + ((dttm >>> 20) & 0x1ff)
  if (month < 1 || month > 12 || day < 1 || day > 31 || hour > 23 || minute > 59) return undefined
  const two = (value: number): string => String(value).padStart(2, '0')
  return `${year}-${two(month)}-${two(day)}T${two(hour)}:${two(minute)}:00Z`
}

interface Atrd {
  initials: string
  ibst: number
  tag: number
}

function readAtrd(table: Uint8Array, offset: number): Atrd {
  const view = viewOf(table)
  const cch = Math.min(view.getUint16(offset, true), XST_INITIALS_CHARS)
  let initials = ''
  for (let char = 0; char < cch; char++)
    initials += String.fromCharCode(view.getUint16(offset + 2 + char * 2, true))
  return {
    initials,
    ibst: view.getUint16(offset + 20, true),
    tag: view.getInt32(offset + 26, true),
  }
}

interface AnnotationBookmarks {
  tags: number[]
  startCps: number[]
  ibkls: number[]
  endCps: number[]
}

function readAnnotationBookmarks(fib: FibInfo, table: Uint8Array): AnnotationBookmarks | undefined {
  if (fib.lcbPlcfAtnBkf === 0 || fib.lcbPlcfAtnBkl === 0) return undefined
  checkRange(table, fib.fcPlcfAtnBkf, fib.lcbPlcfAtnBkf, 'annotation bookmark start table')
  checkRange(table, fib.fcPlcfAtnBkl, fib.lcbPlcfAtnBkl, 'annotation bookmark end table')
  if (fib.lcbPlcfAtnBkf < 4 + BKF_BYTES || (fib.lcbPlcfAtnBkf - 4) % (4 + BKF_BYTES) !== 0)
    badAtn(`the annotation bookmark start table has an invalid length (lcb ${fib.lcbPlcfAtnBkf})`)
  const count = (fib.lcbPlcfAtnBkf - 4) / (4 + BKF_BYTES)
  const startCps = readCps(table, fib.fcPlcfAtnBkf, count)
  const view = viewOf(table)
  const bkfStart = fib.fcPlcfAtnBkf + (count + 1) * 4
  const ibkls: number[] = []
  for (let index = 0; index < count; index++)
    ibkls.push(view.getUint16(bkfStart + index * BKF_BYTES, true))
  // PlcfAtnBkl is CP-only in every version this reader supports; tolerate the 2-byte BKL layout too
  let endCount: number
  if (fib.lcbPlcfAtnBkl % 4 === 0) endCount = fib.lcbPlcfAtnBkl / 4 - 1
  else if ((fib.lcbPlcfAtnBkl - 4) % 6 === 0) endCount = (fib.lcbPlcfAtnBkl - 4) / 6
  else badAtn(`the annotation bookmark end table has an invalid length (lcb ${fib.lcbPlcfAtnBkl})`)
  const endCps = readCps(table, fib.fcPlcfAtnBkl, endCount)
  let tags: number[] = []
  if (fib.lcbSttbfAtnBkmk > 0) {
    checkRange(table, fib.fcSttbfAtnBkmk, fib.lcbSttbfAtnBkmk, 'annotation bookmark name table')
    // Word stores each ATNBE (bmc, lTag, lTagOld) as the 10 cbExtra bytes of an empty string
    tags = readSttb(table, fib.fcSttbfAtnBkmk, fib.lcbSttbfAtnBkmk).extra.map((atnbe) =>
      atnbe.length >= 6 ? viewOf(atnbe).getInt32(2, true) : NO_BOOKMARK,
    )
  }
  return { tags, startCps, ibkls, endCps }
}

/** split one comment story like a header story: guard mark off, trailing paragraph mark is the last break */
function storyParagraphs(raw: string): { text: string; paragraphs: string[] } {
  let text = raw
  if (text.endsWith('\r')) text = text.slice(0, -1)
  const rawParagraphs = text.split('\r')
  if (rawParagraphs.length > 1 && rawParagraphs[rawParagraphs.length - 1] === '')
    rawParagraphs.pop()
  return { text, paragraphs: rawParagraphs.map(mapWordControlChars) }
}

/**
 * Parse the comments of the main document. Missing tables with declared
 * comment text are a warning (comments omitted); structurally invalid tables
 * raise 'bad_atn'. Comments anchored outside the main text are skipped with
 * a warning.
 */
export function parseComments(source: CommentsSource, formatter?: StoryFormatter): CommentsResult {
  const { fib, wordDocument, table, pieces } = source
  const warnings: string[] = []
  if (fib.lcbPlcfandRef === 0 || fib.lcbPlcfandTxt === 0) {
    if (fib.ccpAtn > 0)
      warnings.push(
        'the document declares comment text but has no comment table; comments are omitted',
      )
    return { comments: [], warnings }
  }
  checkRange(table, fib.fcPlcfandRef, fib.lcbPlcfandRef, 'comment reference table')
  checkRange(table, fib.fcPlcfandTxt, fib.lcbPlcfandTxt, 'comment text table')
  if (
    fib.lcbPlcfandRef < 4 + ATRD_PRE10_BYTES ||
    (fib.lcbPlcfandRef - 4) % (4 + ATRD_PRE10_BYTES) !== 0
  )
    badAtn(
      `the comment reference table has an invalid length (lcb ${fib.lcbPlcfandRef}, expected 34n+4)`,
    )
  const count = (fib.lcbPlcfandRef - 4) / (4 + ATRD_PRE10_BYTES)
  if (fib.lcbPlcfandTxt % 4 !== 0 || fib.lcbPlcfandTxt / 4 - 2 !== count)
    badAtn(
      `the comment text table has ${fib.lcbPlcfandTxt / 4} CPs, but ${count} comment(s) need ${count + 2}`,
    )
  const referenceCps = readCps(table, fib.fcPlcfandRef, count + 1)
  const atrdStart = fib.fcPlcfandRef + (count + 1) * 4
  const atrds: Atrd[] = []
  for (let index = 0; index < count; index++)
    atrds.push(readAtrd(table, atrdStart + index * ATRD_PRE10_BYTES))
  const textCps = readCps(table, fib.fcPlcfandTxt, count + 2)
  for (let index = 0; index + 1 < textCps.length; index++)
    if (textCps[index + 1]! < textCps[index]!)
      badAtn(`the comment text CPs are not ordered (${textCps[index]} > ${textCps[index + 1]})`)

  let authors: string[] = []
  if (fib.lcbSttbfRMark > 0) {
    checkRange(table, fib.fcSttbfRMark, fib.lcbSttbfRMark, 'revision author table')
    authors = readSttb(table, fib.fcSttbfRMark, fib.lcbSttbfRMark).strings
  }
  const bookmarks = readAnnotationBookmarks(fib, table)
  const dates: Array<string | undefined> = []
  if (fib.lcbAtrdExtra > 0) {
    checkRange(table, fib.fcAtrdExtra, fib.lcbAtrdExtra, 'comment date table')
    if (fib.lcbAtrdExtra === count * ATRD_POST10_BYTES) {
      const view = viewOf(table)
      for (let index = 0; index < count; index++)
        dates.push(dttmToIso(view.getUint32(fib.fcAtrdExtra + index * ATRD_POST10_BYTES, true)))
    } else
      warnings.push(
        'the comment date table does not match the comment count; comment dates are omitted',
      )
  }

  const commentStart = fib.ccpText + fib.ccpFtn + fib.ccpHdd
  const commentLength = fib.ccpAtn
  const totalCharacters = pieces.cps.at(-1) ?? 0
  if (commentStart + commentLength > totalCharacters)
    badAtn(
      `the comment document (CPs ${commentStart}..${commentStart + commentLength}) lies outside the ${totalCharacters}-character text`,
    )
  const comments: LegacyComment[] = []
  let authorsMissing = false
  let outside = 0
  for (let index = 0; index < count; index++) {
    const referenceCp = referenceCps[index]!
    if (referenceCp >= fib.ccpText) {
      outside++
      continue
    }
    const atrd = atrds[index]!
    const author = authors[atrd.ibst]
    if (author === undefined) authorsMissing = true
    const relativeStart = Math.min(textCps[index]!, commentLength)
    const relativeEnd = Math.min(textCps[index + 1]!, commentLength)
    let raw = ''
    try {
      if (relativeEnd > relativeStart)
        raw = decodeWordDocumentRange(
          wordDocument,
          pieces,
          commentStart + relativeStart,
          commentStart + relativeEnd,
        )
    } catch (error) {
      if (isLegacyDocError(error)) badAtn(`comment ${index} cannot be read: ${error.message}`)
      throw error
    }
    const story = storyParagraphs(raw)
    const comment: LegacyComment = {
      id: index,
      author: author ?? '',
      initials: atrd.initials,
      referenceCp,
      paragraphs: story.paragraphs,
    }
    const date = dates[index]
    if (date) comment.date = date
    if (formatter && story.text !== '') {
      const formatted = formatter.paragraphs(
        commentStart + relativeStart,
        commentStart + relativeStart + story.text.length,
      )
      if (formatted.length > 1 && formatted[formatted.length - 1]!.text === '') formatted.pop()
      if (formatted.length === story.paragraphs.length) comment.formatted = formatted
    }
    if (bookmarks && atrd.tag !== NO_BOOKMARK) {
      const bookmarkIndex = bookmarks.tags.length > 0 ? bookmarks.tags.indexOf(atrd.tag) : index
      const ibkl = bookmarks.ibkls[bookmarkIndex]
      const start = bookmarks.startCps[bookmarkIndex]
      const end = ibkl === undefined ? undefined : bookmarks.endCps[ibkl]
      if (start === undefined || end === undefined)
        warnings.push(
          `comment ${index} refers to a missing annotation bookmark; it was imported without a text range`,
        )
      // an empty bookmark is how Word stores a comment on a collapsed selection
      else if (end > start && end <= fib.ccpText) {
        comment.rangeStartCp = start
        comment.rangeEndCp = end
      }
    }
    comments.push(comment)
  }
  if (authorsMissing)
    warnings.push(
      'some comment authors are missing from the revision author table; their names are empty',
    )
  if (outside > 0)
    warnings.push(
      `${outside} comment(s) anchored in headers, footers or footnotes are not converted yet; they were omitted`,
    )
  return { comments, warnings }
}
