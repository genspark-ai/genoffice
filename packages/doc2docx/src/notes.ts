/**
 * Footnotes and endnotes (MS-DOC 2.8.17-2.8.20).
 *
 * PlcffndRef (FIB pair 2) holds n+1 body CPs plus n 2-byte FRDs (nAuto 0 =
 * custom mark, otherwise auto-numbered); the character at each CP is the
 * reference mark (0x02 for auto marks, the mark itself for custom ones).
 * PlcffndTxt (pair 3) is a CP-only PLC with n+2 entries over the footnote
 * document, which starts at CP ccpText; each note's text ends with its own
 * paragraph mark and the document ends with one guard paragraph mark. The
 * endnote tables (pairs 46/47) and the endnote document (after the header and
 * annotation documents) have the same shape.
 */
import { LegacyDocError, isLegacyDocError } from './errors'
import type { FibInfo } from './fib'
import type { PieceTable } from './clx'
import type { FormattedParagraph, NoteMark } from './formatting'
import type { StoryFormatter } from './headers'
import { decodeWordDocumentRange, mapWordControlChars } from './text'

export type NoteKind = 'footnote' | 'endnote'

export interface LegacyNote {
  id: number
  kind: NoteKind
  /** nAuto 0: the mark is literal text (kept in the note's runs), not a numbered mark */
  custom: boolean
  /** the note text with resolved formatting; an auto mark is a run with note.role 'self' */
  paragraphs: FormattedParagraph[]
}

export interface LegacyDocNotes {
  footnotes: LegacyNote[]
  endnotes: LegacyNote[]
  /** the reference marks in the body, by CP */
  marks: NoteMark[]
  warnings: string[]
}

export interface NotesSource {
  fib: FibInfo
  wordDocument: Uint8Array
  table: Uint8Array
  pieces: PieceTable
}

const AUTO_MARK = '\u0002'

function badNote(detail: string): never {
  throw new LegacyDocError('bad_note', detail)
}

function u32s(table: Uint8Array, fc: number, count: number): number[] {
  const view = new DataView(table.buffer, table.byteOffset, table.byteLength)
  const out: number[] = []
  for (let index = 0; index < count; index++) out.push(view.getUint32(fc + index * 4, true))
  return out
}

function tables(fib: FibInfo, kind: NoteKind) {
  return kind === 'footnote'
    ? {
        ccp: fib.ccpFtn,
        ref: { fc: fib.fcPlcffndRef, lcb: fib.lcbPlcffndRef },
        txt: { fc: fib.fcPlcffndTxt, lcb: fib.lcbPlcffndTxt },
        storyStart: fib.ccpText,
      }
    : {
        ccp: fib.ccpEdn,
        ref: { fc: fib.fcPlcfendRef, lcb: fib.lcbPlcfendRef },
        txt: { fc: fib.fcPlcfendTxt, lcb: fib.lcbPlcfendTxt },
        storyStart: fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpAtn,
      }
}

function plainParagraphs(raw: string, startCp: number): FormattedParagraph[] {
  let cp = startCp
  return raw.split('\r').map((paragraph) => {
    const text = mapWordControlChars(paragraph)
    const out: FormattedParagraph = {
      startCp: cp,
      endCp: cp + paragraph.length,
      text,
      formatting: {},
      runs: text === '' ? [] : [{ text, formatting: {} }],
    }
    cp += paragraph.length + 1
    return out
  })
}

function parseKind(
  source: NotesSource,
  kind: NoteKind,
  formatter: StoryFormatter | undefined,
  warnings: string[],
): { notes: LegacyNote[]; marks: NoteMark[] } {
  const { fib, table, wordDocument, pieces } = source
  const { ccp, ref, txt, storyStart } = tables(fib, kind)
  const plural = `${kind}s`
  if (ref.lcb === 0 || txt.lcb === 0) {
    if (ccp > 0)
      warnings.push(
        `the document declares ${kind} text but has no ${kind} ${
          txt.lcb === 0 ? 'text' : 'reference'
        } table; ${plural} are omitted`,
      )
    return { notes: [], marks: [] }
  }
  if (ref.lcb < 10 || (ref.lcb - 4) % 6 !== 0)
    badNote(`the ${kind} reference table has an invalid length (lcb ${ref.lcb}, expected 6n+4)`)
  if (txt.lcb < 12 || txt.lcb % 4 !== 0)
    badNote(
      `the ${kind} text table has an invalid length (lcb ${txt.lcb}, expected a multiple of 4)`,
    )
  for (const [pair, what] of [
    [ref, 'reference'],
    [txt, 'text'],
  ] as const)
    if (pair.fc > table.length || pair.lcb > table.length - pair.fc)
      badNote(
        `the ${kind} ${what} table range (${pair.fc}+${pair.lcb}) lies outside the ${table.length}-byte table stream`,
      )
  const count = (ref.lcb - 4) / 6
  const textCount = txt.lcb / 4 - 2
  if (textCount !== count)
    badNote(`the document has ${count} ${kind} references but ${textCount} ${kind} texts`)
  const refCps = u32s(table, ref.fc, count + 1)
  const textCps = u32s(table, txt.fc, count + 2)
  for (const cps of [refCps, textCps])
    for (let index = 0; index + 1 < cps.length; index++)
      if (cps[index + 1]! < cps[index]!)
        badNote(`the ${kind} CPs are not ordered (${cps[index]} > ${cps[index + 1]})`)
  const totalCharacters = pieces.cps.at(-1) ?? 0
  if (storyStart + ccp > totalCharacters)
    badNote(
      `the ${kind} document (CPs ${storyStart}..${storyStart + ccp}) lies outside the ${totalCharacters}-character text`,
    )
  const view = new DataView(table.buffer, table.byteOffset, table.byteLength)
  const frdStart = ref.fc + (count + 1) * 4
  const notes: LegacyNote[] = []
  const marks: NoteMark[] = []
  for (let index = 0; index < count; index++) {
    const id = index + 1
    const custom = view.getUint16(frdStart + index * 2, true) === 0
    const refCp = refCps[index]!
    if (refCp < fib.ccpText) marks.push({ cp: refCp, kind, id, role: 'reference', custom })
    else
      warnings.push(
        `${kind} ${id} is referenced outside the main document; the reference mark was omitted`,
      )
    const start = storyStart + Math.min(textCps[index]!, ccp)
    const end = storyStart + Math.min(textCps[index + 1]!, ccp)
    let raw = ''
    if (end > start) {
      try {
        raw = decodeWordDocumentRange(wordDocument, pieces, start, end)
      } catch (error) {
        if (isLegacyDocError(error)) badNote(`${kind} ${id} cannot be read: ${error.message}`)
        throw error
      }
    }
    if (raw.endsWith('\r')) raw = raw.slice(0, -1)
    const selfMarks: NoteMark[] = []
    const markAt = custom ? -1 : raw.indexOf(AUTO_MARK)
    if (markAt >= 0) selfMarks.push({ cp: start + markAt, kind, id, role: 'self', custom: false })
    let paragraphs = formatter?.paragraphs(start, start + raw.length, selfMarks) ?? []
    if (paragraphs.length === 0) {
      paragraphs = plainParagraphs(raw, start)
      if (selfMarks.length > 0)
        paragraphs[0]!.runs.unshift({ text: '', formatting: {}, note: selfMarks[0]! })
    }
    notes.push({ id, kind, custom, paragraphs })
  }
  return { notes, marks }
}

/**
 * Parse the footnotes and endnotes of a document. Broken structural data is a
 * 'bad_note' error; a document that declares note text without the tables
 * yields a warning and no notes.
 */
export function parseNotes(source: NotesSource, formatter?: StoryFormatter): LegacyDocNotes {
  const warnings: string[] = []
  const footnotes = parseKind(source, 'footnote', formatter, warnings)
  const endnotes = parseKind(source, 'endnote', formatter, warnings)
  return {
    footnotes: footnotes.notes,
    endnotes: endnotes.notes,
    marks: [...footnotes.marks, ...endnotes.marks].sort((left, right) => left.cp - right.cp),
    warnings,
  }
}
