/**
 * Legacy Word 97-2003 (.doc) import: OLE2/CFB container + FIB + CLX piece table
 * -> body text (stage 1) and a per-run / per-paragraph formatting model from
 * the CHPX/PAPX FKPs (stage 2). Styles/tables/images are stage 3; the whole
 * package is internal only.
 */
import { openCompoundFile } from './compound'
import { parseFib, validateFibTableRanges, type FibInfo } from './fib'
import { parseClx, type PieceTable } from './clx'
import { extractBodyText } from './text'
import { readFormatting, type FormattedParagraph } from './formatting'
import { LegacyDocError } from './errors'

export { LEGACY_DOC_ERROR_CODES, LegacyDocError, isLegacyDocError } from './errors'
export type { LegacyDocErrorCode } from './errors'
export { LEGACY_DOC_LIMITS, openCompoundFile } from './compound'
export type { CompoundFile, CompoundFileLimits } from './compound'
export {
  FIB_IDENT,
  MAX_SUPPORTED_NFIB,
  MIN_SUPPORTED_NFIB,
  parseFib,
  validateFibTableRanges,
} from './fib'
export type { FibInfo, FibPair } from './fib'
export { parseClx } from './clx'
export type { ParseClxOptions, Piece, PieceTable } from './clx'
export { createWordControlMapper, extractBodyText, mapWordControlChars } from './text'
export type { LegacyDocText, WordControlMapper } from './text'
export { parsePlcBte } from './plcf'
export type { PlcBte } from './plcf'
export { FKP_PAGE_BYTES, readChpxFkp, readPapxFkp } from './fkp'
export type { ChpxRun, PapxEntry } from './fkp'
export {
  SPRM_LONG_PARAGRAPH,
  SPRM_LONG_TABLE,
  iterateSprms,
  readSprmI16,
  readSprmU16,
  readSprmU32,
  readSprmU8,
  sprmOperandSize,
} from './sprm'
export type { Sprm } from './sprm'
export { cpOfFc, fcOfCp, readFormatting } from './formatting'
export type {
  CharacterFormatting,
  FcResolution,
  FormattedParagraph,
  FormattedRun,
  FormattingSource,
  ParagraphFormatting,
  UnderlineKind,
} from './formatting'

/** Everything stage 1 reads from one .doc file. */
export interface LegacyDoc {
  fib: FibInfo
  pieces: PieceTable
  /** raw WordDocument stream bytes (FIB + text pieces + later-stage tables) */
  wordDocument: Uint8Array
  /** raw table stream bytes (0Table or 1Table, per FibBase.fWhichTblStm) */
  table: Uint8Array
  /** piece text in CP order with Word control marks still in place */
  rawText: string
  /** mapped body text, paragraphs separated by newlines */
  text: string
  /** mapped paragraphs (the final paragraph mark leaves a trailing empty one) */
  paragraphs: string[]
  /** other OLE2 streams by name (the Data stream is needed for sprmPHugePapx) */
  stream(name: string): Uint8Array | undefined
}

/**
 * Open a Word 97-2003 .doc and read its body text. Throws LegacyDocError with
 * a stable code for every malformed / unsupported / oversized input.
 */
export function openLegacyDoc(bytes: Uint8Array): LegacyDoc {
  const compound = openCompoundFile(bytes)
  const wordDocument = compound.stream('WordDocument')
  if (!wordDocument)
    throw new LegacyDocError('missing_stream', 'the OLE2 container has no WordDocument stream')
  const fib = parseFib(wordDocument)

  const table = compound.stream(fib.tableStream)
  if (!table)
    throw new LegacyDocError(
      'missing_stream',
      `the header names the ${fib.tableStream} stream (fWhichTblStm) but the container has no such stream`,
    )
  validateFibTableRanges(fib, table.length)
  const pieces = parseClx(table, fib.fcClx, fib.lcbClx, {
    ccpText: fib.ccpText,
    wordDocumentLength: wordDocument.length,
  })
  const extracted = extractBodyText(wordDocument, pieces, fib.ccpText)
  return {
    fib,
    pieces,
    wordDocument,
    table,
    ...extracted,
    stream: (name) => compound.stream(name),
  }
}

/** Just the body text of a .doc (the common read path). */
export function readLegacyDocText(bytes: Uint8Array): { text: string; paragraphs: string[] } {
  const doc = openLegacyDoc(bytes)
  return { text: doc.text, paragraphs: doc.paragraphs }
}

/** Stage 2 result: the paragraph text plus its formatting model. */
export interface LegacyDocModel {
  fib: FibInfo
  text: string
  paragraphs: FormattedParagraph[]
  /** non-fatal problems (e.g. paragraph properties in a missing Data stream) */
  warnings: string[]
}

/** Open a .doc and read body text plus the run/paragraph formatting model. */
export function readLegacyDocModel(bytes: Uint8Array): LegacyDocModel {
  const doc = openLegacyDoc(bytes)
  const warnings: string[] = []
  const paragraphs = readFormatting(doc, warnings)
  return { fib: doc.fib, text: doc.text, paragraphs, warnings }
}

/* stage 3a: .docx writer */
export {
  conversionWarnings,
  convertLegacyDocToDocx,
  generateBodyParagraphsXml,
  toDocxRun,
  toParaFormat,
  writeLegacyDocx,
} from './docx-writer'
