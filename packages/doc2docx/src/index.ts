/**
 * Legacy Word 97-2003 (.doc) import: OLE2/CFB container + FIB + CLX piece table
 * -> body text (stage 1), a per-run / per-paragraph formatting model from the
 * CHPX/PAPX FKPs (stage 2), the STSH stylesheet resolved into that model
 * (stage 3b) and a .docx writer (stage 3a). Tables/images and the docs UI are
 * later stages; the whole package is internal only.
 */
import { openCompoundFile } from './compound'
import { parseFib, validateFibTableRanges, type FibInfo } from './fib'
import { parseClx, type PieceTable } from './clx'
import {
  decodeWordDocumentRange,
  extractBodyText,
  normalizeSectionMarks,
  textFromRaw,
} from './text'
import { readFormatting, readFormattingDetailed, type FormattedParagraph } from './formatting'
import {
  readResolvedFormatting,
  readStyleSheet,
  resolveFormatting,
  fieldRangesWithin,
  markFormattingEntry,
  resolveParagraphFormattingFor,
  resolveParagraphMarkFormatting,
  resolveRunsForRange,
} from './resolve'
import type { FormattingRead, FormattingSource } from './formatting'
import type { StyleSheet } from './styles'
import { parseSections, sectionMarkCps, type SectionInfo } from './sections'
import { parseHeaderStories, type SectionHeaderFooter, type StoryFormatter } from './headers'
import { parseInlineImages } from './images'
import { parseFloatingPictures, type FloatingPicturesResult } from './drawings'
import { parseTextBoxStories, textBoxDocumentStart } from './textboxes'
import { parseListTables } from './lists'
import { parseFields } from './fields'
import { buildLegacyDocBody, type LegacyDocBody } from './tables'
import { parseNotes, type LegacyDocNotes } from './notes'
import { parseComments, type LegacyComment } from './comments'
import { LegacyDocError } from './errors'

export { LEGACY_DOC_ERROR_CODES, LegacyDocError, isLegacyDocError } from './errors'
export type { LegacyDocErrorCode } from './errors'
export {
  LEGACY_DOC_LIMITS,
  OLE2_HEADER_BYTES,
  OLE2_MAGIC,
  hasOle2Signature,
  openCompoundFile,
  openOle2Container,
} from './compound'
export type {
  CompoundFile,
  CompoundFileLimits,
  Ole2FailureCode,
  Ole2Limits,
  Ole2OpenResult,
  Ole2Streams,
} from './compound'
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
export {
  applyCharacterGrpprl,
  applyParagraphGrpprl,
  characterFormattingOf,
  cpOfFc,
  fcOfCp,
  readFormatting,
  readFormattingDetailed,
} from './formatting'
export type {
  CharacterFormatting,
  CharacterIndents,
  FcResolution,
  FormattedParagraph,
  FormattedRun,
  FormattingRead,
  FormattingSource,
  NoteMark,
  ParagraphBorder,
  ParagraphBorderSide,
  ParagraphContext,
  ParagraphFormatting,
  RawCharacterRun,
  RawParagraphSprms,
  TabStop,
  ToggleMode,
  UnderlineKind,
} from './formatting'
export { parseStsh } from './styles'
export type { Style, StyleKind, StyleSheet } from './styles'
export { fontNamesFor, parseFontTable } from './fonts'
export type { DefaultFontIndices, FontNames, FontTable } from './fonts'
export { isBulletLevel, levelTextForWriter, parseListTables } from './lists'
export type { ListDefinition, ListLevel, ListTables } from './lists'
export { parseFields } from './fields'
export type { FieldParseResult, HyperlinkField } from './fields'
export { DEFAULT_SECTION_GEOMETRY, parseSections } from './sections'
export { EMU_PER_TWIP, imageExtentEmu, parseInlineImages } from './images'
export type { ImageBlock, InlineImageResult } from './images'
export { parseFloatingPictures, readBlip } from './drawings'
export type {
  FloatingPicture,
  FloatingPicturesResult,
  FloatingShape,
  FloatingTextBox,
  PictureMedia,
  PictureMediaType,
} from './drawings'
export { parseTextBoxStories, textBoxDocumentStart } from './textboxes'
export type { TextBoxStories, TextBoxStoriesSource, TextBoxStory } from './textboxes'
export type { SectionBreakType, SectionInfo, SectionMargins, SectionSource } from './sections'
export { parseHeaderStories } from './headers'
export { parseNotes } from './notes'
export type { LegacyDocNotes, LegacyNote, NoteKind, NotesSource } from './notes'
export type {
  HeaderFooterResult,
  HeaderFooterStory,
  HeaderStoriesSource,
  SectionHeaderFooter,
  StoryFormatter,
} from './headers'
export {
  DEFAULT_CHARACTER_STYLE_ISTD,
  DEFAULT_PARAGRAPH_STYLE_ISTD,
  readResolvedFormatting,
  readStyleSheet,
  resolveFormatting,
  resolveParagraphFormattingFor,
  resolveRunsForRange,
} from './resolve'
export type { FormattedLinkRange, ParagraphRunRequest } from './resolve'
export { parseComments } from './comments'
export type { CommentsResult, CommentsSource, LegacyComment } from './comments'
export { buildLegacyDocBody } from './tables'
export type {
  BodyBlock,
  CellBorders,
  LegacyDocBody,
  TableBorders,
  TableCell,
  TableCellPadding,
  TableModel,
  TableRow,
} from './tables'

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
  const body = extractBodyText(wordDocument, pieces, fib.ccpText)
  const normalized = normalizeSectionMarks(body.rawText, sectionMarkCps(fib, table))
  const extracted = normalized === body.rawText ? body : textFromRaw(normalized)
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

/** Open a .doc and read body text plus the direct-formatting model (stage 1/2 read). */
export function readLegacyDocModel(bytes: Uint8Array): LegacyDocModel {
  const doc = openLegacyDoc(bytes)
  const warnings: string[] = []
  const paragraphs = readFormatting(doc, warnings)
  return { fib: doc.fib, text: doc.text, paragraphs, warnings }
}

/**
 * Open a .doc and resolve its STSH stylesheet into the formatting model
 * (stage 3b read): the paragraphs carry style-based plus direct formatting.
 * Documents without an STSH keep their direct formatting.
 */
export function readResolvedLegacyDocModel(bytes: Uint8Array): LegacyDocModel {
  const doc = openLegacyDoc(bytes)
  const warnings: string[] = []
  const paragraphs = readResolvedFormatting(doc, warnings)
  return { fib: doc.fib, text: doc.text, paragraphs, warnings }
}

/** Stage 3c-1 layout: section geometry plus the header/footer stories of every section. */
export interface LegacyDocLayout {
  sections: SectionInfo[]
  /** one entry per section (even/odd/first header and footer stories) */
  headerFooter: SectionHeaderFooter[]
  /** non-fatal layout problems (missing header table, flattened columns, ...) */
  warnings: string[]
}

/**
 * Header/footer, note and text box stories live past the body in the same
 * text stream, so their formatting comes from a second formatting read over
 * the whole text (only when the document has such stories).
 */
function wholeTextSource(doc: LegacyDoc): FormattingSource {
  const total = doc.pieces.cps.at(-1) ?? 0
  return {
    ...doc,
    fib: { ...doc.fib, ccpText: total },
    rawText: decodeWordDocumentRange(doc.wordDocument, doc.pieces, 0, total),
  }
}

function storyFormatter(source: FormattingSource, styles: StyleSheet | undefined): StoryFormatter {
  const read = readFormattingDetailed(source, [])
  const fields = parseFields(source.rawText).simpleFields
  return {
    paragraphs(startCp, endCp, noteMarks = []) {
      const out: FormattedParagraph[] = []
      read.paragraphs.forEach((paragraph, index) => {
        if (paragraph.startCp < startCp || paragraph.endCp > endCp) return
        const raw = read.paragraphSprms[index]!
        const request = {
          startCp: paragraph.startCp,
          endCp: paragraph.endCp,
          styleIstd: raw.istd,
          papxSprms: raw.sprms,
        }
        out.push({
          startCp: paragraph.startCp,
          endCp: paragraph.endCp,
          text: paragraph.text,
          formatting: resolveParagraphFormattingFor(source, styles, request),
          runs: resolveRunsForRange(
            read,
            source,
            styles,
            request,
            [],
            [],
            fieldRangesWithin(fields, paragraph.startCp, paragraph.endCp),
            noteMarks,
          ),
          ...markFormattingEntry(resolveParagraphMarkFormatting(read, source, styles, request)),
        })
      })
      return out
    },
  }
}

function commentsOf(
  doc: LegacyDoc,
  formatter: StoryFormatter | undefined,
): {
  comments: LegacyComment[]
  warnings: string[]
} {
  if (doc.fib.ccpAtn === 0 && doc.fib.lcbPlcfandRef === 0) return { comments: [], warnings: [] }
  return parseComments(doc, formatter)
}

/** the formatting read shared by the header/footer, note, comment and text box stories */
interface Stories {
  source?: FormattingSource
  formatter?: StoryFormatter
  drawings: FloatingPicturesResult
  warnings: string[]
}

function readStories(doc: LegacyDoc, styles: StyleSheet | undefined): Stories {
  const { fib } = doc
  if (fib.ccpHdd + fib.ccpFtn + fib.ccpEdn + fib.ccpAtn + fib.ccpTxbx + fib.ccpHdrTxbx === 0)
    return { drawings: parseFloatingPictures(doc), warnings: [] }
  const source = wholeTextSource(doc)
  const formatter = storyFormatter(source, styles)
  const warnings: string[] = []
  const imagesIn = (startCp: number, length: number) => {
    if (length === 0) return []
    const result = parseInlineImages(source, { startCp, endCp: startCp + length })
    warnings.push(...result.warnings)
    return result.images.map((image) => ({ ...image, cp: image.cp - startCp }))
  }
  const textBoxStart = textBoxDocumentStart(fib)
  const textBoxes = parseTextBoxStories(doc, formatter, {
    body: imagesIn(textBoxStart, fib.ccpTxbx),
    header: imagesIn(textBoxStart + fib.ccpTxbx, fib.ccpHdrTxbx),
  })
  return { source, formatter, drawings: parseFloatingPictures(doc, textBoxes), warnings }
}

function layoutOf(doc: LegacyDoc, stories: Stories): LegacyDocLayout {
  const sections = parseSections(doc)
  const headerStart = doc.fib.ccpText + doc.fib.ccpFtn
  const images =
    stories.source && doc.fib.ccpHdd > 0
      ? parseInlineImages(stories.source, {
          startCp: headerStart,
          endCp: headerStart + doc.fib.ccpHdd,
        })
      : undefined
  const headerFooter = parseHeaderStories(
    doc,
    sections,
    stories.formatter,
    stories.drawings.header,
    images?.images.map((image) => ({ ...image, cp: image.cp - headerStart })),
  )
  return {
    sections,
    headerFooter: headerFooter.stories,
    warnings: [
      ...sections.flatMap((section) => section.warnings),
      ...headerFooter.warnings,
      ...(images?.warnings ?? []),
      ...stories.warnings,
    ],
  }
}

function bodyOf(
  doc: LegacyDoc,
  read: FormattingRead,
  styles: StyleSheet | undefined,
  warnings: string[],
  stories: Stories,
): LegacyDocBody {
  const images = parseInlineImages(doc)
  const lists = parseListTables(doc.fib, doc.table)
  const fields = parseFields(doc.rawText)
  const notes: LegacyDocNotes = parseNotes(doc, stories.formatter)
  const comments = commentsOf(doc, stories.formatter)
  const body = buildLegacyDocBody(doc, read, {
    styles,
    images: images.images,
    lists,
    fields,
    drawings: stories.drawings.body,
    noteMarks: notes.marks,
    comments: comments.comments,
  })
  return {
    blocks: body.blocks,
    warnings: [
      ...new Set([
        ...warnings,
        ...body.warnings,
        ...images.warnings,
        ...lists.warnings,
        ...fields.warnings,
        ...stories.drawings.warnings,
        ...stories.warnings,
        ...notes.warnings,
        ...comments.warnings,
      ]),
    ],
    ...(body.lists ? { lists: body.lists } : {}),
    ...(notes.footnotes.length + notes.endnotes.length > 0 ? { notes } : {}),
    ...(body.comments ? { comments: body.comments } : {}),
  }
}

/**
 * Read the reconstructed body of a .doc: paragraphs, real tables and inline
 * PNG/JPEG pictures in CP order (stages 3c-2/3c-3). Cell character formatting
 * is resolved from the stylesheet when the document has one; cell paragraph
 * formatting stays plain.
 */
export function readLegacyDocBody(bytes: Uint8Array): LegacyDocBody {
  const doc = openLegacyDoc(bytes)
  const warnings: string[] = []
  const read = readFormattingDetailed(doc, warnings)
  const styles = readStyleSheet(doc, warnings)
  return bodyOf(doc, read, styles, warnings, readStories(doc, styles))
}

/** Read the section geometry and header/footer stories of a .doc. */
export function readLegacyDocLayout(bytes: Uint8Array): LegacyDocLayout {
  const doc = openLegacyDoc(bytes)
  return layoutOf(doc, readStories(doc, readStyleSheet(doc, [])))
}

/** Everything the .docx writer needs, read from one parse of the file. */
export interface LegacyDocRead {
  model: LegacyDocModel
  layout: LegacyDocLayout
  body: LegacyDocBody
  /** the document has an STSH and the model carries style-resolved formatting */
  stylesResolved: boolean
}

/** Open a .doc once and read the model, the layout and the body from that parse. */
export function readLegacyDoc(bytes: Uint8Array): LegacyDocRead {
  const doc = openLegacyDoc(bytes)
  const warnings: string[] = []
  const read = readFormattingDetailed(doc, warnings)
  const styles = readStyleSheet(doc, warnings)
  const paragraphs = styles ? resolveFormatting(read, doc, styles, warnings) : read.paragraphs
  const model: LegacyDocModel = { fib: doc.fib, text: doc.text, paragraphs, warnings }
  const stories = readStories(doc, styles)
  return {
    model,
    layout: layoutOf(doc, stories),
    body: bodyOf(doc, read, styles, [], stories),
    stylesResolved: styles !== undefined,
  }
}

/* stage 3a/3c-1/3c-2: .docx writer */
export {
  conversionWarnings,
  convertLegacyDocToDocx,
  generateBodyBlocksXml,
  generateBodyParagraphsXml,
  generateTableXml,
  toDocxRun,
  toParaFormat,
  writeLegacyDocx,
} from './docx-writer'
