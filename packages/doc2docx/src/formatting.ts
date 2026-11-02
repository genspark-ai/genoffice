/**
 * Stage 2: the per-run / per-paragraph formatting model on top of the stage-1
 * text reader. CHPX FKPs carry character properties per FC run; PAPX FKPs carry
 * paragraph properties keyed by paragraph start FC. Properties are read as
 * sprms and folded in file order; a missing sprm stays undefined (stage 3 adds
 * the style sheet that supplies the defaults).
 */
import { LegacyDocError } from './errors'
import type { FibInfo } from './fib'
import type { PieceTable } from './clx'
import { readChpxFkp, readPapxFkp, type PapxEntry } from './fkp'
import { parsePlcBte } from './plcf'
import { iterateSprms, readSprmI16, readSprmU16, readSprmU8, readSprmU32, type Sprm } from './sprm'
import { createWordControlMapper, mapWordControlChars } from './text'

/** underline kinds of sprmCKul; 'inherit'/'invert' are the toggle-style operands */
export type UnderlineKind =
  | 'none'
  | 'single'
  | 'words'
  | 'double'
  | 'dotted'
  | 'thick'
  | 'dash'
  | 'dotDash'
  | 'dotDotDash'
  | 'wave'
  | `other:${number}`
  | 'inherit'
  | 'invert'

export interface CharacterFormatting {
  bold?: boolean
  italic?: boolean
  strike?: boolean
  outline?: boolean
  shadow?: boolean
  smallCaps?: boolean
  caps?: boolean
  hidden?: boolean
  /** font size in half-points (sprmCHps: 28 = 14pt) */
  fontSizeHalfPoints?: number
  underline?: UnderlineKind
  /** Word's color index (sprmCIco), resolved by palette in stage 3 */
  colorIndex?: number
  verticalAlign?: 'baseline' | 'superscript' | 'subscript'
  /** highlight color index (sprmCHighlight) */
  highlight?: number
  /** font indices (sprmCRgFtc0/1/2), resolved by the font table in stage 3 */
  ftcAscii?: number
  ftcHAnsi?: number
  ftcEastAsia?: number
  /** character style index (sprmCIstd) */
  styleIndex?: number
  /** raised/lowered position in half-points (sprmCHpsPos) */
  positionHalfPoints?: number
  /** minimum kerning in half-points (sprmCHpsKern) */
  kerningHalfPoints?: number
}

export interface ParagraphFormatting {
  /** paragraph style index (istd from the PapxInFkp) */
  styleIndex?: number
  alignment?: 'left' | 'center' | 'right' | 'justify' | 'distribute'
  /** sprmPFKeep: keep lines together */
  keepLines?: boolean
  /** sprmPFKeepFollow: keep with next */
  keepNext?: boolean
  pageBreakBefore?: boolean
  indentLeftTwips?: number
  indentRightTwips?: number
  indentFirstLineTwips?: number
  spacingBeforeTwips?: number
  spacingAfterTwips?: number
  /** sprmPDyaLine: dyaLine in twips plus the multiple-line-spacing flag */
  lineSpacing?: { dyaLine: number; multiple: boolean }
  outlineLevel?: number
  listLevel?: number
  listId?: number
  inTable?: boolean
  tableRowEnd?: boolean
}

export interface FormattedRun {
  text: string
  formatting: CharacterFormatting
}

export interface FormattedParagraph {
  /** CP of the first character of the paragraph text (excluding its paragraph mark) */
  startCp: number
  /** CP one past the last text character (the paragraph mark is not included) */
  endCp: number
  text: string
  formatting: ParagraphFormatting
  /** text runs; concatenating their text reproduces `text` */
  runs: FormattedRun[]
}

/** the parts of a parsed .doc the formatting layer reads */
export interface FormattingSource {
  fib: FibInfo
  pieces: PieceTable
  wordDocument: Uint8Array
  table: Uint8Array
  /** body text headline (stage-1 `rawText`, control marks intact) */
  rawText: string
  /** other OLE2 streams (the Data stream for sprmPHugePapx); optional in tests */
  stream?(name: string): Uint8Array | undefined
}

/* Character sprms (MS-DOC 2.6.x) */
const SPRM_CF_BOLD = 0x0835
const SPRM_CF_ITALIC = 0x0836
const SPRM_CF_STRIKE = 0x0837
const SPRM_CF_OUTLINE = 0x0838
const SPRM_CF_SHADOW = 0x0839
const SPRM_CF_SMALL_CAPS = 0x083a
const SPRM_CF_CAPS = 0x083b
const SPRM_CF_VANISH = 0x083c
const SPRM_C_KUL = 0x2a3e
const SPRM_C_ICO = 0x2a42
const SPRM_C_HPS = 0x4a43
const SPRM_C_ISS = 0x2a48
const SPRM_C_HIGHLIGHT = 0x2a0c
const SPRM_C_RG_FTC0 = 0x4a4f
const SPRM_C_RG_FTC1 = 0x4a50
const SPRM_C_RG_FTC2 = 0x4a51
const SPRM_C_ISTD = 0x4a30
const SPRM_C_HPS_POS = 0x4845
const SPRM_C_HPS_KERN = 0x484b

/* Paragraph sprms */
const SPRM_P_JC80 = 0x2403
const SPRM_P_JC = 0x2461
const SPRM_P_F_KEEP = 0x2405
const SPRM_P_F_KEEP_FOLLOW = 0x2406
const SPRM_P_F_PAGE_BREAK_BEFORE = 0x2407
const SPRM_P_ILVL = 0x260a
const SPRM_P_ILFO = 0x460b
const SPRM_P_DXA_RIGHT = 0x845d
const SPRM_P_DXA_RIGHT_OLD = 0x840e
const SPRM_P_DXA_LEFT = 0x845e
const SPRM_P_DXA_LEFT_OLD = 0x840f
const SPRM_P_DXA_LEFT1 = 0x8460
const SPRM_P_DXA_LEFT1_OLD = 0x8411
const SPRM_P_DYA_LINE = 0x6412
const SPRM_P_DYA_BEFORE = 0xa413
const SPRM_P_DYA_AFTER = 0xa414
const SPRM_P_OUT_LVL = 0x2640
const SPRM_P_F_IN_TABLE = 0x2416
const SPRM_P_F_TTP = 0x2417
const SPRM_P_HUGE_PAPX = 0x6645
const SPRM_P_HUGE_PAPX_2 = 0x6646

const TOGGLE_INHERIT = 0x80
const TOGGLE_INVERT = 0x81

const UNDERLINE_KINDS: Record<number, UnderlineKind> = {
  0: 'none',
  1: 'single',
  2: 'words',
  3: 'double',
  4: 'dotted',
  6: 'thick',
  7: 'dash',
  9: 'dotDash',
  10: 'dotDotDash',
  11: 'wave',
}

const ALIGNMENTS: Record<number, ParagraphFormatting['alignment']> = {
  0: 'left',
  1: 'center',
  2: 'right',
  3: 'justify',
  4: 'distribute',
}

const VERTICAL_ALIGNS: Record<number, CharacterFormatting['verticalAlign']> = {
  0: 'baseline',
  1: 'superscript',
  2: 'subscript',
}

/**
 * How ToggleOperand values 0x80/0x81 are applied. 'direct' is the stage-2
 * direct-formatting model: without a style sheet there is nothing to inherit
 * from, so both leave the field unset. 'inherit' is style resolution: 0x80
 * keeps the value inherited from the previous layer and 0x81 inverts it (the
 * base of the chain is off).
 */
export type ToggleMode = 'direct' | 'inherit'

/**
 * ToggleOperand: 0 = off, 1 = on, 0x80 = inherit, 0x81 = invert. The style
 * resolver uses 'inherit'; the direct-formatting model keeps 'direct'.
 */
function applyToggle(target: object, key: string, operand: Uint8Array, mode: ToggleMode): void {
  const value = operand[0]
  if (value === undefined) return
  const record = target as Record<string, unknown>
  if (value === TOGGLE_INHERIT) {
    if (mode === 'inherit') return // keep the inherited value
    delete record[key]
    return
  }
  if (value === TOGGLE_INVERT) {
    if (mode === 'inherit') record[key] = !(record[key] ?? false)
    else delete record[key]
    return
  }
  record[key] = value !== 0
}

function underlineKind(value: number): UnderlineKind {
  if (value === TOGGLE_INHERIT) return 'inherit'
  if (value === TOGGLE_INVERT) return 'invert'
  return UNDERLINE_KINDS[value] ?? `other:${value}`
}

function applyCharacterSprm(format: CharacterFormatting, sprm: Sprm, mode: ToggleMode): void {
  switch (sprm.opcode) {
    case SPRM_CF_BOLD:
      applyToggle(format, 'bold', sprm.operand, mode)
      break
    case SPRM_CF_ITALIC:
      applyToggle(format, 'italic', sprm.operand, mode)
      break
    case SPRM_CF_STRIKE:
      applyToggle(format, 'strike', sprm.operand, mode)
      break
    case SPRM_CF_OUTLINE:
      applyToggle(format, 'outline', sprm.operand, mode)
      break
    case SPRM_CF_SHADOW:
      applyToggle(format, 'shadow', sprm.operand, mode)
      break
    case SPRM_CF_SMALL_CAPS:
      applyToggle(format, 'smallCaps', sprm.operand, mode)
      break
    case SPRM_CF_CAPS:
      applyToggle(format, 'caps', sprm.operand, mode)
      break
    case SPRM_CF_VANISH:
      applyToggle(format, 'hidden', sprm.operand, mode)
      break
    case SPRM_C_KUL: {
      const value = readSprmU8(sprm.operand)
      if (value !== undefined) format.underline = underlineKind(value)
      break
    }
    case SPRM_C_ICO: {
      const value = readSprmU8(sprm.operand)
      if (value !== undefined) format.colorIndex = value
      break
    }
    case SPRM_C_HPS: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.fontSizeHalfPoints = value
      break
    }
    case SPRM_C_ISS: {
      const value = readSprmU8(sprm.operand)
      const alignment = value === undefined ? undefined : VERTICAL_ALIGNS[value]
      if (alignment !== undefined) format.verticalAlign = alignment
      break
    }
    case SPRM_C_HIGHLIGHT: {
      const value = readSprmU8(sprm.operand)
      if (value !== undefined) format.highlight = value
      break
    }
    case SPRM_C_RG_FTC0: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.ftcAscii = value
      break
    }
    case SPRM_C_RG_FTC1: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.ftcHAnsi = value
      break
    }
    case SPRM_C_RG_FTC2: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.ftcEastAsia = value
      break
    }
    case SPRM_C_ISTD: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.styleIndex = value
      break
    }
    case SPRM_C_HPS_POS: {
      const value = readSprmI16(sprm.operand)
      if (value !== undefined) format.positionHalfPoints = value
      break
    }
    case SPRM_C_HPS_KERN: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.kerningHalfPoints = value
      break
    }
    default:
      // unknown/unsupported character sprm: skipped by the operand-size rule
      break
  }
}

/** Apply a raw CHPX grpprl to an existing character format (used by the style resolver). */
export function applyCharacterGrpprl(
  format: CharacterFormatting,
  grpprl: Uint8Array,
  mode: ToggleMode = 'direct',
): void {
  for (const sprm of iterateSprms(grpprl)) applyCharacterSprm(format, sprm, mode)
}

function characterFormattingOf(grpprl: Uint8Array): CharacterFormatting {
  const format: CharacterFormatting = {}
  applyCharacterGrpprl(format, grpprl)
  return format
}

/** What paragraph-sprm application needs besides the format (huge PAPX reads the Data stream). */
export interface ParagraphContext {
  source: FormattingSource
  warnings: string[]
}

function applyParagraphSprm(
  format: ParagraphFormatting,
  sprm: Sprm,
  context: ParagraphContext,
  depth: number,
  mode: ToggleMode,
): void {
  switch (sprm.opcode) {
    case SPRM_P_JC80:
    case SPRM_P_JC: {
      const value = readSprmU8(sprm.operand)
      const alignment = value === undefined ? undefined : ALIGNMENTS[value]
      if (alignment !== undefined) format.alignment = alignment
      break
    }
    case SPRM_P_F_KEEP:
      applyToggle(format, 'keepLines', sprm.operand, mode)
      break
    case SPRM_P_F_KEEP_FOLLOW:
      applyToggle(format, 'keepNext', sprm.operand, mode)
      break
    case SPRM_P_F_PAGE_BREAK_BEFORE:
      applyToggle(format, 'pageBreakBefore', sprm.operand, mode)
      break
    case SPRM_P_ILVL: {
      const value = readSprmU8(sprm.operand)
      if (value !== undefined) format.listLevel = value
      break
    }
    case SPRM_P_ILFO: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.listId = value
      break
    }
    case SPRM_P_DXA_RIGHT:
    case SPRM_P_DXA_RIGHT_OLD: {
      const value = readSprmI16(sprm.operand)
      if (value !== undefined) format.indentRightTwips = value
      break
    }
    case SPRM_P_DXA_LEFT:
    case SPRM_P_DXA_LEFT_OLD: {
      const value = readSprmI16(sprm.operand)
      if (value !== undefined) format.indentLeftTwips = value
      break
    }
    case SPRM_P_DXA_LEFT1:
    case SPRM_P_DXA_LEFT1_OLD: {
      const value = readSprmI16(sprm.operand)
      if (value !== undefined) format.indentFirstLineTwips = value
      break
    }
    case SPRM_P_DYA_LINE: {
      const dyaLine = readSprmI16(sprm.operand)
      const multiple = readSprmU16(sprm.operand, 2)
      if (dyaLine !== undefined && multiple !== undefined)
        format.lineSpacing = { dyaLine, multiple: multiple !== 0 }
      break
    }
    case SPRM_P_DYA_BEFORE: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.spacingBeforeTwips = value
      break
    }
    case SPRM_P_DYA_AFTER: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.spacingAfterTwips = value
      break
    }
    case SPRM_P_OUT_LVL: {
      const value = readSprmU8(sprm.operand)
      if (value !== undefined) format.outlineLevel = value
      break
    }
    case SPRM_P_F_IN_TABLE:
      applyToggle(format, 'inTable', sprm.operand, mode)
      break
    case SPRM_P_F_TTP:
      applyToggle(format, 'tableRowEnd', sprm.operand, mode)
      break
    case SPRM_P_HUGE_PAPX:
    case SPRM_P_HUGE_PAPX_2:
      applyHugePapx(format, sprm, context, depth, mode)
      break
    default:
      // unknown/unsupported paragraph sprm: skipped by the operand-size rule
      break
  }
}

const MAX_HUGE_PAPX_DEPTH = 4

/**
 * sprmPHugePapx: the rest of the paragraph's properties live in the Data
 * stream at the 4-byte FC operand: a signed 16-bit cb followed by the grpprl.
 * A missing Data stream keeps the properties read so far and records a warning.
 */
function applyHugePapx(
  format: ParagraphFormatting,
  sprm: Sprm,
  context: ParagraphContext,
  depth: number,
  mode: ToggleMode,
): void {
  if (depth >= MAX_HUGE_PAPX_DEPTH) {
    context.warnings.push('nested sprmPHugePapx records were ignored (too deep)')
    return
  }
  const fc = readSprmU32(sprm.operand)
  const data = context.source.stream?.('Data')
  if (fc === undefined || !data) {
    context.warnings.push(
      'a paragraph stores extra properties in the Data stream, which this file does not have; they were ignored',
    )
    return
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
  if (fc + 2 > data.length) {
    context.warnings.push(
      'sprmPHugePapx points outside the Data stream; the extra paragraph properties were ignored',
    )
    return
  }
  const cb = view.getInt16(fc, true)
  if (cb <= 0 || fc + 2 + cb > data.length) {
    context.warnings.push(
      'sprmPHugePapx has an invalid property block length; the extra paragraph properties were ignored',
    )
    return
  }
  const grpprl = data.slice(fc + 2, fc + 2 + cb)
  applyParagraphGrpprl(format, grpprl, context, mode, depth + 1)
}

/** Apply a raw PAPX grpprl to an existing paragraph format (used by the style resolver). */
export function applyParagraphGrpprl(
  format: ParagraphFormatting,
  grpprl: Uint8Array,
  context: ParagraphContext,
  mode: ToggleMode = 'direct',
  depth = 0,
): void {
  for (const sprm of iterateSprms(grpprl)) applyParagraphSprm(format, sprm, context, depth, mode)
}

function paragraphFormattingOf(entry: PapxEntry, context: ParagraphContext): ParagraphFormatting {
  const format: ParagraphFormatting = { styleIndex: entry.istd }
  applyParagraphGrpprl(format, entry.grpprl, context)
  return format
}

/** FC -> CP resolution result: a CP, an unaligned byte offset, or a byte outside the text pieces */
export type FcResolution = { cp: number } | 'unaligned' | 'outside'

/** Map a character-start FC to its CP; a character-end FC (piece end) maps to the piece's end CP. */
export function cpOfFc(pieceTable: PieceTable, fc: number): FcResolution {
  for (const piece of pieceTable.pieces) {
    const width = piece.compressed ? 1 : 2
    const byteLength = (piece.cpEnd - piece.cpStart) * width
    if (fc >= piece.fc && fc < piece.fc + byteLength) {
      const delta = fc - piece.fc
      if (delta % width !== 0) return 'unaligned'
      return { cp: piece.cpStart + delta / width }
    }
    if (fc === piece.fc + byteLength) return { cp: piece.cpEnd }
  }
  return 'outside'
}

/** Inverse of cpOfFc: the byte offset of a CP in the WordDocument stream (undefined outside the pieces). */
export function fcOfCp(pieceTable: PieceTable, cp: number): number | undefined {
  for (const piece of pieceTable.pieces) {
    if (cp < piece.cpStart || cp > piece.cpEnd) continue
    if (cp === piece.cpEnd && piece !== pieceTable.pieces[pieceTable.pieces.length - 1]) continue
    const width = piece.compressed ? 1 : 2
    return piece.fc + (cp - piece.cpStart) * width
  }
  return undefined
}

const SENTINEL_FC = 0xffffffff

/** An FC boundary from a formatting table to a body CP, with end clamping for FKPs. */
function cpAtFc(source: FormattingSource, fc: number, what: string): number {
  const ccpText = source.fib.ccpText
  if (fc === SENTINEL_FC) return ccpText
  const resolved = cpOfFc(source.pieces, fc)
  if (resolved === 'unaligned')
    throw new LegacyDocError(
      'bad_fkp',
      `${what} at byte ${fc} does not land on a character boundary`,
    )
  if (resolved === 'outside') {
    const pieces = source.pieces.pieces
    const first = pieces[0]
    const last = pieces[pieces.length - 1]
    if (first && fc < first.fc) return 0
    if (last) {
      const width = last.compressed ? 1 : 2
      const lastEnd = last.fc + (last.cpEnd - last.cpStart) * width
      if (fc >= lastEnd) return ccpText
    }
    throw new LegacyDocError('bad_fkp', `${what} at byte ${fc} does not map to the document text`)
  }
  return Math.min(resolved.cp, ccpText)
}

const NO_SPRMS = new Uint8Array(0)

interface CpCharacterRun {
  cpStart: number
  cpEnd: number
  formatting: CharacterFormatting
  /** the raw CHPX grpprl, kept for style resolution */
  sprms: Uint8Array
}

function collectCharacterRuns(source: FormattingSource): CpCharacterRun[] {
  const plc = parsePlcBte(source.table, source.fib.fcPlcfbteChpx, source.fib.lcbPlcfbteChpx)
  const runs: CpCharacterRun[] = []
  for (const page of plc.pages) {
    for (const run of readChpxFkp(source.wordDocument, page)) {
      if (run.fcEnd <= run.fcStart) continue
      const cpStart = cpAtFc(source, run.fcStart, 'a CHPX run start')
      const cpEnd = Math.min(cpAtFc(source, run.fcEnd, 'a CHPX run end'), source.fib.ccpText)
      if (cpEnd <= cpStart) continue
      runs.push({
        cpStart,
        cpEnd,
        formatting: characterFormattingOf(run.grpprl),
        sprms: run.grpprl,
      })
    }
  }
  runs.sort((left, right) => left.cpStart - right.cpStart)
  return runs
}

interface ParagraphSpan {
  startCp: number
  endCp: number
  text: string
}

/** split the raw body text on paragraph marks, exactly like stage 1 (trailing empty paragraph included) */
function paragraphSpans(rawText: string): ParagraphSpan[] {
  const spans: ParagraphSpan[] = []
  let start = 0
  for (let index = 0; index < rawText.length; index++) {
    if (rawText[index] !== '\r') continue
    spans.push({
      startCp: start,
      endCp: index,
      text: mapWordControlChars(rawText.slice(start, index)),
    })
    start = index + 1
  }
  spans.push({
    startCp: start,
    endCp: rawText.length,
    text: mapWordControlChars(rawText.slice(start)),
  })
  return spans
}

interface CollectParagraphResult {
  formats: Map<number, ParagraphFormatting>
  raw: Map<number, RawParagraphSprms>
}

function collectParagraphFormatting(
  source: FormattingSource,
  spans: ParagraphSpan[],
  warnings: string[],
): CollectParagraphResult {
  const context: ParagraphContext = { source, warnings }
  const formats = new Map<number, ParagraphFormatting>()
  const raw = new Map<number, RawParagraphSprms>()
  const spanIndexByStartCp = new Map<number, number>()
  spans.forEach((span, index) => spanIndexByStartCp.set(span.startCp, index))
  const plc = parsePlcBte(source.table, source.fib.fcPlcfbtePapx, source.fib.lcbPlcfbtePapx)
  for (const page of plc.pages) {
    for (const entry of readPapxFkp(source.wordDocument, page)) {
      const cp = cpAtFc(source, entry.fcStart, 'a PAPX paragraph start')
      const index = spanIndexByStartCp.get(cp)
      if (index === undefined) {
        warnings.push(
          `a paragraph formatting record starts at CP ${cp}, which is not a paragraph start; it was ignored`,
        )
        continue
      }
      formats.set(index, paragraphFormattingOf(entry, context))
      raw.set(index, { istd: entry.istd, sprms: entry.grpprl })
    }
  }
  return { formats, raw }
}

function sameFormatting(left: object, right: object): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function buildRuns(
  span: ParagraphSpan,
  characterRuns: CpCharacterRun[],
  rawText: string,
): FormattedRun[] {
  const boundaries = new Set<number>([span.startCp, span.endCp])
  for (const run of characterRuns) {
    if (run.cpEnd <= span.startCp || run.cpStart >= span.endCp) continue
    boundaries.add(Math.max(run.cpStart, span.startCp))
    boundaries.add(Math.min(run.cpEnd, span.endCp))
  }
  const points = [...boundaries].sort((left, right) => left - right)
  const mapper = createWordControlMapper()
  const formatted: FormattedRun[] = []
  for (let index = 0; index + 1 < points.length; index++) {
    const segmentStart = points[index]!
    const segmentEnd = points[index + 1]!
    if (segmentEnd <= segmentStart) continue
    const covering = characterRuns.find(
      (run) => run.cpStart <= segmentStart && run.cpEnd >= segmentEnd,
    )
    const formatting = covering?.formatting ?? {}
    const text = mapper.map(rawText.slice(segmentStart, segmentEnd))
    if (text === '') continue
    const previous = formatted[formatted.length - 1]
    if (previous && sameFormatting(previous.formatting, formatting)) previous.text += text
    else formatted.push({ text, formatting })
  }
  return formatted
}

/** The PAPX istd and grpprl behind one paragraph (defaults when the paragraph has no PAPX). */
export interface RawParagraphSprms {
  istd: number
  sprms: Uint8Array
}

/** One raw CHPX run: the CP range and its grpprl. */
export interface RawCharacterRun {
  cpStart: number
  cpEnd: number
  sprms: Uint8Array
}

/** The stage-2 formatting read plus the raw sprms the style resolver replays. */
export interface FormattingRead {
  paragraphs: FormattedParagraph[]
  /** one entry per paragraph, aligned with `paragraphs` */
  paragraphSprms: RawParagraphSprms[]
  /** the CHPX runs in CP order */
  characterRuns: RawCharacterRun[]
}

/**
 * Read the formatting model of a stage-1 document: one FormattedParagraph per
 * stage-1 paragraph (trailing empty paragraph included), each with paragraph
 * properties and character runs covering its text, plus the raw PAPX/CHPX
 * sprms for the style resolver. Warnings (e.g. a paragraph whose properties
 * live in a missing Data stream) are appended to `warnings`.
 */
export function readFormattingDetailed(
  source: FormattingSource,
  warnings: string[] = [],
): FormattingRead {
  const spans = paragraphSpans(source.rawText)
  const characterRuns = collectCharacterRuns(source)
  const collected = collectParagraphFormatting(source, spans, warnings)
  const paragraphs = spans.map((span, index) => ({
    startCp: span.startCp,
    endCp: span.endCp,
    text: span.text,
    formatting: collected.formats.get(index) ?? {},
    runs: buildRuns(span, characterRuns, source.rawText),
  }))
  return {
    paragraphs,
    paragraphSprms: spans.map(
      (_, index) => collected.raw.get(index) ?? { istd: 0, sprms: NO_SPRMS },
    ),
    characterRuns: characterRuns.map((run) => ({
      cpStart: run.cpStart,
      cpEnd: run.cpEnd,
      sprms: run.sprms,
    })),
  }
}

/** Read just the stage-2 direct-formatting model (the stage-2/3a public read). */
export function readFormatting(
  source: FormattingSource,
  warnings: string[] = [],
): FormattedParagraph[] {
  return readFormattingDetailed(source, warnings).paragraphs
}
