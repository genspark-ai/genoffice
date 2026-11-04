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
import {
  iterateSprms,
  iterateSprmsLenient,
  readSprmI16,
  readSprmU16,
  readSprmU8,
  readSprmU32,
  type Sprm,
} from './sprm'
import { createWordControlMapper, mapWordControlChars } from './text'
import type { FontNames } from './fonts'

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
  /** sprmCFBiDi: right-to-left run */
  rtl?: boolean
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
  /** sprmCCv RGB color as 'RRGGBB', or 'auto'; wins over colorIndex when both are set */
  colorRgb?: string
  verticalAlign?: 'baseline' | 'superscript' | 'subscript'
  /** highlight color index (sprmCHighlight) */
  highlight?: number
  /** font indices (sprmCRgFtc0 ascii, 1 East Asian, 2 other/hAnsi, sprmCFtcBi), resolved by the font table */
  ftcAscii?: number
  ftcHAnsi?: number
  ftcEastAsia?: number
  ftcBi?: number
  /** font names resolved from the ftc indices through the font table */
  fontNames?: FontNames
  /** character style index (sprmCIstd) */
  styleIndex?: number
  /** raised/lowered position in half-points (sprmCHpsPos) */
  positionHalfPoints?: number
  /** minimum kerning in half-points (sprmCHpsKern) */
  kerningHalfPoints?: number
  /** expanded/condensed character spacing in twips (sprmCDxaSpace) */
  spacingTwips?: number
  /** horizontal character scale in percent (sprmCCharScale) */
  scalePercent?: number
  /** Windows LCIDs (sprmCRgLid0/1): the Latin and the East Asian language of the run */
  languageId?: number
  eastAsianLanguageId?: number
}

export interface TabStop {
  positionTwips: number
  alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar'
  leader?: 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot'
}

export interface ParagraphBorder {
  /** line width in eighths of a point (dptLineWidth) */
  widthEighths: number
  /** gap between the text and the line in points (dptSpace) */
  spacePoints: number
  /** 'RRGGBB'; undefined = automatic */
  color?: string
}

export type ParagraphBorderSide = 'top' | 'left' | 'bottom' | 'right'

/** character-unit indents (sprmPDxcLeft/Left1/Right), in hundredths of a character */
export interface CharacterIndents {
  left?: number
  right?: number
  firstLine?: number
}

/** Positioned paragraph frame (sprmPPc / sprmPDxaAbs / sprmPDyaAbs ...), written as w:framePr. */
export interface ParagraphFrame {
  /** sprmPDxaAbs: twips, or the alignment codes -4 center, -8 right, -12 inside, -16 outside */
  xTwips?: number
  /** sprmPDyaAbs: twips, or -4 top, -8 center, -12 bottom, -16 inside, -20 outside */
  yTwips?: number
  /** sprmPDxaWidth; absent or 0 = auto width */
  widthTwips?: number
  /** sprmPDyaHeight (bit 15 = exact); absent or 0 = auto height */
  heightTwips?: number
  heightExact?: boolean
  hAnchor?: 'text' | 'margin' | 'page'
  vAnchor?: 'text' | 'margin' | 'page'
  /** sprmPWr (ST_Wrap order) */
  wrap?: 'auto' | 'notBeside' | 'around' | 'none' | 'tight' | 'through'
  hSpaceTwips?: number
  vSpaceTwips?: number
}

/** wrap/clearance alone (e.g. inherited from a style) does not make the paragraph a frame */
export function isPositionedFrame(frame: ParagraphFrame | undefined): frame is ParagraphFrame {
  return (
    frame !== undefined &&
    (frame.xTwips !== undefined ||
      frame.yTwips !== undefined ||
      frame.hAnchor !== undefined ||
      Boolean(frame.widthTwips) ||
      Boolean(frame.heightTwips))
  )
}

export interface ParagraphFormatting {
  /** paragraph style index (istd from the PapxInFkp) */
  styleIndex?: number
  alignment?: 'left' | 'center' | 'right' | 'justify' | 'distribute'
  /** the alignment came from sprmPJc (logical, flips in a bidi paragraph) rather than sprmPJc80 (physical) */
  alignmentLogical?: boolean
  /** sprmPFBiDi: right-to-left paragraph */
  bidi?: boolean
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
  /** custom tab stops after applying sprmPChgTabs / sprmPChgTabsPapx, sorted by position */
  tabStops?: TabStop[]
  /** paragraph borders (sprmPBrcXxx80 / sprmPBrcXxx); a 'none' border removes the side */
  borders?: Partial<Record<ParagraphBorderSide, ParagraphBorder>>
  charIndents?: CharacterIndents
  /** set when any frame sprm is present; Word treats such a paragraph as positioned */
  frame?: ParagraphFrame
  /** sprmPShd80 / sprmPShd: the paragraph background as OOXML writes it */
  shading?: Shading
}

/** A Shd / Shd80 as OOXML writes it: w:val pattern, w:color (fore), w:fill (back). */
export interface Shading {
  pattern: string
  color?: string
  fill?: string
}

/** Word color index (sprmCIco 1..16) -> 'RRGGBB' */
export const COLOR_BY_INDEX: Record<number, string> = {
  1: '000000',
  2: '0000FF',
  3: '00FFFF',
  4: '00FF00',
  5: 'FF00FF',
  6: 'FF0000',
  7: 'FFFF00',
  8: 'FFFFFF',
  9: '000080',
  10: '008080',
  11: '008000',
  12: '800080',
  13: '800000',
  14: '808000',
  15: '808080',
  16: 'C0C0C0',
}

/** A run that belongs to a hyperlink; the writer emits a w:hyperlink around it. */
export interface FormattedLink {
  /** external URL, or '#bookmark' for an internal anchor */
  href: string
  tooltip?: string
  /** written without the Hyperlink character style (TOC entries) */
  plain?: boolean
}

/** A comment anchor emitted between runs: range start, range end or the reference mark. */
export interface CommentMark {
  kind: 'start' | 'end' | 'reference'
  id: number
}

export interface FormattedRun {
  text: string
  formatting: CharacterFormatting
  /** set when the run's text sits inside a hyperlink field result */
  link?: FormattedLink
  /** field instruction when the run is the cached result of a live field (PAGE, NUMPAGES, ...) */
  field?: string
  /** the run is a footnote/endnote mark (text holds the custom mark, if any) */
  note?: NoteMark
  /** an empty run standing for comment anchors at this position */
  commentMarks?: CommentMark[]
  /** an empty run standing for the inline picture whose \x01 character sits at this CP */
  pictureCp?: number
}

/** A footnote/endnote mark character in the text: a body reference or the note's own mark. */
export interface NoteMark {
  cp: number
  kind: 'footnote' | 'endnote'
  /** 1-based note id (w:id in the written parts) */
  id: number
  /** reference in the body, or the mark that opens the note's own text */
  role: 'reference' | 'self'
  /** nAuto 0: the character at cp is the custom mark itself (auto marks are 0x02) */
  custom: boolean
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
  /** formatting of the paragraph mark (the CHPX covering the \r); sizes empty paragraphs */
  markFormatting?: CharacterFormatting
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
const SPRM_C_FTC_BI = 0x4a5e
const SPRM_C_ISTD = 0x4a30
const SPRM_C_HPS_POS = 0x4845
const SPRM_C_HPS_KERN = 0x484b
const SPRM_C_DXA_SPACE = 0x8840
const SPRM_C_CHAR_SCALE = 0x4852
const SPRM_C_CV = 0x6870
const SPRM_C_LID0_80 = 0x486d
const SPRM_C_LID1_80 = 0x486e
const SPRM_C_LID0 = 0x4873
const SPRM_C_LID1 = 0x4874

/* Paragraph sprms */
const SPRM_P_JC80 = 0x2403
const SPRM_P_JC = 0x2461
const SPRM_P_F_BIDI = 0x2441
const SPRM_C_F_BIDI = 0x085a
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
const SPRM_P_CHG_TABS_PAPX = 0xc60d
const SPRM_P_CHG_TABS = 0xc615
const SPRM_P_DXC_RIGHT = 0x4455
const SPRM_P_SHD80 = 0x442d
const SPRM_P_SHD = 0xc64d
const SPRM_P_DXC_LEFT = 0x4456
const SPRM_P_DXC_LEFT1 = 0x4457
const SPRM_P_DXA_ABS = 0x8418
const SPRM_P_DYA_ABS = 0x8419
const SPRM_P_DXA_WIDTH = 0x841a
const SPRM_P_PC = 0x261b
const SPRM_P_WR = 0x2423
const SPRM_P_DYA_HEIGHT = 0x442b
const SPRM_P_DYA_FROM_TEXT = 0x842e
const SPRM_P_DXA_FROM_TEXT = 0x842f
const FRAME_H_ANCHORS = ['text', 'margin', 'page'] as const
const FRAME_V_ANCHORS = ['margin', 'page', 'text'] as const
const FRAME_WRAPS = ['auto', 'notBeside', 'around', 'none', 'tight', 'through'] as const
/** Brc80 paragraph borders (top, left, bottom, right) and their 8-byte Brc twins */
const BORDER_SIDES_80: Record<number, ParagraphBorderSide> = {
  0x6424: 'top',
  0x6425: 'left',
  0x6426: 'bottom',
  0x6427: 'right',
}
const BORDER_SIDES: Record<number, ParagraphBorderSide> = {
  0xc64e: 'top',
  0xc64f: 'left',
  0xc650: 'bottom',
  0xc651: 'right',
}
const TAB_ALIGNMENTS: TabStop['alignment'][] = ['left', 'center', 'right', 'decimal', 'bar']
const TAB_LEADERS: Array<TabStop['leader']> = [
  undefined,
  'dot',
  'hyphen',
  'underscore',
  'heavy',
  'middleDot',
]

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
    case SPRM_C_F_BIDI:
      applyToggle(format, 'rtl', sprm.operand, mode)
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
      if (value !== undefined) format.ftcEastAsia = value
      break
    }
    case SPRM_C_RG_FTC2: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.ftcHAnsi = value
      break
    }
    case SPRM_C_FTC_BI: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.ftcBi = value
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
    case SPRM_C_DXA_SPACE: {
      const value = readSprmI16(sprm.operand)
      if (value !== undefined) format.spacingTwips = value
      break
    }
    case SPRM_C_CHAR_SCALE: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined && value > 0) format.scalePercent = value
      break
    }
    case SPRM_C_CV: {
      const color = colorRefOf(sprm.operand, 0)
      if (color !== undefined) format.colorRgb = color
      break
    }
    case SPRM_C_LID0_80:
    case SPRM_C_LID0: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.languageId = value
      break
    }
    case SPRM_C_LID1_80:
    case SPRM_C_LID1: {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) format.eastAsianLanguageId = value
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

/** Direct character formatting of one CHPX grpprl (no style resolution). */
export function characterFormattingOf(grpprl: Uint8Array): CharacterFormatting {
  const format: CharacterFormatting = {}
  applyCharacterGrpprl(format, grpprl)
  return format
}

/** a COLORREF (red, green, blue, fAuto) as 'RRGGBB' or 'auto' */
function colorRefOf(bytes: Uint8Array, offset: number): string | undefined {
  if (offset + 4 > bytes.length) return undefined
  if (bytes[offset + 3] === 0xff) return 'auto'
  return [bytes[offset]!, bytes[offset + 1]!, bytes[offset + 2]!]
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()
}

/** MS-DOC Ipat -> ST_Shd (0 = clear/auto, 1 = solid foreground) */
const SHADING_PATTERNS: Record<number, string> = {
  0: 'clear',
  1: 'solid',
  2: 'pct5',
  3: 'pct10',
  4: 'pct20',
  5: 'pct25',
  6: 'pct30',
  7: 'pct40',
  8: 'pct50',
  9: 'pct60',
  10: 'pct70',
  11: 'pct75',
  12: 'pct80',
  13: 'pct90',
  14: 'horzStripe',
  15: 'vertStripe',
  16: 'reverseDiagStripe',
  17: 'diagStripe',
  18: 'horzCross',
  19: 'diagCross',
  20: 'thinHorzStripe',
  21: 'thinVertStripe',
  22: 'thinReverseDiagStripe',
  23: 'thinDiagStripe',
  24: 'thinHorzCross',
  25: 'thinDiagCross',
  35: 'pct5',
  36: 'pct10',
  37: 'pct12',
  38: 'pct15',
  39: 'pct20',
  40: 'pct20',
  41: 'pct25',
  42: 'pct35',
  43: 'pct35',
  44: 'pct37',
  45: 'pct45',
  46: 'pct45',
  47: 'pct50',
  48: 'pct50',
  49: 'pct55',
  50: 'pct60',
  51: 'pct62',
  52: 'pct65',
  53: 'pct65',
  54: 'pct70',
  55: 'pct75',
  56: 'pct80',
  57: 'pct85',
  58: 'pct87',
  59: 'pct90',
  60: 'pct95',
  61: 'pct95',
  62: 'pct95',
}

function shadingFrom(
  pattern: number,
  fore: string | undefined,
  back: string | undefined,
): Shading | undefined {
  const color = fore !== undefined && fore !== 'auto' ? fore : undefined
  const fill = back !== undefined && back !== 'auto' ? back : undefined
  if (color === undefined && fill === undefined) return undefined
  return {
    pattern: SHADING_PATTERNS[pattern] ?? 'clear',
    ...(color !== undefined ? { color } : {}),
    ...(fill !== undefined ? { fill } : {}),
  }
}

/** a Shd80 (2 bytes: icoFore 5 bits, icoBack 5 bits, ipat 6 bits); undefined = no shading */
export function shading80At(bytes: Uint8Array, offset: number): Shading | undefined {
  if (offset + 2 > bytes.length) return undefined
  const value = bytes[offset]! | (bytes[offset + 1]! << 8)
  if (value === 0xffff) return undefined
  return shadingFrom(value >> 10, COLOR_BY_INDEX[value & 0x1f], COLOR_BY_INDEX[(value >> 5) & 0x1f])
}

/** a Shd (10 bytes: cvFore, cvBack, ipat); undefined = no shading */
export function shadingAt(bytes: Uint8Array, offset: number): Shading | undefined {
  if (offset + 10 > bytes.length) return undefined
  return shadingFrom(
    bytes[offset + 8]! | (bytes[offset + 9]! << 8),
    colorRefOf(bytes, offset),
    colorRefOf(bytes, offset + 4),
  )
}

/** the colour Word paints behind a shaded area (a solid pattern paints the foreground) */
export function shadingFillOf(shading: Shading): string | undefined {
  return shading.pattern === 'solid' ? (shading.color ?? shading.fill) : shading.fill
}

/** a Brc80 (4 bytes) or Brc (8 bytes); undefined when the side has no border */
function borderOf(operand: Uint8Array, wide: boolean): ParagraphBorder | undefined | null {
  if (operand.length < (wide ? 8 : 4)) return undefined
  const base = wide ? 4 : 0
  const widthEighths = operand[base]!
  const brcType = operand[base + 1]!
  if (brcType === 0 || brcType === 0xff || widthEighths === 0) return null
  const spacePoints = operand[base + 2 + (wide ? 0 : 1)]! & 0x1f
  const color = wide ? colorRefOf(operand, 0) : COLOR_BY_INDEX[operand[2]!]
  return {
    widthEighths,
    spacePoints,
    ...(color !== undefined && color !== 'auto' ? { color } : {}),
  }
}

function setBorder(
  format: ParagraphFormatting,
  side: ParagraphBorderSide,
  operand: Uint8Array,
  wide: boolean,
): void {
  const border = borderOf(operand, wide)
  if (border === undefined) return
  const borders = { ...format.borders }
  if (border === null) delete borders[side]
  else borders[side] = border
  if (Object.keys(borders).length === 0) delete format.borders
  else format.borders = borders
}

/**
 * sprmPChgTabsPapx / sprmPChgTabs: delete the listed positions from the
 * inherited stops, then add the new ones (dxa + TBD: jc bits 0-2, leader 3-5).
 * sprmPChgTabs also carries a close tolerance per deleted stop.
 */
function applyTabChange(
  format: ParagraphFormatting,
  operand: Uint8Array,
  withClose: boolean,
): void {
  let offset = 0
  const deleteCount = operand[offset++]
  if (deleteCount === undefined) return
  const deleted: number[] = []
  for (let index = 0; index < deleteCount; index++) {
    const value = readSprmI16(operand, offset)
    if (value === undefined) return
    deleted.push(value)
    offset += 2
  }
  if (withClose) offset += deleteCount * 2
  const addCount = operand[offset++]
  if (addCount === undefined) return
  const added: number[] = []
  for (let index = 0; index < addCount; index++) {
    const value = readSprmI16(operand, offset)
    if (value === undefined) return
    added.push(value)
    offset += 2
  }
  if (offset + addCount > operand.length) return
  const stops = (format.tabStops ?? []).filter((stop) => !deleted.includes(stop.positionTwips))
  added.forEach((positionTwips, index) => {
    const tbd = operand[offset + index]!
    const alignment = TAB_ALIGNMENTS[tbd & 0x07] ?? 'left'
    const leader = TAB_LEADERS[(tbd >> 3) & 0x07]
    const stop: TabStop = { positionTwips, alignment, ...(leader ? { leader } : {}) }
    const existing = stops.findIndex((candidate) => candidate.positionTwips === positionTwips)
    if (existing >= 0) stops[existing] = stop
    else stops.push(stop)
  })
  stops.sort((left, right) => left.positionTwips - right.positionTwips)
  if (stops.length === 0) delete format.tabStops
  else format.tabStops = stops
}

function setCharIndent(
  format: ParagraphFormatting,
  key: keyof CharacterIndents,
  operand: Uint8Array,
): void {
  const value = readSprmI16(operand)
  if (value === undefined) return
  format.charIndents = { ...format.charIndents, [key]: value }
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
      if (alignment !== undefined) {
        format.alignment = alignment
        format.alignmentLogical = sprm.opcode === SPRM_P_JC
      }
      break
    }
    case SPRM_P_F_BIDI:
      applyToggle(format, 'bidi', sprm.operand, mode)
      break
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
    case SPRM_P_CHG_TABS_PAPX:
      applyTabChange(format, sprm.operand, false)
      break
    case SPRM_P_CHG_TABS:
      applyTabChange(format, sprm.operand, true)
      break
    case SPRM_P_DXC_LEFT:
      setCharIndent(format, 'left', sprm.operand)
      break
    case SPRM_P_DXC_LEFT1:
      setCharIndent(format, 'firstLine', sprm.operand)
      break
    case SPRM_P_DXA_ABS:
    case SPRM_P_DYA_ABS:
    case SPRM_P_DXA_WIDTH:
    case SPRM_P_DXA_FROM_TEXT:
    case SPRM_P_DYA_FROM_TEXT: {
      const value = readSprmI16(sprm.operand)
      if (value === undefined) break
      const frame = (format.frame ??= {})
      if (sprm.opcode === SPRM_P_DXA_ABS) frame.xTwips = value
      else if (sprm.opcode === SPRM_P_DYA_ABS) frame.yTwips = value
      else if (sprm.opcode === SPRM_P_DXA_WIDTH) frame.widthTwips = value
      else if (sprm.opcode === SPRM_P_DXA_FROM_TEXT) frame.hSpaceTwips = value
      else frame.vSpaceTwips = value
      break
    }
    case SPRM_P_PC: {
      const value = readSprmU8(sprm.operand)
      if (value === undefined) break
      const frame = (format.frame ??= {})
      frame.vAnchor = FRAME_V_ANCHORS[(value >> 4) & 3] ?? 'margin'
      frame.hAnchor = FRAME_H_ANCHORS[(value >> 6) & 3] ?? 'text'
      break
    }
    case SPRM_P_WR: {
      const value = readSprmU8(sprm.operand)
      if (value === undefined) break
      ;(format.frame ??= {}).wrap = FRAME_WRAPS[value] ?? 'auto'
      break
    }
    case SPRM_P_DYA_HEIGHT: {
      const value = readSprmU16(sprm.operand)
      if (value === undefined) break
      const frame = (format.frame ??= {})
      frame.heightTwips = value & 0x7fff
      frame.heightExact = (value & 0x8000) !== 0
      break
    }
    case SPRM_P_DXC_RIGHT:
      setCharIndent(format, 'right', sprm.operand)
      break
    case SPRM_P_SHD80:
      if (sprm.operand.length >= 2) format.shading = shading80At(sprm.operand, 0)
      break
    case SPRM_P_SHD:
      // SHDOperand's cb is the spra-6 length byte iterateSprms already consumed
      if (sprm.operand.length >= 10) format.shading = shadingAt(sprm.operand, 0)
      break
    default: {
      const side80 = BORDER_SIDES_80[sprm.opcode]
      if (side80) setBorder(format, side80, sprm.operand, false)
      const side = BORDER_SIDES[sprm.opcode]
      if (side) setBorder(format, side, sprm.operand, true)
      // other unknown/unsupported paragraph sprms are skipped by the operand-size rule
      break
    }
  }
}

const MAX_HUGE_PAPX_DEPTH = 4

function isHugePapx(opcode: number): boolean {
  return opcode === SPRM_P_HUGE_PAPX || opcode === SPRM_P_HUGE_PAPX_2
}

/**
 * sprmPHugePapx: the rest of the paragraph's properties live in the Data
 * stream at the 4-byte FC operand: a signed 16-bit cb followed by the grpprl.
 * A missing Data stream or an invalid block records a warning and yields nothing.
 */
function hugePapxGrpprl(sprm: Sprm, context: ParagraphContext): Uint8Array | undefined {
  const fc = readSprmU32(sprm.operand)
  const data = context.source.stream?.('Data')
  if (fc === undefined || !data) {
    context.warnings.push(
      'a paragraph stores extra properties in the Data stream, which this file does not have; they were ignored',
    )
    return undefined
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
  if (fc + 2 > data.length) {
    context.warnings.push(
      'sprmPHugePapx points outside the Data stream; the extra paragraph properties were ignored',
    )
    return undefined
  }
  const cb = view.getInt16(fc, true)
  if (cb <= 0 || fc + 2 + cb > data.length) {
    context.warnings.push(
      'sprmPHugePapx has an invalid property block length; the extra paragraph properties were ignored',
    )
    return undefined
  }
  return data.slice(fc + 2, fc + 2 + cb)
}

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
  const grpprl = hugePapxGrpprl(sprm, context)
  if (grpprl) applyParagraphGrpprl(format, grpprl, context, mode, depth + 1)
}

/**
 * The grpprl with every sprmPHugePapx replaced by the Data-stream block it
 * points at, so readers that look for specific sprms (the table reader's TAP
 * sprms live there) see the paragraph's complete property set.
 */
export function expandHugePapx(
  grpprl: Uint8Array,
  context: ParagraphContext,
  depth = 0,
): Uint8Array {
  const sprms = readParagraphSprms(grpprl, context)
  if (!sprms.some((sprm) => isHugePapx(sprm.opcode))) return grpprl
  const parts: Uint8Array[] = []
  for (const sprm of sprms) {
    if (!isHugePapx(sprm.opcode)) {
      parts.push(grpprl.subarray(sprm.start, sprm.end))
      continue
    }
    if (depth >= MAX_HUGE_PAPX_DEPTH) {
      context.warnings.push('nested sprmPHugePapx records were ignored (too deep)')
      continue
    }
    const inner = hugePapxGrpprl(sprm, context)
    if (inner) parts.push(expandHugePapx(inner, context, depth + 1))
  }
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

/** the sprms of a PAPX grpprl; a truncated tail is reported once and dropped */
function readParagraphSprms(grpprl: Uint8Array, context: ParagraphContext): Sprm[] {
  const read = iterateSprmsLenient(grpprl)
  if (read.truncated !== undefined)
    context.warnings.push(
      'a paragraph property record is truncated (' +
        read.truncated +
        '); the properties after that point were ignored',
    )
  return read.sprms
}

/** Apply a raw PAPX grpprl to an existing paragraph format (used by the style resolver). */
export function applyParagraphGrpprl(
  format: ParagraphFormatting,
  grpprl: Uint8Array,
  context: ParagraphContext,
  mode: ToggleMode = 'direct',
  depth = 0,
): void {
  for (const sprm of readParagraphSprms(grpprl, context))
    applyParagraphSprm(format, sprm, context, depth, mode)
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
        // a cell/row-end paragraph (after a \x07 mark) is a table unit, read by
        // stage 3c-2; anything else is a genuinely misplaced record
        const afterMark = cp > 0 && source.rawText[cp - 1] === '\u0007'
        if (!afterMark)
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
  const paragraphs = spans.map((span, index) => {
    const mark = characterRuns.find((run) => run.cpStart <= span.endCp && run.cpEnd > span.endCp)
    return {
      startCp: span.startCp,
      endCp: span.endCp,
      text: span.text,
      formatting: collected.formats.get(index) ?? {},
      runs: buildRuns(span, characterRuns, source.rawText),
      ...(mark ? { markFormatting: mark.formatting } : {}),
    }
  })
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
