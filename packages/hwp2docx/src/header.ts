/**
 * Contents/header.xml: the shared reference tables every body element points
 * into by id — fonts per script slot, character shapes (hh:charPr), paragraph
 * shapes (hh:paraPr) and border/fill sets (hh:borderFill). Read once into flat
 * lookups; the body walker never touches header XML itself.
 */

import type { ParaAlign, ParaFormat } from '../../docx-engine/src/index'
import { hwpToTwips } from './units'
import {
  attr,
  descendants,
  elements,
  firstChild,
  firstDescendant,
  intAttr,
  unwrapSwitch,
  type XNode,
} from './xml'

export interface CharShape {
  sizeHalfPoints?: number
  /** hex without '#' */
  color?: string
  shading?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  vertAlign?: 'superscript' | 'subscript'
  /** HANGUL slot face (the East Asian font) */
  fontHangul?: string
  /** LATIN slot face */
  fontLatin?: string
  /** horizontal scale percent of the HANGUL slot (hh:ratio) */
  scalePct?: number
  /** letter spacing percent of the HANGUL slot (hh:spacing), relative to size */
  spacingPct?: number
}

export interface BorderSide {
  /** OOXML w:val */
  val: string
  /** eighths of a point */
  sz: number
  color: string
}

export interface BorderFill {
  left: BorderSide | null
  right: BorderSide | null
  top: BorderSide | null
  bottom: BorderSide | null
  /** cell / paragraph background, hex without '#' */
  fill?: string
}

export interface ParaShape {
  format: ParaFormat
  /** hh:heading type: OUTLINE / NUMBER / BULLET paragraphs carry auto numbering */
  /** East Asian line breaking and auto spacing, written into w:pPr as-is (Word defaults are all on) */
  eastAsian?: { autoSpaceDE: boolean; autoSpaceDN: boolean; wordWrap: boolean }
  /** hh:lineSpacing PERCENT value: line height = font size × percent (Hangul's rule) */
  linePercent?: number
  numbering?: 'outline' | 'number' | 'bullet'
  level?: number
  /** hh:heading idRef: the bullet or numbering the head comes from */
  headingIdRef?: string
}

export interface HeaderTables {
  charShape(id: string | undefined): CharShape
  paraShape(id: string | undefined): ParaShape
  borderFill(id: string | undefined): BorderFill | null
}

const ALIGN: Record<string, ParaAlign> = {
  LEFT: 'left',
  RIGHT: 'right',
  CENTER: 'center',
  JUSTIFY: 'justify',
  DISTRIBUTE: 'distribute',
  DISTRIBUTE_SPACE: 'distribute',
}

/** HWPX border line types → OOXML w:val (unknown shapes fall back to single) */
const BORDER_VAL: Record<string, string> = {
  SOLID: 'single',
  DASH: 'dashed',
  DOT: 'dotted',
  DASH_DOT: 'dotDash',
  DASH_DOT_DOT: 'dotDotDash',
  LONG_DASH: 'dashed',
  CIRCLE: 'dotted',
  DOUBLE_SLIM: 'double',
  SLIM_THICK: 'thinThickSmallGap',
  THICK_SLIM: 'thickThinSmallGap',
  SLIM_THICK_SLIM: 'thinThickThinSmallGap',
  WAVE: 'wave',
  DOUBLE_WAVE: 'doubleWave',
  THICK_3D: 'threeDEmboss',
  THICK_3D_REVERS: 'threeDEngrave',
  '3D': 'threeDEmboss',
  '3D_REVERS': 'threeDEngrave',
}

function hexColor(value: string | undefined): string | undefined {
  if (!value || value === 'none') return undefined
  const match = /^#?([0-9a-f]{6})([0-9a-f]{2})?$/i.exec(value)
  if (!match) return undefined
  // #AARRGGBB (alpha first) appears in a few producers: keep the RGB tail
  if (match[2]) return (match[1].slice(2) + match[2]).toUpperCase()
  return match[1].toUpperCase()
}

function widthToEighths(width: string | undefined): number {
  const mm = Number.parseFloat(width ?? '0.1')
  if (!Number.isFinite(mm) || mm <= 0) return 4
  // mm → pt → eighths of a point; OOXML caps w:sz at 96 (12pt), floor 2
  return Math.min(96, Math.max(2, Math.round(((mm * 72) / 25.4) * 8)))
}

function borderSide(node: XNode | null): BorderSide | null {
  if (!node) return null
  const type = attr(node, 'type') ?? 'NONE'
  if (type === 'NONE') return null
  const val = BORDER_VAL[type] ?? 'single'
  const total = widthToEighths(attr(node, 'width'))
  return {
    val,
    // Hangul's width is the whole compound line; Word's w:sz is one stroke of
    // it and the gap repeats it (a 0.5 mm double drew 4pt apart, Hangul 2pt)
    sz: COMPOUND.has(val) ? Math.max(2, Math.round(total / 3)) : total,
    color: hexColor(attr(node, 'color')) ?? '000000',
  }
}

/** compound lines: Word draws two strokes of w:sz with a gap of the same width */
export const COMPOUND = new Set([
  'double',
  'thinThickSmallGap',
  'thickThinSmallGap',
  'thinThickThinSmallGap',
])

function slotValue(node: XNode | null, slot: 'hangul' | 'latin'): string | undefined {
  return attr(node, slot)
}

export function readHeader(root: XNode[]): HeaderTables {
  // -- fonts: lang → id → face
  const fonts = new Map<string, Map<string, string>>()
  for (const face of descendants(root, 'hh:fontface')) {
    const lang = attr(face, 'lang') ?? ''
    const byId = new Map<string, string>()
    for (const font of elements(face)) {
      if (font.name !== 'hh:font') continue
      const id = attr(font.node, 'id')
      const name = attr(font.node, 'face')
      if (id !== undefined && name) byId.set(id, name)
    }
    fonts.set(lang, byId)
  }
  const fontOf = (lang: string, id: string | undefined): string | undefined =>
    id === undefined ? undefined : fonts.get(lang)?.get(id)

  // -- character shapes
  const chars = new Map<string, CharShape>()
  for (const node of descendants(root, 'hh:charPr')) {
    const id = attr(node, 'id')
    if (id === undefined) continue
    const shape: CharShape = {}
    const height = intAttr(node, 'height', 0)
    if (height > 0) shape.sizeHalfPoints = Math.max(2, Math.round(height / 50))
    const color = hexColor(attr(node, 'textColor'))
    if (color && color !== '000000') shape.color = color
    const shade = hexColor(attr(node, 'shadeColor'))
    if (shade && shade !== 'FFFFFF') shape.shading = shade
    const fontRef = firstChild(node, 'hh:fontRef')
    shape.fontHangul = fontOf('HANGUL', slotValue(fontRef, 'hangul'))
    shape.fontLatin = fontOf('LATIN', slotValue(fontRef, 'latin'))
    const ratio = Number(slotValue(firstChild(node, 'hh:ratio'), 'hangul'))
    if (Number.isFinite(ratio) && ratio > 0 && ratio !== 100) shape.scalePct = ratio
    const spacing = Number(slotValue(firstChild(node, 'hh:spacing'), 'hangul'))
    if (Number.isFinite(spacing) && spacing !== 0) shape.spacingPct = spacing
    for (const child of elements(node)) {
      switch (child.name) {
        case 'hh:bold':
          shape.bold = true
          break
        case 'hh:italic':
          shape.italic = true
          break
        case 'hh:underline': {
          const type = attr(child.node, 'type') ?? 'NONE'
          if (type !== 'NONE') shape.underline = true
          break
        }
        case 'hh:strikeout': {
          // Hangul writes shape="3D" (or NONE) on every charPr as the "off" value
          const shapeName = attr(child.node, 'shape') ?? 'NONE'
          if (shapeName !== 'NONE' && shapeName !== '3D') shape.strike = true
          break
        }
        case 'hh:supscript':
          shape.vertAlign = 'superscript'
          break
        case 'hh:subscript':
          shape.vertAlign = 'subscript'
          break
      }
    }
    chars.set(id, shape)
  }

  // -- paragraph shapes
  const paras = new Map<string, ParaShape>()
  for (const node of descendants(root, 'hh:paraPr')) {
    const id = attr(node, 'id')
    if (id === undefined) continue
    const format: ParaFormat = {}
    const content = unwrapSwitch(node)
    const find = (name: string): XNode | null =>
      content.find((child) => Object.prototype.hasOwnProperty.call(child, name)) ?? null
    const align = ALIGN[attr(find('hh:align'), 'horizontal') ?? '']
    if (align) format.align = align
    const margin = find('hh:margin')
    if (margin) {
      const value = (name: string): number | undefined => {
        const child = firstChild(margin, name)
        if (!child) return undefined
        const unit = attr(child, 'unit') ?? 'HWPUNIT'
        if (unit !== 'HWPUNIT') return undefined
        return hwpToTwips(intAttr(child, 'value', 0))
      }
      const left = value('hc:left')
      const right = value('hc:right')
      const first = value('hc:intent')
      const before = value('hc:prev')
      const after = value('hc:next')
      if (left) format.indentLeft = left
      if (right) format.indentRight = right
      if (first) format.indentFirstLine = first
      // Hangul's negative (hanging-style) indent starts line 1 at the left margin and
      // pushes the following lines in; Word's hanging indent hangs line 1 out
      // of the left indent instead, so the left indent grows by the same amount
      if (first && first < 0) format.indentLeft = (left ?? 0) - first
      format.spaceBefore = Math.max(0, before ?? 0)
      format.spaceAfter = Math.max(0, after ?? 0)
    }
    const line = find('hh:lineSpacing')
    if (line) {
      // PERCENT is the font size times the percent: paragraphBlocks sizes those
      // lines from shape.linePercent once it knows the paragraph's glyphs
      const type = attr(line, 'type') ?? 'PERCENT'
      const value = intAttr(line, 'value', 160)
      if (type === 'FIXED' && value > 0) {
        format.lineRule = 'exact'
        format.lineRawTwips = hwpToTwips(value)
      } else if (type === 'AT_LEAST' && value > 0) {
        format.lineRule = 'atLeast'
        format.lineRawTwips = hwpToTwips(value)
      }
    }
    const brk = find('hh:breakSetting')
    if (attr(brk, 'pageBreakBefore') === '1') format.pageBreakBefore = true
    if (attr(brk, 'keepWithNext') === '1') format.keepNext = true
    if (attr(brk, 'keepLines') === '1') format.keepLines = true
    const shape: ParaShape = { format }
    // Hangul defaults: no auto spacing around Latin/digits; Word defaults are
    // the opposite, so always spell them out. breakNonLatinWord reads backwards
    // from its name: Hangul's saved line layouts break KEEP_WORD paragraphs
    // between syllables (573 of 1191 corpus breaks fell mid-word) and
    // BREAK_WORD ones at spaces (1 of 107), so only BREAK_WORD keeps Word's
    // whole-word wrapping and KEEP_WORD (the default) turns w:wordWrap off.
    // That holds for every writer: the PDFs Hangul 2018 and 2022 export keep
    // most words whole but still split them at particles and endings
    // (강화|하는, 마련하겠|습니다), which whole-word wrapping misses by more
    // (corpus: 2,530 vs 2,605 of 2,710 line ends matched)
    const auto = find('hh:autoSpacing')
    shape.eastAsian = {
      autoSpaceDE: attr(auto, 'eAsianEng') === '1',
      autoSpaceDN: attr(auto, 'eAsianNum') === '1',
      wordWrap: attr(brk, 'breakNonLatinWord') === 'BREAK_WORD',
    }
    if (line && (attr(line, 'type') ?? 'PERCENT') === 'PERCENT') {
      const percent = intAttr(line, 'value', 160)
      if (percent > 0) shape.linePercent = percent
    }
    const heading = find('hh:heading')
    const headingType = attr(heading, 'type')
    if (headingType === 'OUTLINE') shape.numbering = 'outline'
    else if (headingType === 'NUMBER') shape.numbering = 'number'
    else if (headingType === 'BULLET') shape.numbering = 'bullet'
    if (shape.numbering) {
      shape.level = intAttr(heading, 'level', 0)
      shape.headingIdRef = attr(heading, 'idRef')
    }
    paras.set(id, shape)
  }

  // -- border/fill sets
  const fills = new Map<string, BorderFill>()
  for (const node of descendants(root, 'hh:borderFill')) {
    const id = attr(node, 'id')
    if (id === undefined) continue
    const brush = firstDescendant(node, 'hc:winBrush')
    fills.set(id, {
      left: borderSide(firstChild(node, 'hh:leftBorder')),
      right: borderSide(firstChild(node, 'hh:rightBorder')),
      top: borderSide(firstChild(node, 'hh:topBorder')),
      bottom: borderSide(firstChild(node, 'hh:bottomBorder')),
      fill: hexColor(attr(brush, 'faceColor')),
    })
  }

  return {
    charShape: (id) => (id !== undefined ? chars.get(id) : undefined) ?? {},
    paraShape: (id) => (id !== undefined ? paras.get(id) : undefined) ?? { format: {} },
    borderFill: (id) => (id !== undefined ? fills.get(id) : undefined) ?? null,
  }
}
