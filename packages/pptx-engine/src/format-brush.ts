/**
 * Format Painter — pure functions with no side effects, easy to unit test.
 *
 * Format scope:
 *  - Shapes/text boxes: the modelable part of spPr (fill / stroke / presetGeometry / adjust)
 *  - Text: default run format (fontFamily/fontSize/bold/italic/underline/color of the
 *          first run of the first paragraph) + paragraph alignment (first paragraph's align)
 *  - Cross-type: text→shape applies only the spPr part; shape→text applies only the text part
 *
 * Uses the existing editFill / editStroke / editText IPC channels; the caller takes the
 * CopiedFormat to the target element and calls each patch function field by field.
 */

import type { Fill, Stroke, TextElement } from './types'

// ── Format containers ────────────────────────────────────────────────────

/** Copied shape appearance format (spPr layer) */
export interface ShapeFormat {
  fill?: Fill
  stroke?: Stroke
  /** Corner-radius adjustments for preset geometry such as roundRect */
  adjust?: Record<string, number>
}

/** Copied default text run format (first run of the first paragraph) */
export interface TextRunFormat {
  fontFamily?: string
  fontSize?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
}

/** Copied paragraph format (first paragraph) */
export interface ParagraphFormat {
  align?: 'left' | 'center' | 'right' | 'justify'
}

/** Full container of formats copied by the format painter */
export interface CopiedFormat {
  /** Source element type */
  sourceType: 'text' | 'shape' | 'picture'
  shape?: ShapeFormat
  run?: TextRunFormat
  paragraph?: ParagraphFormat
}

// ── Extraction (source element → CopiedFormat) ─────────────────────────

/**
 * Extract format from a TextElement (type='text'|'shape').
 * Only extracts the "write-back-able" parts; ignores passthrough/group/picture.
 */
export function extractFormat(el: TextElement): CopiedFormat {
  const shape: ShapeFormat = {}
  if (el.fill) shape.fill = el.fill
  if (el.stroke) shape.stroke = el.stroke
  if (el.adjust && Object.keys(el.adjust).length > 0) shape.adjust = el.adjust

  const run: TextRunFormat = {}
  const para: ParagraphFormat = {}

  const firstPara = el.text?.paragraphs?.[0]
  if (firstPara) {
    if (firstPara.align) para.align = firstPara.align
    const firstRun = firstPara.runs?.[0]
    if (firstRun) {
      if (firstRun.fontFamily != null) run.fontFamily = firstRun.fontFamily
      if (firstRun.fontSize != null) run.fontSize = firstRun.fontSize
      if (firstRun.bold != null) run.bold = firstRun.bold
      if (firstRun.italic != null) run.italic = firstRun.italic
      if (firstRun.underline != null) run.underline = firstRun.underline
      if (firstRun.color != null) run.color = firstRun.color
    }
  }

  return {
    sourceType: el.type as 'text' | 'shape',
    shape: Object.keys(shape).length > 0 ? shape : undefined,
    run: Object.keys(run).length > 0 ? run : undefined,
    paragraph: Object.keys(para).length > 0 ? para : undefined,
  }
}

// ── Application (CopiedFormat → target element delta) ──────────────────

export interface FormatPatchResult {
  /** Fill change (undefined = no change) */
  fill?: Fill
  /** Stroke change (undefined = no change; null = remove stroke) */
  stroke?: Fill extends Fill ? Stroke | null : Stroke | null
  /** Text run change (undefined = no change) */
  runFormat?: TextRunFormat
  /** Paragraph alignment change (undefined = no change) */
  align?: 'left' | 'center' | 'right' | 'justify'
}

/**
 * Apply a CopiedFormat to the target element and compute the delta to modify.
 * Cross-type application takes the intersection:
 *   - target is picture: apply stroke only (no fill/text)
 *   - target is shape: apply spPr + text run + alignment
 *   - target is text: apply spPr + text run + alignment
 */
export function applyFormat(
  fmt: CopiedFormat,
  targetType: 'text' | 'shape' | 'picture',
): FormatPatchResult {
  const result: FormatPatchResult = {}

  if (targetType === 'picture') {
    // Picture: apply stroke only
    if (fmt.shape?.stroke !== undefined) result.stroke = fmt.shape.stroke
    return result
  }

  // shape / text: apply everything
  if (fmt.shape?.fill !== undefined) result.fill = fmt.shape.fill
  if (fmt.shape?.stroke !== undefined) result.stroke = fmt.shape.stroke
  if (fmt.run && Object.keys(fmt.run).length > 0) result.runFormat = fmt.run
  if (fmt.paragraph?.align) result.align = fmt.paragraph.align

  return result
}
