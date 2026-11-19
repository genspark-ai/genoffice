/**
 * Tiptap entry: the same HWP/HWPX → .docx filter GenOffice opens with, read
 * back through docx-engine and emitted as standard Tiptap JSON (StarterKit,
 * TextStyle/Color/FontFamily/FontSize, Highlight, TextAlign, Subscript/
 * Superscript, Link, Image, Table). There is no second HWP reader: anything
 * the docx filter learns, a Tiptap editor gets with it, and both stay
 * comparable on the same document.
 *
 * Letter spacing rides on textStyle as `letterSpacing` (a host declares it).
 * Images are emitted inside paragraphs, so the Image extension needs
 * `inline: true`. Attributes a stock extension does not declare are dropped
 * by the editor schema, never misread.
 */

import {
  parseDocx,
  readSections,
  type Block,
  type CellBorders,
  type ParaAlign,
  type ParaFormat,
  type Run,
  type TableCell,
  type TableModel,
  type TableParagraph,
} from '../../docx-engine/src/index'
import { convertHwpToDocx } from './hwp'
import type { HwpxToDocxResult } from './hwpx-docx'

export interface TiptapMark {
  type: string
  attrs?: Record<string, unknown>
}

export interface TiptapNode {
  type: string
  attrs?: Record<string, unknown>
  content?: TiptapNode[]
  marks?: TiptapMark[]
  text?: string
}

/** the first section's sheet, in points: a host lays the editor out on it */
export interface TiptapPage {
  widthPt: number
  heightPt: number
  marginTopPt: number
  marginRightPt: number
  marginBottomPt: number
  marginLeftPt: number
}

export interface TiptapDocument {
  /** Tiptap document JSON (`editor.commands.setContent(doc)`) */
  doc: TiptapNode
  /** docx blocks with no Tiptap counterpart (section/body passthrough parts) */
  skipped: number
  page?: TiptapPage
}

export type HwpToTiptapResult = HwpxToDocxResult & TiptapDocument

const TWIPS_PER_PX = 15

const ALIGN: Partial<Record<ParaAlign, string>> = {
  left: 'left',
  center: 'center',
  right: 'right',
  justify: 'justify',
  // Hangul's distributed alignment has no CSS counterpart beyond justify
  distribute: 'justify',
}

/** Any Hancom word-processor file → Tiptap JSON, through the docx filter */
export async function convertHwpToTiptap(bytes: Uint8Array): Promise<HwpToTiptapResult> {
  const converted = await convertHwpToDocx(bytes)
  return { ...converted, ...(await docxToTiptap(converted.docx, { hwpOrigin: true })) }
}

/**
 * A .docx (as written by the filter, or any other) → Tiptap JSON. `hwpOrigin`:
 * the filter wrote it, so letter spacing carries Word's doubled double-byte
 * share (hwpx-docx runOf), which CSS undoes per glyph.
 */
export async function docxToTiptap(
  docx: Uint8Array,
  options: { hwpOrigin?: boolean } = {},
): Promise<TiptapDocument> {
  const parsed = await parseDocx(docx)
  hwpOrigin = options.hwpOrigin === true
  const content: TiptapNode[] = []
  let skipped = 0
  for (const block of parsed.blocks) {
    const node = blockNode(block)
    if (node) content.push(node)
    else skipped++
  }
  // a Tiptap document needs at least one block
  if (content.length === 0) content.push({ type: 'paragraph' })
  const sheet = readSections(parsed)[0]?.settings
  const pt = (twips: number) => Math.round(twips / 2) / 10
  const page = sheet
    ? {
        widthPt: pt(sheet.pageWidth),
        heightPt: pt(sheet.pageHeight),
        marginTopPt: pt(sheet.marginTop),
        marginRightPt: pt(sheet.marginRight),
        marginBottomPt: pt(sheet.marginBottom),
        marginLeftPt: pt(sheet.marginLeft),
      }
    : undefined
  return { doc: { type: 'doc', content }, skipped, ...(page ? { page } : {}) }
}

let hwpOrigin = false

/** CSS font-family for a run: the face as the document names it */
function fontFamilyCss(face: string): string {
  return face
}

/** paragraph geometry as inline CSS for a `style` attribute (stock Paragraph drops it) */
function paragraphCss(format: ParaFormat): string {
  const pt = (twips: number) => `${Math.round(twips / 2) / 10}pt`
  const css: string[] = []
  if (format.lineRawTwips && (format.lineRule === 'exact' || format.lineRule === 'atLeast')) {
    css.push(`line-height:${pt(format.lineRawTwips)}`)
  }
  if (format.spaceBefore) css.push(`margin-top:${pt(format.spaceBefore)}`)
  if (format.spaceAfter) css.push(`margin-bottom:${pt(format.spaceAfter)}`)
  if (format.indentLeft) css.push(`padding-left:${pt(format.indentLeft)}`)
  if (format.indentRight) css.push(`padding-right:${pt(format.indentRight)}`)
  if (format.indentFirstLine) css.push(`text-indent:${pt(format.indentFirstLine)}`)
  return css.join(';')
}

function blockNode(block: Block): TiptapNode | null {
  switch (block.type) {
    case 'table':
      return block.table ? tableNode(block.table) : null
    case 'image':
      return block.imageDataUrl
        ? paragraphNode({ align: block.imageAlign }, [
            imageNode(block.imageDataUrl, block.imageWidthPx, block.imageHeightPx),
          ])
        : null
    case 'heading': {
      const node = paragraphNode(block.format ?? {}, inlineNodes(block.runs ?? []))
      return {
        ...node,
        type: 'heading',
        attrs: { ...node.attrs, level: Math.min(6, Math.max(1, block.level ?? 1)) },
      }
    }
    case 'paragraph':
    case 'listItem':
      return paragraphNode(block.format ?? {}, inlineNodes(block.runs ?? []))
    default:
      return null
  }
}

function paragraphNode(format: ParaFormat, content: TiptapNode[]): TiptapNode {
  const align = format.align ? ALIGN[format.align] : undefined
  const style = paragraphCss(format)
  const attrs = {
    ...(align && align !== 'left' ? { textAlign: align } : {}),
    ...(style ? { style } : {}),
  }
  return {
    type: 'paragraph',
    ...(Object.keys(attrs).length > 0 ? { attrs } : {}),
    ...(content.length > 0 ? { content } : {}),
  }
}

function imageNode(src: string, width?: number, height?: number): TiptapNode {
  return {
    type: 'image',
    attrs: {
      src,
      ...(width ? { width: Math.round(width) } : {}),
      ...(height ? { height: Math.round(height) } : {}),
    },
  }
}

/** runs → text / hardBreak / image nodes; a line break inside a run is a hard break */
function inlineNodes(runs: Run[]): TiptapNode[] {
  const out: TiptapNode[] = []
  for (const run of runs) {
    if (run.vanish) continue
    if (run.image?.dataUrl) {
      out.push(imageNode(run.image.dataUrl, run.image.widthPx, run.image.heightPx))
    }
    if (!run.text) continue
    run.text.split('\n').forEach((line, index) => {
      if (index > 0) out.push({ type: 'hardBreak' })
      for (const piece of spacingPieces(line, run)) {
        const marks = runMarks(run, piece.half)
        // runs the docx splits for Word-only reasons fold back together
        const last = out[out.length - 1]
        if (last?.type === 'text' && JSON.stringify(last.marks ?? []) === JSON.stringify(marks)) {
          last.text += piece.text
        } else {
          out.push({ type: 'text', text: piece.text, ...(marks.length > 0 ? { marks } : {}) })
        }
      }
    })
  }
  return out
}

const FULL_WIDTH = /[\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uff60]/

/**
 * Hangul's letter spacing is a percent of each glyph's own advance; CSS adds
 * one length per glyph. Full-width glyphs take the whole value, half-width
 * ones (Latin, digits, spaces, hyphens) half of it.
 */
function spacingPieces(text: string, run: Run): Array<{ text: string; half: boolean }> {
  if (!text) return []
  if (!hwpOrigin || !run.charSpacingTwips) return [{ text, half: false }]
  const out: Array<{ text: string; half: boolean }> = []
  for (const ch of text) {
    const half = !FULL_WIDTH.test(ch)
    const last = out[out.length - 1]
    if (last && last.half === half) last.text += ch
    else out.push({ text: ch, half })
  }
  return out
}

function runMarks(run: Run, halfWidth = false): TiptapMark[] {
  const marks: TiptapMark[] = []
  if (run.bold) marks.push({ type: 'bold' })
  if (run.italic) marks.push({ type: 'italic' })
  if (run.underline) marks.push({ type: 'underline' })
  if (run.strike || run.dstrike) marks.push({ type: 'strike' })
  if (run.vertAlign === 'superscript') marks.push({ type: 'superscript' })
  if (run.vertAlign === 'subscript') marks.push({ type: 'subscript' })
  if (run.link?.href) marks.push({ type: 'link', attrs: { href: run.link.href } })
  const style: Record<string, string> = {}
  if (run.color && run.color !== 'auto') style.color = `#${run.color}`
  const font = run.eastAsiaFont ?? run.font
  if (font) style.fontFamily = fontFamilyCss(font)
  if (run.sizeHalfPoints) style.fontSize = `${run.sizeHalfPoints / 2}pt`
  if (run.charSpacingTwips) {
    // the docx carries half of Hangul's percent for Word's doubled double-byte
    // spacing (hwpx-docx runOf): whole again on full-width glyphs, as it is on
    // half-width ones (spacingPieces)
    const twips = hwpOrigin && !halfWidth ? run.charSpacingTwips * 2 : run.charSpacingTwips
    style.letterSpacing = `${Math.round(twips * 5) / 100}pt`
  }
  if (Object.keys(style).length > 0) marks.push({ type: 'textStyle', attrs: style })
  const highlight = run.shading ?? run.highlight
  if (highlight && /^[0-9a-f]{6}$/i.test(highlight)) {
    marks.push({ type: 'highlight', attrs: { color: `#${highlight}` } })
  }
  return marks
}

// -- tables -------------------------------------------------------------------

/**
 * docx-engine keeps Word's shape: horizontal spans as gridSpan, vertical ones
 * as a restart cell over continue cells. Tiptap wants rowspan on the origin
 * and no cell in the slots it covers.
 */
function tableNode(table: TableModel): TiptapNode {
  const colPx = (table.colWidthsTwips ?? []).map((tw) => Math.max(1, Math.round(tw / TWIPS_PER_PX)))
  // grid column of every cell, then how far each restart reaches down
  const placed = table.rows.map((row) => {
    let col = 0
    return row.map((cell) => {
      const at = col
      col += Math.max(1, cell.colSpan ?? 1)
      return { cell, col: at }
    })
  })
  const rows: TiptapNode[] = placed.map((row, r) => {
    const cells: TiptapNode[] = []
    for (const { cell, col } of row) {
      if (cell.vMerge === 'continue') continue
      let rowspan = 1
      if (cell.vMerge === 'restart') {
        while (
          r + rowspan < placed.length &&
          placed[r + rowspan].some((next) => next.col === col && next.cell.vMerge === 'continue')
        ) {
          rowspan++
        }
      }
      cells.push(cellNode(cell, col, rowspan, colPx))
    }
    return { type: 'tableRow', content: cells.length > 0 ? cells : [emptyCell()] }
  })
  // a table set off the left edge (w:jc right / center) keeps its place
  const style =
    table.align === 'right'
      ? 'margin-left:auto;margin-right:0'
      : table.align === 'center'
        ? 'margin-left:auto;margin-right:auto'
        : ''
  return { type: 'table', ...(style ? { attrs: { style } } : {}), content: rows }
}

function cellNode(cell: TableCell, col: number, rowspan: number, colPx: number[]): TiptapNode {
  const colspan = Math.max(1, cell.colSpan ?? 1)
  const widths = colPx.slice(col, col + colspan)
  const content = cellContent(cell)
  return {
    type: 'tableCell',
    attrs: {
      colspan,
      rowspan,
      colwidth: widths.length === colspan ? widths : null,
      ...(cell.fill && /^[0-9a-f]{6}$/i.test(cell.fill)
        ? { backgroundColor: `#${cell.fill}` }
        : {}),
      ...(cellCss(cell) ? { style: cellCss(cell) } : {}),
    },
    content: content.length > 0 ? content : [{ type: 'paragraph' }],
  }
}

/** cell borders and vertical alignment as inline CSS */
function cellCss(cell: TableCell): string {
  const css = cell.borders ? [borderCss(cell.borders)] : []
  if (cell.vAlign === 'center') css.push('vertical-align:middle')
  else if (cell.vAlign === 'bottom') css.push('vertical-align:bottom')
  return css.join(';')
}

const CSS_BORDER: Record<string, string> = {
  dashed: 'dashed',
  dotted: 'dotted',
  double: 'double',
  thinThickSmallGap: 'double',
  thickThinSmallGap: 'double',
}

/**
 * Cell borders as inline CSS for a `style` attribute: stock TableCell has no
 * border attribute, so a host that wants form lines declares `style` as a
 * global attribute; without it the editor drops the value.
 */
function borderCss(borders: CellBorders): string {
  const sides = ['top', 'right', 'bottom', 'left'] as const
  return sides
    .map((side) => {
      const line = borders[side]
      if (!line || line.style === 'nil' || line.style === 'none') return `border-${side}:none`
      const px = Math.max(1, Math.round(((line.szEighths ?? 4) / 8) * (96 / 72) * 10) / 10)
      const color = line.color && line.color !== 'auto' ? `#${line.color}` : '#000'
      return `border-${side}:${px}px ${CSS_BORDER[line.style] ?? 'solid'} ${color}`
    })
    .join(';')
}

/** cell paragraphs with nested tables placed before the paragraph they anchor to */
function cellContent(cell: TableCell): TiptapNode[] {
  const paras: TableParagraph[] =
    cell.richParas ?? cell.paras.map((text) => ({ runs: text ? [{ text }] : [] }))
  const nested = cell.nestedTables ?? []
  const anchors = cell.nestedTableAnchors ?? nested.map(() => paras.length)
  const out: TiptapNode[] = []
  for (let i = 0; i <= paras.length; i++) {
    nested.forEach((table, t) => {
      if (Math.min(anchors[t] ?? paras.length, paras.length) === i) out.push(tableNode(table))
    })
    if (i < paras.length) out.push(paragraphNode(paras[i], inlineNodes(paras[i].runs)))
  }
  return out
}

function emptyCell(): TiptapNode {
  return { type: 'tableCell', attrs: { colspan: 1, rowspan: 1 }, content: [{ type: 'paragraph' }] }
}
