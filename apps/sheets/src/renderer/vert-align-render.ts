/**
 * Univer never paints a whole-cell superscript or subscript. A plain-value
 * cell takes the single-`fillText` paint path, whose font string is built from
 * ff/fs/it/bl/ul/st/ol/cl only — Univer's getFontFormat leaves `va` out — so a
 * cell whose font carries <vertAlign> (set from the Home ribbon, or read from
 * an Excel file) looks exactly like baseline text while the ribbon button
 * lights up and the saved file is right. The document layout path does honor
 * `va` on a text run (smaller glyph on a raised or lowered baseline), and rich
 * text and rotated cells already go down it. This hands the font cache a
 * rich-text twin of a plain `va` cell so it takes that path too. The cell
 * model is untouched: edits, copies, sorts and saves still read `v` + `s.va`.
 */
import {
  BaselineOffset,
  CellValueType,
  getDisplayValueFromCell,
  type ICellData,
  type IDocumentData,
  type IStyleData,
  type ITextStyle,
} from '@univerjs/core'
import { SpreadsheetSkeleton } from '@univerjs/engine-render'

export function hasVertAlign(style: IStyleData | null | undefined): boolean {
  return style?.va === BaselineOffset.SUBSCRIPT || style?.va === BaselineOffset.SUPERSCRIPT
}

/**
 * The run style the plain path would have painted with (the keys Univer's
 * getFontFormat copies), plus the vertical alignment it drops.
 */
export function vertAlignRunStyle(style: IStyleData): ITextStyle {
  const run: ITextStyle = {}
  if (style.ff) run.ff = style.ff
  if (style.fs) run.fs = style.fs
  if (style.it) run.it = style.it
  if (style.bl) run.bl = style.bl
  if (style.ul) run.ul = style.ul
  if (style.st) run.st = style.st
  if (style.ol) run.ol = style.ol
  if (style.cl) run.cl = style.cl
  if (style.va) run.va = style.va
  return run
}

/// What the plain painter reads off the cell to decide default alignment:
/// numbers right, booleans centered, everything else left.
function valueTypeOf(cell: ICellData): CellValueType {
  if (cell.t !== undefined && cell.t !== null) return cell.t
  if (typeof cell.v === 'number') return CellValueType.NUMBER
  if (typeof cell.v === 'boolean') return CellValueType.BOOLEAN
  return CellValueType.STRING
}

/**
 * A rich-text twin of a plain-value cell whose style carries `va`, in the
 * same shape the loader builds for file rich text (so every layout patch that
 * recognizes a 'rich-cell' document treats it alike), or undefined when the
 * cell should keep its own paint path.
 */
export function vertAlignTwin(
  cell: ICellData | undefined,
  style: IStyleData | null | undefined,
): ICellData | undefined {
  if (!cell || cell.p || cell.v === null || cell.v === undefined || !style) return undefined
  if (!hasVertAlign(style)) return undefined
  const text = getDisplayValueFromCell(cell)
  if (text.length === 0) return undefined
  // Univer document streams use \r as the paragraph break and \n as the
  // section break; a raw \n would split the cell into sections.
  const dataStream = `${text.replace(/\r\n?|\n/g, '\r')}\r\n`
  const paragraphs: Array<{ startIndex: number }> = []
  for (let i = 0; i < dataStream.length; i += 1) {
    if (dataStream[i] === '\r') paragraphs.push({ startIndex: i })
  }
  const p: IDocumentData = {
    id: 'rich-cell',
    body: {
      dataStream,
      textRuns: [{ st: 0, ed: dataStream.length - 2, ts: vertAlignRunStyle(style) }],
      paragraphs,
      sectionBreaks: [{ startIndex: dataStream.length - 1 }],
    },
    // Univer merges the cell's alignment config over this; the value type is
    // what keeps a number right-aligned and a boolean centered when the cell
    // has no explicit alignment (the plain painter reads it off cellData.t).
    documentStyle: { renderConfig: { cellValueType: valueTypeOf(cell) } },
  }
  return { ...cell, p }
}

interface SkeletonLike {
  _renderRawFormula?: boolean
  _setFontStylesCache(
    row: number,
    col: number,
    cell: ICellData | undefined,
    style: IStyleData | undefined,
  ): void
}

let installed = false

export function installVertAlignRender(): void {
  if (installed) return
  installed = true
  const proto = SpreadsheetSkeleton.prototype as unknown as SkeletonLike
  const original = proto._setFontStylesCache
  if (typeof original !== 'function') return
  proto._setFontStylesCache = function (this: SkeletonLike, row, col, cell, style) {
    // Show Formulas paints formulas (and the ' of a forced string) through
    // the plain path; only a displayed value gets the twin.
    const twin = this._renderRawFormula ? undefined : vertAlignTwin(cell, style)
    original.call(this, row, col, twin ?? cell, style)
  }
}
