import type { TableCell, TableModel, TableParagraph } from '@genoffice/docx-engine'
import { textHasComplexScript } from '../line-metrics'
import { borderTotalPt } from '../editor/border-metrics'
import type { MeasureRun, RunDefaults, TextMeasure } from './measure'
import { nestedFloorTwips } from './distribute'
import { DEFAULT_CELL_MAR, MIN_CONTENT_SLACK } from './constants'

/** a measured width on the px grid the renderer draws columns on (convert.ts rounds twips / 15):
 * Word fits the bare word + margins (probe 2026-11-09: 30 m -> 30 x 175.7 + 216 exactly) and a
 * column rounded down by half a px overflows in Chromium (573 / 1260 kernel probe cells), ceil never */
function pxGridTwips(px: number, padTwips: number): number {
  return 15 * Math.ceil(px + padTwips / 15)
}

/** positive left-aligned w:tblInd (twips): the part of the column the table cannot use */
export function tableIndentTwips(model: TableModel): number {
  return model.align !== 'center' && model.align !== 'right' && (model.indentTwips ?? 0) > 0
    ? model.indentTwips!
    : 0
}

/** side cell margins of a table's own cells (twips) */
export function sideMarginTwips(model: TableModel): number {
  const mar = model.cellMarTwips
  return (mar?.left ?? DEFAULT_CELL_MAR) + (mar?.right ?? DEFAULT_CELL_MAR)
}

/** side padding of one cell (twips): its own w:tcMar, else the table's */
export function cellSidePadTwips(cell: TableCell, model: TableModel): number {
  const mar = cell.cellMarTwips ?? model.cellMarTwips
  return (mar?.left ?? DEFAULT_CELL_MAR) + (mar?.right ?? DEFAULT_CELL_MAR)
}

/** paragraph indent inside a cell eats line width like the word itself does;
 * list fallbacks mirror the .doc-li per-level CSS defaults (0.55in + 0.3in/level) */
function paraIndentPx(para: TableParagraph): number {
  const leftTw = para.indentLeft ?? (para.list ? 792 + 432 * (para.list.ilvl || 0) : 0)
  const firstTw = Math.max(para.indentFirstLine ?? 0, 0)
  return Math.max(leftTw + firstTw, 0) / 15
}

/** widest paragraph of a cell (px), paragraph indent included. `word`: the
 * widest unbreakable unit (min-content); `line`: the unwrapped line (max-content).
 * An inline picture is an unbreakable unit of its own width; floating ones
 * take no cell width. */
export function cellMeasurePx(
  cell: TableCell,
  mode: 'word' | 'line',
  measure: TextMeasure,
  defaults?: RunDefaults,
): number {
  const measureText = (runs: MeasureRun[]) =>
    mode === 'word' ? measure.wordPx(runs) : measure.linePx(runs)
  let max = 0
  if (cell.richParas?.length) {
    for (const para of cell.richParas) {
      const runs: MeasureRun[] = []
      let imgMax = 0
      let imgSum = 0
      for (const r of para.runs) {
        if (r.image && !r.image.wrap && (r.image.widthPx ?? 0) > 0) {
          imgMax = Math.max(imgMax, r.image.widthPx!)
          imgSum += r.image.widthPx!
        }
        if (!r.text) continue
        // Latin words render in the ascii slot; `font` may be the East Asian face;
        // complex-script text (and every character of an rtl run) takes the cs chain
        const family =
          r.csFont && (r.cs || textHasComplexScript(r.text))
            ? r.csFont
            : (r.fontAscii ?? r.font ?? defaults?.font)
        const size = r.sizeHalfPoints ?? defaults?.sizeHalfPoints
        const eastAsiaLang = r.eastAsiaLang ?? para.eastAsiaLang ?? defaults?.eastAsiaLang
        runs.push({
          text: r.text,
          ...(family ? { fontFamily: family } : {}),
          ...(size ? { sizeHalfPoints: size } : {}),
          ...(r.bold ? { bold: true } : {}),
          ...(r.italic ? { italic: true } : {}),
          ...(eastAsiaLang ? { eastAsiaLang } : {}),
        })
      }
      const textPx = runs.length ? measureText(runs) : 0
      const px = mode === 'word' ? Math.max(textPx, imgMax) : textPx + imgSum
      if (px <= 0) continue
      max = Math.max(max, px + paraIndentPx(para))
    }
  } else {
    for (const text of cell.paras) {
      if (!text) continue
      max = Math.max(
        max,
        measureText([
          {
            text,
            ...(cell.bold ? { bold: true } : {}),
            ...(defaults?.font ? { fontFamily: defaults.font } : {}),
            ...(defaults?.sizeHalfPoints ? { sizeHalfPoints: defaults.sizeHalfPoints } : {}),
            ...(defaults?.eastAsiaLang ? { eastAsiaLang: defaults.eastAsiaLang } : {}),
          },
        ]),
      )
    }
  }
  return max
}

/** half of the straddled outer vertical borders of a nested table (twips) */
function nestedBorderTwips(nt: TableModel): number {
  return ((borderTotalPt(nt.borders?.left) + borderTotalPt(nt.borders?.right)) / 2) * 20
}

/** narrowest width a nested table can take inside its cell: fixed grids keep
 * their declared width, pct tables follow the cell, autofit ones shrink to
 * every column's min-content (an empty cell keeps its side margins). Word
 * straddles the outer vertical borders, so the box is half a border wider on
 * each side (a regression sample: eight nesting levels, each parent column
 * = nested grid + 216 + 10 twips for sz 4 borders). */
export function nestedTableMinTwips(
  nt: TableModel,
  measure: TextMeasure,
  defaults?: RunDefaults,
): number {
  const widths = nt.colWidthsTwips
  if (!widths?.length || nt.widthPct) return 0
  const indent = tableIndentTwips(nt)
  const border = nestedBorderTwips(nt)
  // a nested dxa table keeps its declared width too: the parent column grows to
  // it and the other columns give the growth back (probe 2026-10-29: nested dxa
  // 6000 in a 5000 column -> 6226 = 6000 + 216 + border, siblings 2000/3000 -> 1536/2237)
  if (nt.fixedLayout || nt.dxaWidth) return widths.reduce((a, b) => a + b, 0) + indent + border
  const mins = minContentColTwips(nt, widths.length, measure, false, defaults)
  return (
    mins.reduce((sum, m, i) => sum + nestedFloorTwips(widths[i], m, sideMarginTwips(nt)), 0) +
    indent +
    border
  )
}

/** widest a nested table wants to be in its cell (twips): its grid plus indent
 * and the straddled outer borders; pct tables follow the cell (0) */
export function nestedTablePreferredTwips(nt: TableModel): number {
  const widths = nt.colWidthsTwips
  if (!widths?.length || nt.widthPct) return 0
  return widths.reduce((a, b) => a + b, 0) + tableIndentTwips(nt) + nestedBorderTwips(nt)
}

/** the single-span cells of a row that demand column width, with their grid column;
 * merged-away and vertical-text cells don't */
function* widthCells(model: TableModel, colCount: number): Generator<[TableCell, number]> {
  for (const row of model.rows) {
    let column = 0
    for (const cell of row) {
      const start = column
      column += cell.colSpan ?? 1
      if ((cell.colSpan ?? 1) !== 1 || start >= colCount) continue
      if (cell.vMerge === 'continue' || cell.textDirection) continue
      yield [cell, start]
    }
  }
}

/** per grid column (twips): widest unbreakable word in single-span cells, plus side
 * padding; `bare` drops the slack (a grid Word laid out already holds its words) */
export function minContentColTwips(
  model: TableModel,
  colCount: number,
  measure: TextMeasure,
  bare = false,
  defaults?: RunDefaults,
): number[] {
  const slack = bare ? 1 : measure.measured ? 1 : MIN_CONTENT_SLACK
  const mins = new Array<number>(colCount).fill(0)
  for (const [cell, start] of widthCells(model, colCount)) {
    const pad = cellSidePadTwips(cell, model)
    for (const nt of cell.nestedTables ?? []) {
      const ntMin = nestedTableMinTwips(nt, measure, defaults)
      if (ntMin > 0) mins[start] = Math.max(mins[start], ntMin + pad)
    }
    const wordPx = cellMeasurePx(cell, 'word', measure, defaults)
    if (wordPx <= 0) continue
    const twips =
      measure.measured && !bare ? pxGridTwips(wordPx, pad) : Math.ceil(wordPx * slack * 15) + pad
    mins[start] = Math.max(mins[start], twips)
  }
  return mins
}

/** per grid column (twips): widest unwrapped paragraph in single-span cells plus
 * side padding (max-content) */
export function maxContentColTwips(
  model: TableModel,
  colCount: number,
  measure: TextMeasure,
  defaults?: RunDefaults,
): number[] {
  const slack = measure.measured ? 1 : MIN_CONTENT_SLACK
  const maxes = new Array<number>(colCount).fill(0)
  for (const [cell, start] of widthCells(model, colCount)) {
    const pad = cellSidePadTwips(cell, model)
    for (const nt of cell.nestedTables ?? []) {
      const ntMax = Math.max(
        nestedTablePreferredTwips(nt),
        nestedTableMinTwips(nt, measure, defaults),
      )
      if (ntMax > 0) maxes[start] = Math.max(maxes[start], ntMax + pad)
    }
    const linePx = cellMeasurePx(cell, 'line', measure, defaults)
    if (linePx <= 0) continue
    const twips = measure.measured ? pxGridTwips(linePx, pad) : Math.ceil(linePx * slack * 15) + pad
    maxes[start] = Math.max(maxes[start], twips)
  }
  return maxes
}
