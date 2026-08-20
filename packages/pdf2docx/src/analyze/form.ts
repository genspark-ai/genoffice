/**
 * Checkbox form-table detection (P6). Form templates (NIST SSP/BIA and the
 * like) lay out FORMCHECKBOX rows as borderless tables: one row of
 * "☐ Label   ☐ Label   ☐ Label" cells, often followed by a full-width
 * description row merged across the columns. Lattice detection cannot see
 * them (no cell borders beyond the checkbox squares themselves) and the
 * stream detector rejects single rows by design, so this pass keys on the
 * strongest evidence available: a small drawn square hugging each unit's
 * left edge. The squares also become visible again as '☐' glyphs — the PDF
 * draws them as vector paths that would otherwise be dropped.
 *
 * Suppression bias stays "miss rather than misfire": EVERY unit of the row
 * must carry its own checkbox, and description rows join only while the
 * vertical gap stays clearly intra-table.
 */
import type { Rect } from '../geometry'
import { intersectArea, median, rectUnion, rectUnionAll, verticalOverlapRatio } from '../geometry'
import type { PageShapes, PdfChar, Stroke, TableBlock, TableCellBlock } from '../ir'
import { scriptOf } from '../script'
import { groupIntoBlocks } from './blocks'
import { analyzeChars } from './chars'
import { isSpaceCode } from './lines'
import { groupStrokes } from './table'
import type { LineUnit } from './units'
import { clusterUnitRows, type UnitRow } from './units'

/** checkbox square side, points (Word/LO render form checkboxes at ~9–13pt) */
const CHECKBOX_MIN_PT = 5
const CHECKBOX_MAX_PT = 18
/** |w − h| beyond this share of the longer side is not a square */
const CHECKBOX_ASPECT_TOL = 0.35
/** unit left edge must start within this many ems right of its square */
const CHECKBOX_GAP_MAX_EMS = 1.2
/** …and may overlap it at most this many points (glyph boxes jitter) */
const CHECKBOX_GAP_MIN_PT = -1
/** square and unit must share this much vertical extent */
const CHECKBOX_V_OVERLAP_MIN = 0.5
/** a form row needs at least this many checkbox-led cells */
const FORM_MIN_COLS = 2
/** consecutive form rows further apart than this × row height stay separate tables */
const FORM_ROW_GAP_MAX_RATIO = 1.6
/** a description row joins while its gap stays under this × form-row height */
const CONT_GAP_MAX_RATIO = 0.75
/** …and its left edge sits within this many ems of the table's left edge */
const CONT_LEFT_TOL_EMS = 2
/** at most this many description rows join one table */
const CONT_MAX_ROWS = 4
/** checkbox evidence is strong — well above the stream detector's threshold */
const FORM_CONFIDENCE = 0.9

/** Small stroked squares (checkbox outlines) among the page's strokes. */
export function detectCheckboxSquares(strokes: readonly Stroke[]): Rect[] {
  const squares: Rect[] = []
  for (const group of groupStrokes(strokes)) {
    const h = group.filter((s) => s.orientation === 'h')
    const v = group.filter((s) => s.orientation === 'v')
    // a drawn rect outline: both edge pairs, nothing much beyond them
    if (h.length < 2 || v.length < 2 || group.length > 8) continue
    const box = rectUnionAll(group.map((s) => s.box))
    const w = box.x1 - box.x0
    const hgt = box.y1 - box.y0
    if (w < CHECKBOX_MIN_PT || w > CHECKBOX_MAX_PT) continue
    if (hgt < CHECKBOX_MIN_PT || hgt > CHECKBOX_MAX_PT) continue
    if (Math.abs(w - hgt) > CHECKBOX_ASPECT_TOL * Math.max(w, hgt)) continue
    squares.push(box)
  }
  return squares
}

const isVisible = (c: PdfChar): boolean => !isSpaceCode(c.code) && c.code > 0x1f

/** the '☐' glyph standing in for a drawn checkbox square */
function checkboxChar(square: Rect, unit: LineUnit): PdfChar {
  const anchor = unit.chars.find(isVisible) ?? unit.chars[0]!
  return {
    code: 0x2610,
    text: '☐',
    box: square,
    looseBox: square,
    originX: square.x0,
    originY: unit.baseline,
    angle: 0,
    fontSize: anchor.fontSize,
    fontWeight: anchor.fontWeight,
    fontFamily: anchor.fontFamily,
    italic: false,
    color: anchor.color,
    isGenerated: true, // synthesized — not in the content stream
    isHyphen: false,
    script: scriptOf(0x2610),
  }
}

/** the checkbox square hugging the unit's left edge, if any */
function squareOf(unit: LineUnit, squares: readonly Rect[]): Rect | null {
  for (const square of squares) {
    const gap = unit.box.x0 - square.x1
    if (gap < CHECKBOX_GAP_MIN_PT || gap > CHECKBOX_GAP_MAX_EMS * unit.fontSize) continue
    if (verticalOverlapRatio(square, unit.box) < CHECKBOX_V_OVERLAP_MIN) continue
    return square
  }
  return null
}

/** a qualifying form row: every unit is checkbox-led */
interface FormRow {
  row: UnitRow
  squares: Rect[]
  /** full extent including the squares */
  box: Rect
}

function formRowOf(row: UnitRow, squares: readonly Rect[]): FormRow | null {
  if (row.units.length < FORM_MIN_COLS) return null
  const matched: Rect[] = []
  for (const unit of row.units) {
    const square = squareOf(unit, squares)
    if (!square) return null
    matched.push(square)
  }
  return { row, squares: matched, box: rectUnion(row.box, rectUnionAll(matched)) }
}

/** one detected table: consecutive form rows + trailing description rows */
interface FormCandidate {
  formRows: FormRow[]
  contRows: UnitRow[]
}

function buildFormTable(cand: FormCandidate): TableBlock {
  const cols = cand.formRows[0]!.row.units.length
  // per-column extent (square left edge → unit right edge) across form rows
  const colSpans: Array<{ lo: number; hi: number }> = Array.from({ length: cols }, (_, c) => ({
    lo: Math.min(...cand.formRows.map((f) => f.squares[c]!.x0)),
    hi: Math.max(...cand.formRows.map((f) => f.row.units[c]!.box.x1)),
  }))
  const box = rectUnionAll([...cand.formRows.map((f) => f.box), ...cand.contRows.map((r) => r.box)])
  const xs: number[] = [box.x0]
  for (let c = 1; c < cols; c++) xs.push((colSpans[c - 1]!.hi + colSpans[c]!.lo) / 2)
  xs.push(box.x1)

  const rows: TableCellBlock[][] = cand.formRows.map((f) => {
    const rowBox = f.box
    return f.row.units.map((unit, c) => {
      const cellBox: Rect = { x0: xs[c]!, x1: xs[c + 1]!, y0: rowBox.y0, y1: rowBox.y1 }
      const chars = [checkboxChar(f.squares[c]!, unit), ...unit.chars]
      return { box: cellBox, gridSpan: 1, blocks: groupIntoBlocks(analyzeChars(chars)) }
    })
  })
  if (cand.contRows.length > 0) {
    const contBox = rectUnionAll(cand.contRows.map((r) => r.box))
    const cellBox: Rect = { x0: box.x0, x1: box.x1, y0: contBox.y0, y1: contBox.y1 }
    const chars = cand.contRows.flatMap((r) => r.units.flatMap((u) => u.chars))
    rows.push([{ box: cellBox, gridSpan: cols, blocks: groupIntoBlocks(analyzeChars(chars)) }])
  }
  const colWidthsPt = xs.slice(1).map((x, i) => x - xs[i]!)
  return { kind: 'table', box, colWidthsPt, rows, confidence: FORM_CONFIDENCE, form: true }
}

export interface DetectedFormTables {
  tables: TableBlock[]
  /** units that stay in the regular flow */
  remainingUnits: LineUnit[]
}

/**
 * Full checkbox-form pass over one page's units. `excludeBoxes` are already
 * detected (lattice) table regions — squares inside them are cell borders'
 * business, not form evidence.
 */
export function detectFormTables(
  units: readonly LineUnit[],
  shapes: PageShapes,
  excludeBoxes: readonly Rect[] = [],
): DetectedFormTables {
  const squares = detectCheckboxSquares(shapes.strokes).filter(
    (sq) => !excludeBoxes.some((b) => intersectArea(sq, b) > 0),
  )
  if (squares.length === 0) return { tables: [], remainingUnits: [...units] }

  const rows = clusterUnitRows(units)
  const tables: TableBlock[] = []
  const consumed = new Set<LineUnit>()

  for (let i = 0; i < rows.length; i++) {
    const first = formRowOf(rows[i]!, squares)
    if (!first) continue

    // consecutive same-width form rows stack into one grid
    const formRows: FormRow[] = [first]
    let j = i + 1
    while (j < rows.length) {
      const prev = formRows[formRows.length - 1]!
      const next = formRowOf(rows[j]!, squares)
      if (!next || next.row.units.length !== first.row.units.length) break
      const gap = prev.box.y0 - next.box.y1
      if (gap > FORM_ROW_GAP_MAX_RATIO * Math.max(prev.box.y1 - prev.box.y0, 1)) break
      formRows.push(next)
      j++
    }

    // tight-following full-width description rows join as a merged row
    const formRowH = median(formRows.map((f) => f.row.box.y1 - f.row.box.y0)) || 12
    const tableLeft = Math.min(...formRows.map((f) => f.box.x0))
    const contRows: UnitRow[] = []
    while (j < rows.length && contRows.length < CONT_MAX_ROWS) {
      const row = rows[j]!
      if (row.units.length !== 1) break
      const unit = row.units[0]!
      if (squareOf(unit, squares)) break // a new checkbox row, not a description
      const above = contRows[contRows.length - 1]?.box ?? formRows[formRows.length - 1]!.box
      if (above.y0 - row.box.y1 > CONT_GAP_MAX_RATIO * formRowH) break
      if (Math.abs(unit.box.x0 - tableLeft) > CONT_LEFT_TOL_EMS * unit.fontSize) break
      contRows.push(row)
      j++
    }

    tables.push(buildFormTable({ formRows, contRows }))
    for (const f of formRows) for (const u of f.row.units) consumed.add(u)
    for (const r of contRows) for (const u of r.units) consumed.add(u)
    i = j - 1
  }

  tables.sort((a, b) => b.box.y1 - a.box.y1)
  return { tables, remainingUnits: units.filter((u) => !consumed.has(u)) }
}
