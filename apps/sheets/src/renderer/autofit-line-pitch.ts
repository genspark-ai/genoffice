/**
 * Excel autofits a wrapped row to `lines x font row height`, where the font
 * row height is the single-line default row height of that font and size
 * (Calibri 10 -> 12.75 pt, Calibri 11 -> 15 pt, Calibri 12 -> 15.75 pt,
 * Arial 10 -> 12.75 pt). That pitch is the OS/2 win ascent and descent,
 * each rounded to whole pixels at the GDI ppem, plus 2 px. Univer's measure
 * instead stacks the canvas font bounding box (Calibri's hhea pair is 1.0 em,
 * about 10 pt per line), so a five-line Calibri 10 cell comes out ~20 %
 * short. Rescale the measured plain-text wrap height to Excel's pitch; the
 * rendered line breaks are untouched.
 */
import { type ICellData, type IStyleData, type Nullable, WrapStrategy } from '@univerjs/core'
import {
  DEFAULT_PADDING_DATA,
  FontCache,
  getFontStyleString,
  SpreadsheetSkeleton,
} from '@univerjs/engine-render'

/// usWinAscent / usWinDescent over unitsPerEm. Listed so the pitch does not
/// depend on which substitute the machine draws with; the ClearType set
/// (Calibri, Cambria, Candara, Consolas, Constantia, Corbel) also has hhea
/// metrics that differ from its win pair, so canvas metrics cannot stand in.
const WIN_METRICS: Record<string, readonly [ascent: number, descent: number]> = {
  calibri: [1950 / 2048, 550 / 2048],
  carlito: [1950 / 2048, 550 / 2048],
  cambria: [1946 / 2048, 455 / 2048],
  caladea: [1946 / 2048, 455 / 2048],
  candara: [1950 / 2048, 550 / 2048],
  consolas: [1884 / 2048, 514 / 2048],
  constantia: [1950 / 2048, 550 / 2048],
  corbel: [1950 / 2048, 550 / 2048],
  arial: [1854 / 2048, 434 / 2048],
  'liberation sans': [1854 / 2048, 434 / 2048],
  'times new roman': [1825 / 2048, 443 / 2048],
  'liberation serif': [1825 / 2048, 443 / 2048],
  verdana: [2059 / 2048, 430 / 2048],
  tahoma: [2049 / 2048, 423 / 2048],
  aptos: [2068 / 2048, 563 / 2048],
  'aptos narrow': [2068 / 2048, 563 / 2048],
}

const ROW_PADDING_PX = 2

export interface LineMetrics {
  fontBoundingBoxAscent: number
  fontBoundingBoxDescent: number
}

function familyKey(family: string): string {
  const first = family.split(',')[0] ?? ''
  return first
    .trim()
    .replace(/^['"]|['"]$/g, '')
    .toLowerCase()
}

/// Excel's single-line row height in px for a font and size; `metrics` is the
/// canvas font bounding box at that size and only serves unlisted families.
export function excelRowPitchPx(family: string, sizePt: number, metrics: LineMetrics): number {
  const px = (sizePt * 96) / 72
  const ppem = Math.round(px)
  const [ascent, descent] = WIN_METRICS[familyKey(family)] ?? [
    metrics.fontBoundingBoxAscent / px,
    metrics.fontBoundingBoxDescent / px,
  ]
  if (!(ascent + descent > 0)) return 0
  return Math.round(ascent * ppem) + Math.round(descent * ppem) + ROW_PADDING_PX
}

/// Converts Univer's plain-text wrap measure (lines x canvas line box +
/// padding) into `lines x Excel pitch`.
export function excelWrapHeight(measured: number, style: IStyleData): number {
  const { fontCache, fontFamily, fontSize } = getFontStyleString(style)
  const metrics = FontCache.getMeasureText('A', fontCache)
  const lineHeight = metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent
  if (!(lineHeight > 0)) return measured
  const padding = (style.pd?.t ?? DEFAULT_PADDING_DATA.t) + (style.pd?.b ?? DEFAULT_PADDING_DATA.b)
  const lines = Math.round((measured - padding) / lineHeight)
  if (lines < 1) return measured
  const pitch = excelRowPitchPx(fontFamily, style.fs ?? fontSize, metrics)
  return pitch > 0 ? lines * pitch : measured
}

export interface PitchSkeletonLike {
  worksheet: {
    getCell(row: number, col: number): Nullable<ICellData>
    getComposedCellStyleByCellData(row: number, col: number, cell: Nullable<ICellData>): IStyleData
  }
  calculateAutoHeightForCell(row: number, col: number): number | undefined
}

const patched = new WeakSet<PitchSkeletonLike>()

/// Only the plain-text wrap branch of the stock measure is rescaled; rich
/// text, rotated text and interceptor-provided heights keep Univer's value.
export function installAutofitLinePitch(
  proto: PitchSkeletonLike = SpreadsheetSkeleton.prototype as unknown as PitchSkeletonLike,
): void {
  if (patched.has(proto)) return
  patched.add(proto)
  const original = proto.calculateAutoHeightForCell
  proto.calculateAutoHeightForCell = function (this: PitchSkeletonLike, row, col) {
    const measured = original.call(this, row, col)
    if (!measured) return measured
    const cell = this.worksheet.getCell(row, col) as
      (ICellData & { interceptorAutoHeight?: unknown }) | null | undefined
    if (!cell || cell.p || cell.v === undefined || cell.v === null || cell.interceptorAutoHeight) {
      return measured
    }
    const style = this.worksheet.getComposedCellStyleByCellData(row, col, cell)
    if (!style || style.tb !== WrapStrategy.WRAP || style.tr?.a || style.tr?.v) return measured
    return excelWrapHeight(measured, style)
  }
}
