/**
 * Excel autofits a wrapped row to `lines x font row height`, where the font
 * row height is the single-line default row height of that font and size
 * (Calibri 10 -> 12.75 pt, Calibri 11 -> 15 pt, Arial 10 -> 12.75 pt). That
 * pitch is the OS/2 win ascent and descent, each rounded to whole pixels at
 * the GDI ppem, plus 2 px, except where Mac Excel (the fidelity reference)
 * is known to print a different value. Univer's measure
 * instead stacks the canvas font bounding box (Calibri's hhea pair is 1.0 em,
 * about 10 pt per line), so a five-line Calibri 10 cell comes out ~20 %
 * short. Rescale the measured plain-text wrap height to Excel's pitch; the
 * rendered line breaks are untouched.
 */
import {
  type ICellData,
  type IDocumentData,
  type IStyleData,
  type Nullable,
  WrapStrategy,
} from '@univerjs/core'
import {
  DEFAULT_PADDING_DATA,
  FontCache,
  getFontStyleString,
  SpreadsheetSkeleton,
} from '@univerjs/engine-render'

type FacePitch = readonly [ascent: number, descent: number, leading?: number]

function faces(metrics: FacePitch, ...names: string[]): Record<string, FacePitch> {
  return Object.fromEntries(names.map((name) => [name.toLowerCase(), metrics]))
}

/// usWinAscent / usWinDescent over unitsPerEm. Listed so the pitch does not
/// depend on which substitute the machine draws with; the ClearType set
/// (Calibri, Cambria, Candara, Consolas, Constantia, Corbel) also has hhea
/// metrics that differ from its win pair, so canvas metrics cannot stand in.
/// The optional third value is GDI's external leading (hhea line height beyond
/// the win pair), which Excel counts in place of its 2 px padding: the Yu
/// family carries a 0.5 em lineGap, hence its 18.75 pt default row.
const WIN_METRICS: Record<string, FacePitch> = {
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
  ...faces(
    [2229 / 2048, 495 / 2048],
    'Malgun Gothic',
    '\uB9D1\uC740 \uACE0\uB515',
    'Malgun Gothic Semilight',
  ),
  ...faces([2171 / 2048, 901 / 2048], 'Meiryo', '\u30E1\u30A4\u30EA\u30AA'),
  ...faces([2171 / 2048, 430 / 2048], 'Meiryo UI'),
  ...faces(
    [220 / 256, 36 / 256],
    'MS Gothic',
    'MS PGothic',
    'MS UI Gothic',
    'MS Mincho',
    'MS PMincho',
    '\uFF2D\uFF33 \u30B4\u30B7\u30C3\u30AF',
    '\uFF2D\uFF33 \uFF30\u30B4\u30B7\u30C3\u30AF',
    '\uFF2D\uFF33 \u660E\u671D',
    '\uFF2D\uFF33 \uFF30\u660E\u671D',
    'SimSun',
    'NSimSun',
    '\u5B8B\u4F53',
    '\u65B0\u5B8B\u4F53',
  ),
  ...faces(
    [2017 / 2048, 619 / 2048, 645 / 2048],
    'Yu Gothic',
    'Yu Gothic Medium',
    'Yu Gothic Light',
    '\u6E38\u30B4\u30B7\u30C3\u30AF',
    '\u6E38\u30B4\u30B7\u30C3\u30AF Medium',
    '\u6E38\u30B4\u30B7\u30C3\u30AF Light',
  ),
  ...faces(
    [2038 / 2048, 598 / 2048, 645 / 2048],
    'Yu Mincho',
    'Yu Mincho Demibold',
    'Yu Mincho Light',
    '\u6E38\u660E\u671D',
  ),
  ...faces([2210 / 2048, 514 / 2048], 'Yu Gothic UI'),
  ...faces(
    [879 / 1024, 145 / 1024],
    'Gulim',
    'GulimChe',
    'Dotum',
    'DotumChe',
    'Batang',
    'BatangChe',
    'Gungsuh',
    'GungsuhChe',
    '\uAD74\uB9BC',
    '\uAD74\uB9BC\uCCB4',
    '\uB3CB\uC6C0',
    '\uB3CB\uC6C0\uCCB4',
    '\uBC14\uD0D5',
    '\uBC14\uD0D5\uCCB4',
    '\uAD81\uC11C',
    '\uAD81\uC11C\uCCB4',
  ),
  ...faces(
    [1802 / 2048, 246 / 2048],
    'BIZ UDGothic',
    'BIZ UDPGothic',
    'BIZ UDMincho',
    'BIZ UDPMincho',
    'BIZ UD\u30B4\u30B7\u30C3\u30AF',
    'BIZ UDP\u30B4\u30B7\u30C3\u30AF',
    'BIZ UD\u660E\u671D',
    'BIZ UDP\u660E\u671D',
  ),
  ...faces(
    [1160 / 1000, 288 / 1000],
    'Noto Sans CJK',
    'Noto Sans CJK SC',
    'Noto Sans CJK TC',
    'Noto Sans CJK HK',
    'Noto Sans CJK JP',
    'Noto Sans CJK KR',
    'Source Han Sans',
    'Source Han Sans SC',
    'Source Han Sans CN',
    'Source Han Sans TC',
    'Source Han Sans TW',
    'Source Han Sans HK',
    'Source Han Sans JP',
    'Source Han Sans KR',
    '\u601D\u6E90\u9ED1\u4F53',
    '\u6E90\u30CE\u89D2\u30B4\u30B7\u30C3\u30AF',
    '\uBCF8\uACE0\uB515',
  ),
  ...faces(
    [2167 / 2048, 536 / 2048],
    'Microsoft YaHei',
    'Microsoft YaHei UI',
    '\u5FAE\u8F6F\u96C5\u9ED1',
  ),
}

const ROW_PADDING_PX = 2
/// Univer's default section linePitch: a cell document never lays a line
/// out shorter than this, whatever the font box.
const DOC_LINE_PITCH_PX = 15.6

/// Where Mac Excel prints a pitch the GDI formula does not give: reference
/// prints of Arial 11 sheets show 15 pt rows and 30 pt two-line wraps (the
/// formula says 14.25 pt), Calibri 12 sheets 15 pt rows (formula 15.75 pt).
const MAC_ROW_PITCH_PX: Record<string, Record<number, number>> = {
  arial: { 11: 20 },
  'liberation sans': { 11: 20 },
  calibri: { 12: 20 },
  carlito: { 12: 20 },
}

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
  const key = familyKey(family)
  const calibrated = MAC_ROW_PITCH_PX[key]?.[sizePt]
  if (calibrated !== undefined) return calibrated
  const px = (sizePt * 96) / 72
  const ppem = Math.round(px)
  const [ascent, descent, leading = 0] = WIN_METRICS[key] ?? [
    metrics.fontBoundingBoxAscent / px,
    metrics.fontBoundingBoxDescent / px,
  ]
  if (!(ascent + descent > 0)) return 0
  const padding = Math.max(ROW_PADDING_PX, Math.round(leading * ppem))
  return Math.round(ascent * ppem) + Math.round(descent * ppem) + padding
}

/// Converts Univer's wrap measure (lines x line box + padding) into
/// `lines x Excel pitch`; `isDoc` for cell documents, whose line box is
/// floored at the section linePitch.
export function excelWrapHeight(measured: number, style: IStyleData, isDoc = false): number {
  const { fontCache, fontFamily, fontSize } = getFontStyleString(style)
  const metrics = FontCache.getMeasureText('A', fontCache)
  const lineHeight = metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent
  if (!(lineHeight > 0)) return measured
  const padding = (style.pd?.t ?? DEFAULT_PADDING_DATA.t) + (style.pd?.b ?? DEFAULT_PADDING_DATA.b)
  const lineBox = isDoc ? Math.max(lineHeight, DOC_LINE_PITCH_PX) : lineHeight
  const lines = Math.round((measured - padding) / lineBox)
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

/// The cell style carrying the one font of a cell document (manual line
/// breaks, uniformly formatted rich text); null when runs mix families or
/// sizes.
function uniformDocStyle(doc: IDocumentData, style: IStyleData): IStyleData | null {
  const body = doc.body
  if (!body || body.tables?.length || body.customBlocks?.length) return null
  const fonts = new Map<string, { ff: Nullable<string>; fs: number | undefined }>()
  const add = (ff = style.ff, fs = style.fs) => fonts.set(`${ff ?? ''}|${fs ?? ''}`, { ff, fs })
  let covered = 0
  for (const run of body.textRuns ?? []) {
    if (run.ts?.va) return null
    add(run.ts?.ff, run.ts?.fs)
    covered += run.ed - run.st
  }
  if (covered < body.dataStream.replace(/\r\n$/, '').length) add()
  if (fonts.size !== 1) return null
  const merged: IStyleData = { ...style }
  for (const { ff, fs } of fonts.values()) {
    if (ff != null) merged.ff = ff
    if (fs != null) merged.fs = fs
  }
  return merged
}

/// Only wrapped cells in a single font are rescaled: plain text, and
/// documents (manual line breaks, uniform rich text) whose runs agree on
/// family and size. Mixed-font rich text, rotated text and
/// interceptor-provided heights keep Univer's value.
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
    if (!cell || cell.interceptorAutoHeight) return measured
    const style = this.worksheet.getComposedCellStyleByCellData(row, col, cell)
    if (!style || style.tb !== WrapStrategy.WRAP || style.tr?.a || style.tr?.v) return measured
    if (cell.p) {
      const docStyle = uniformDocStyle(cell.p, style)
      return docStyle ? excelWrapHeight(measured, docStyle, true) : measured
    }
    if (cell.v === undefined || cell.v === null) return measured
    return excelWrapHeight(measured, style)
  }
}
