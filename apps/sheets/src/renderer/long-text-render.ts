/**
 * Plain (non-wrapped, unrotated) text cells are drawn with one `fillText` of
 * the whole string, clipped to the overflow box. Chromium shapes the entire
 * string on every frame, so a grid of cells holding kilobyte-long paragraphs
 * (prod: 13 columns × thousands of rows of 1-30k-char text) costs ~300 ms per
 * scroll frame while only a few dozen characters are ever visible. Hand the
 * renderer just the slice that can reach the clip box, keeping the edge the
 * alignment anchors on.
 */
import {
  getNumfmtParseValueFilter,
  HorizontalAlign,
  isRealNum,
  Worksheet,
  WrapStrategy,
  type ICellData,
  type IDocumentData,
  type IStyleData,
  type ITextRun,
} from '@univerjs/core'
import {
  DocumentSkeleton,
  DocumentViewModel,
  Font,
  FontCache,
  SpreadsheetSkeleton,
  getDocsSkeletonPageSize,
} from '@univerjs/engine-render'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import { resolveBidiDirection } from './rtl-text-fix'

/** Below this the shaping cost is negligible and slicing would only add work. */
export const LONG_TEXT_MIN_CHARS = 256

/** Proportional fonts make the average-width estimate rough; keep this much spare. */
const SLACK = 1.5
const SLACK_CHARS = 32

/**
 * Which end of a long single line touches the aligned edge. For LTR text the
 * logical head sits at the left edge and the tail at the right; RTL text
 * mirrors that, so a right-aligned Arabic paragraph shows its logical head.
 */
function keepStart(text: string, keep: number, hAlign: number | undefined): number {
  if (hAlign === HorizontalAlign.CENTER) return Math.floor((text.length - keep) / 2)
  const rtl = resolveBidiDirection(text) === 'rtl'
  const tailAtEdge =
    hAlign === HorizontalAlign.RIGHT ? !rtl : hAlign === HorizontalAlign.LEFT && rtl
  return tailAtEdge ? text.length - keep : 0
}

export function visibleSlice(
  text: string,
  textWidth: number,
  availableWidth: number,
  hAlign: number | undefined,
): string {
  if (text.length <= LONG_TEXT_MIN_CHARS || !(textWidth > 0) || !(availableWidth > 0)) return text
  const avgChar = textWidth / text.length
  const keep = Math.ceil((availableWidth / avgChar) * SLACK) + SLACK_CHARS
  if (keep >= text.length) return text
  const start = keepStart(text, keep, hAlign)
  return text.slice(start, start + keep)
}

interface FontCacheLike {
  cellData?: { v?: unknown; p?: unknown; t?: number } | null
  fontString?: string
  wrapStrategy?: number
  vertexAngle?: number
  horizontalAlign?: number
}

interface RenderFontCtxLike {
  startX: number
  endX: number
  fontCache?: FontCacheLike
  overflowRectangle?: { startColumn: number; endColumn: number } | null
  spreadsheetSkeleton?: { columnWidthAccumulation?: number[] }
}

function clipWidth(ctx: RenderFontCtxLike): number {
  const rect = ctx.overflowRectangle
  const acc = ctx.spreadsheetSkeleton?.columnWidthAccumulation
  const right = rect && acc ? acc[rect.endColumn] : undefined
  if (rect && acc && right !== undefined) {
    const left = rect.startColumn > 0 ? (acc[rect.startColumn - 1] ?? 0) : 0
    return right - left
  }
  return ctx.endX - ctx.startX
}

/**
 * Rich-text documents (our importer builds one for every cell holding manual
 * line breaks, joined or not) are laid out glyph by glyph for the render cache
 * and again for auto row height. An unwrapped single paragraph only ever shows
 * what fits across its overflow box, so cap the text handed to the layout and
 * let the overflow pass widen the few cells whose box turns out wider.
 */
export const LONG_DOC_MAX_CHARS = 256

type CellDoc = IDocumentData

export function truncateDocForLayout(
  doc: CellDoc,
  hAlign: number | undefined,
  maxChars = LONG_DOC_MAX_CHARS,
): CellDoc {
  const body = doc.body
  const stream = body?.dataStream
  if (!body || !stream) return doc
  // Trailing "\r\n" is the paragraph + section terminator, not content.
  const textLen = stream.length - 2
  if (textLen <= maxChars) return doc
  const breaks = (body.paragraphs ?? []).filter(
    (para: { startIndex: number }) => para.startIndex < textLen,
  )
  if (breaks.length > 0) return doc
  const start = keepStart(stream.slice(0, textLen), maxChars, hAlign)
  const end = start + maxChars
  const textRuns = (body.textRuns ?? [])
    .filter((run: ITextRun) => run.ed > start && run.st < end)
    .map((run: ITextRun) => ({
      ...run,
      st: Math.max(run.st, start) - start,
      ed: Math.min(run.ed, end) - start,
    }))
  // Ranges, decorations, blocks and tables carry offsets into the full stream.
  const { customRanges: _cr, customDecorations: _cd, customBlocks: _cb, tables: _t, ...rest } = body
  return {
    ...doc,
    body: {
      ...rest,
      dataStream: `${stream.slice(start, end)}\r\n`,
      textRuns,
      paragraphs: [{ startIndex: maxChars }],
      sectionBreaks: [{ startIndex: maxChars + 1 }],
    },
  }
}

/**
 * Auto row height of an unwrapped single-line paragraph does not depend on its
 * length (as long as one run sets the font), so that pass can keep far less.
 */
export const AUTO_HEIGHT_DOC_CHARS = 64

/** Plain strings longer than this are measured by prefix for overflow. */
export const OVERFLOW_MEASURE_CHARS = 2048

let autoHeightPass = false

function uniformRuns(doc: CellDoc): boolean {
  const runs = doc.body?.textRuns ?? []
  if (runs.length <= 1) return true
  const fs = runs[0]?.ts?.fs
  const ff = runs[0]?.ts?.ff
  return runs.every((run: ITextRun) => run.ts?.fs === fs && run.ts?.ff === ff)
}

/** Cells whose layout document was cut, so the overflow pass can widen them. */
const truncatedDocs = new WeakMap<object, { fullLength: number; cap: number }>()
let capOverride: number | null = null

function layoutCellForDoc(cell: ICellData, style: IStyleData | undefined): ICellData {
  const doc = cell.p
  if (!doc?.body?.dataStream) return cell
  // Wrapped, rotated and stacked-vertical text all depend on the whole string.
  if (style?.tb === WrapStrategy.WRAP || style?.tr?.a || style?.tr?.v) return cell
  const cap =
    capOverride ?? (autoHeightPass && uniformRuns(doc) ? AUTO_HEIGHT_DOC_CHARS : LONG_DOC_MAX_CHARS)
  const fullLength = doc.body.dataStream.length - 2
  if (fullLength <= cap) return cell
  const truncated = truncateDocForLayout(doc, style?.ht ?? undefined, cap)
  if (truncated === doc) return cell
  if (!autoHeightPass) truncatedDocs.set(cell, { fullLength, cap })
  return { ...cell, p: truncated }
}

interface OverflowDocsConfig {
  documentSkeleton?: DocumentSkeleton
  cellData?: (ICellData & { s?: IStyleData | string | null }) | undefined
  style?: IStyleData
  horizontalAlign?: number
}

/** Scene pixels a spill can stay visible across: a 1600 px viewport at 25 % zoom shows 6400. */
export const LAYOUT_SPAN_SCENE_PX = 8000

/**
 * The cut document is narrower than the real text, so ask the skeleton how far
 * this cell could spill (up to the layout span, stopped by non-empty
 * neighbours; the merge box for merged cells) and, when that is wider than the
 * laid-out slice, lay out enough of the string to fill it before the overflow
 * pass runs.
 */
function widenTruncatedDoc(
  skeleton: any,
  row: number,
  column: number,
  docsConfig: OverflowDocsConfig,
): boolean {
  const cell = docsConfig.cellData
  const cut = cell ? truncatedDocs.get(cell) : undefined
  if (!cell || !cut || !docsConfig.documentSkeleton) return false
  const docSize = getDocsSkeletonPageSize(docsConfig.documentSkeleton)
  if (!docSize || !(docSize.width > 0)) return false
  const docWidth = docSize.width
  let boxWidth: number
  const coord = skeleton.getCellWithCoordByIndex?.(row, column)
  if (coord?.isMerged || coord?.isMergedMainCell) {
    boxWidth = coord.mergeInfo ? coord.mergeInfo.endX - coord.mergeInfo.startX : 0
  } else {
    const acc: number[] | undefined = skeleton.columnWidthAccumulation
    if (!acc) return false
    const spanFor = (align: number): number => {
      const span = skeleton.getOverflowPosition?.(
        { width: LAYOUT_SPAN_SCENE_PX, height: docSize.height },
        align,
        row,
        column,
        skeleton.getColumnCount(),
      )
      if (!span) return 0
      const left = span.startColumn > 0 ? (acc[span.startColumn - 1] ?? 0) : 0
      return (acc[span.endColumn] ?? 0) - left
    }
    const align = docsConfig.horizontalAlign ?? HorizontalAlign.UNSPECIFIED
    // General alignment resolves per text direction later (RTL text spills
    // leftwards), so probe both sides and take the wider box.
    boxWidth =
      align === HorizontalAlign.UNSPECIFIED
        ? Math.max(spanFor(HorizontalAlign.LEFT), spanFor(HorizontalAlign.RIGHT))
        : spanFor(align)
  }
  if (boxWidth <= docWidth) return false
  const needed = Math.min(cut.fullLength, Math.ceil((cut.cap * boxWidth) / docWidth) + 32)
  if (needed <= cut.cap) return false
  const style =
    docsConfig.style ?? (typeof cell.s === 'object' ? (cell.s as IStyleData) : undefined)
  capOverride = needed
  let model
  try {
    model = skeleton.worksheet.getCellDocumentModel(cell, style, {
      displayRawFormula: skeleton._renderRawFormula,
    })
  } finally {
    capOverride = null
  }
  if (!model?.documentModel) return false
  const widened = DocumentSkeleton.create(
    new DocumentViewModel(model.documentModel),
    skeleton._localeService,
  )
  widened.calculate()
  docsConfig.documentSkeleton = widened
  truncatedDocs.set(cell, { fullLength: cut.fullLength, cap: needed })
  return true
}

let installed = false

export function installLongTextRender(): void {
  if (installed) return
  installed = true
  const worksheetProto = Worksheet.prototype as any
  const origDocModel = worksheetProto.getCellDocumentModel
  // Both the render cache (_setFontStylesCache) and calculateAutoHeightForCell
  // build their document through this one entry point.
  worksheetProto.getCellDocumentModel = function (
    cell: ICellData | undefined,
    style: IStyleData | undefined,
    options: unknown,
  ) {
    return origDocModel.call(this, cell ? layoutCellForDoc(cell, style) : cell, style, options)
  }
  const skeletonProto = SpreadsheetSkeleton.prototype as any
  // The overflow pass measures the whole plain string to find how far it
  // spills; past a few thousand characters the spill already covers any
  // viewport, so measure a prefix instead of shaping the full paragraph.
  const origOverflow = skeletonProto._calculateOverflowCell
  skeletonProto._calculateOverflowCell = function (
    row: number,
    column: number,
    docsConfig: OverflowDocsConfig,
  ) {
    const cellData = docsConfig?.cellData
    const v = cellData?.v
    if (docsConfig.documentSkeleton) {
      widenTruncatedDoc(this, row, column, docsConfig)
      return origOverflow.call(this, row, column, docsConfig)
    }
    if (!cellData || typeof v !== 'string' || v.length <= OVERFLOW_MEASURE_CHARS) {
      return origOverflow.call(this, row, column, docsConfig)
    }
    docsConfig.cellData = { ...cellData, v: v.slice(0, OVERFLOW_MEASURE_CHARS) }
    try {
      return origOverflow.call(this, row, column, docsConfig)
    } finally {
      docsConfig.cellData = cellData
    }
  }
  const origAutoHeight = skeletonProto.calculateAutoHeightForCell
  skeletonProto.calculateAutoHeightForCell = function (row: number, col: number) {
    autoHeightPass = true
    try {
      return origAutoHeight.call(this, row, col)
    } finally {
      autoHeightPass = false
    }
  }
  const fontProto = Font.prototype as any
  const orig = fontProto._renderText
  fontProto._renderText = function (
    ctx: unknown,
    row: number,
    col: number,
    renderFontCtx: RenderFontCtxLike,
    overflowCache: unknown,
  ) {
    const cache = renderFontCtx?.fontCache
    const value = cache?.cellData?.v
    if (
      typeof value !== 'string' ||
      value.length <= LONG_TEXT_MIN_CHARS ||
      cache?.cellData?.p ||
      cache?.wrapStrategy === WrapStrategy.WRAP ||
      (cache?.vertexAngle ?? 0) !== 0 ||
      !cache?.fontString
    ) {
      return orig.call(this, ctx, row, col, renderFontCtx, overflowCache)
    }
    // The overflow pass already measured this exact string, so this is a cache hit.
    const measured = FontCache.getMeasureText(value, cache.fontString)
    const slice = visibleSlice(
      value,
      measured.width,
      clipWidth(renderFontCtx),
      cache.horizontalAlign,
    )
    if (slice === value) return orig.call(this, ctx, row, col, renderFontCtx, overflowCache)
    renderFontCtx.fontCache = { ...cache, cellData: { ...cache.cellData, v: slice } }
    try {
      return orig.call(this, ctx, row, col, renderFontCtx, overflowCache)
    } finally {
      renderFontCtx.fontCache = cache
    }
  }
}

/**
 * Univer's force-string marker (the green "number stored as text" corner) runs
 * isRealNum + numfmt's date/time parser on the raw value of every string cell,
 * on every cell read. Paragraph-sized strings make that the top scroll cost,
 * so gate the handler behind a memoized "could this be numeric" check.
 */
const NUMERIC_LIKE_MAX_CHARS = 64
const numericLikeCache = new Map<string, boolean>()

export function numericLike(value: string): boolean {
  if (value.length > NUMERIC_LIKE_MAX_CHARS) return false
  const cached = numericLikeCache.get(value)
  if (cached !== undefined) return cached
  const result = isRealNum(value) || getNumfmtParseValueFilter(value) != null
  if (numericLikeCache.size >= 50_000) numericLikeCache.clear()
  numericLikeCache.set(value, result)
  return result
}

type CellHandler = (
  cell: unknown,
  pos: { rawData?: ICellData | null },
  next: (c: unknown) => unknown,
) => unknown
interface CellInterceptor {
  handler: CellHandler
  priority?: number
}

const isForceStringInterceptor = (interceptor: CellInterceptor): boolean =>
  typeof interceptor.handler === 'function' &&
  String(interceptor.handler).includes('disableForceStringMark')

function gateForceStringHandler(interceptor: CellInterceptor): void {
  const inner = interceptor.handler
  interceptor.handler = (cell, pos, next) => {
    const raw = pos?.rawData?.v
    if (typeof raw === 'string' && !numericLike(raw)) return next(cell)
    return inner(cell, pos, next)
  }
}

let forceStringInstalled = false

/** Wraps the force-string interceptor already registered and any registered later. */
export function installForceStringMarkGate(service: SheetInterceptorService): void {
  const registered = (service as any)._interceptorsByName?.get(INTERCEPTOR_POINT.CELL_CONTENT) as
    CellInterceptor[] | undefined
  for (const interceptor of registered ?? []) {
    if (isForceStringInterceptor(interceptor)) gateForceStringHandler(interceptor)
  }
  if (forceStringInstalled) return
  forceStringInstalled = true
  const proto = SheetInterceptorService.prototype as any
  const origIntercept = proto.intercept
  proto.intercept = function (name: unknown, interceptor: CellInterceptor) {
    if (name === INTERCEPTOR_POINT.CELL_CONTENT && isForceStringInterceptor(interceptor)) {
      gateForceStringHandler(interceptor)
    }
    return origIntercept.call(this, name, interceptor)
  }
}
