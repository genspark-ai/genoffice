/**
 * Excel resolves a cell's reading order from its content ("context" reading
 * order): the first strong directional character decides the paragraph
 * direction, so "1446<AR>" renders with the Arabic suffix on the LEFT and the
 * digits on the RIGHT. Univer's fast text path draws every line with the
 * canvas default (ltr) paragraph direction, which mirrors such mixed-direction
 * strings. Detect the paragraph direction per UAX#9 P2/P3 (first strong
 * character) and hand it to the canvas bidi algorithm via ctx.direction.
 *
 * textAlign is pinned to 'left' while drawing: the fast path computes line x
 * offsets for a left anchor, and the canvas default 'start' would flip the
 * anchor to the right edge under direction:'rtl'.
 *
 * The same context rule also drives Excel's General HORIZONTAL alignment: a
 * cell whose first strong character is RTL right-aligns (in LTR and RTL
 * sheets alike; ref.pdf verified). Univer resolves General as number→right /
 * boolean→center / else left, so RTL-first text cells are patched to RIGHT at
 * the font-cache choke point (covers paint, overflow direction and clipping),
 * in the document component's page handler (rich text / wrap path) and in the
 * editor bridge state (in-cell editor margin, hence caret and hit-testing).
 */
import { CellValueType, HorizontalAlign } from '@univerjs/core'
import type { ICellData, Nullable } from '@univerjs/core'
import { Documents, SpreadsheetSkeleton, Text } from '@univerjs/engine-render'
import { EditorBridgeService } from '@univerjs/sheets-ui'

const STRONG_RTL_MARK = /[\u200F\u061C]/ // RLM, ALM
const LTR_MARK = '\u200E' // LRM
const RTL_SCRIPT_LETTER =
  /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}]/u
const LETTER = /\p{L}/u

/// First-strong-character scan. Digits, punctuation and spaces are weak or
/// neutral and skipped; RTL-script digits (e.g. Arabic-Indic ٠-٩) are not
/// letters, so they stay weak here too, matching the bidi classes.
export function resolveBidiDirection(text: string): 'ltr' | 'rtl' {
  for (const ch of text) {
    if (STRONG_RTL_MARK.test(ch)) return 'rtl'
    if (ch === LTR_MARK) return 'ltr'
    if (LETTER.test(ch)) return RTL_SCRIPT_LETTER.test(ch) ? 'rtl' : 'ltr'
  }
  return 'ltr'
}

interface CellDocSkeletonLike {
  getViewModel(): {
    getDataModel(): { getBody(): { dataStream?: string } | undefined }
  }
}

interface FontCacheConfigLike {
  horizontalAlign?: HorizontalAlign
  vertexAngle?: number
  centerAngle?: number
  cellData?: Nullable<ICellData>
  documentSkeleton?: CellDocSkeletonLike
}

const isGeneralTextCell = (cell: Nullable<ICellData> | undefined): boolean => {
  const type = cell?.t
  if (type === CellValueType.NUMBER || type === CellValueType.BOOLEAN) return false
  return !(type === undefined && typeof cell?.v === 'number')
}

const cellText = (config: FontCacheConfigLike): string =>
  config.documentSkeleton?.getViewModel().getDataModel().getBody()?.dataStream ??
  String(config.cellData?.v ?? '')

export interface EditorLayoutLike {
  horizontalAlign: HorizontalAlign
  textRotation?: { a?: number; v?: number } | null
  documentModel?: { getBody(): { dataStream?: string } | undefined } | null
}

interface EditCellStateLike {
  documentLayoutObject?: EditorLayoutLike | null
}

/** The editor-bridge counterpart of the General rule: an unrotated,
 *  unaligned cell whose text reads right-to-left edits right-aligned. */
export function rightAlignRtlGeneralEditor(layout: EditorLayoutLike | null | undefined): void {
  if (
    layout &&
    layout.horizontalAlign === HorizontalAlign.UNSPECIFIED &&
    !layout.textRotation?.a &&
    !layout.textRotation?.v &&
    resolveBidiDirection(layout.documentModel?.getBody()?.dataStream ?? '') === 'rtl'
  ) {
    layout.horizontalAlign = HorizontalAlign.RIGHT
  }
}

interface DirectionalContext {
  save(): void
  restore(): void
  direction: CanvasDirection
  textAlign: CanvasTextAlign
}

let installed = false

export function installRtlTextDirectionFix(): void {
  if (installed) return
  installed = true

  // General alignment, single choke point: the font-style cache entry feeds
  // the fast paint path, the overflow-direction calc and the overflow clip,
  // so resolving it here keeps all three consistent.
  const skeletonProto = SpreadsheetSkeleton.prototype as unknown as {
    _calculateOverflowCell(row: number, column: number, config: FontCacheConfigLike): boolean
  }
  const origOverflow = skeletonProto._calculateOverflowCell
  skeletonProto._calculateOverflowCell = function (
    row: number,
    column: number,
    config: FontCacheConfigLike,
  ): boolean {
    if (
      config.horizontalAlign === HorizontalAlign.UNSPECIFIED &&
      !config.vertexAngle &&
      !config.centerAngle &&
      isGeneralTextCell(config.cellData) &&
      resolveBidiDirection(cellText(config)) === 'rtl'
    ) {
      config.horizontalAlign = HorizontalAlign.RIGHT
    }
    return origOverflow.call(this, row, column, config)
  }

  // Rich-text / wrapped cells render through the Documents component, whose
  // page renderConfig carries its own unresolved General alignment.
  const documentsProto = Documents.prototype as unknown as {
    getSkeleton(): CellDocSkeletonLike | null | undefined
    _horizontalHandler(
      pageWidth: number,
      pagePaddingLeft: number,
      pagePaddingRight: number,
      horizontalAlign: HorizontalAlign,
      vertexAngleDeg?: number,
      centerAngleDeg?: number,
      cellValueType?: CellValueType,
    ): number
  }
  const origHorizontal = documentsProto._horizontalHandler
  documentsProto._horizontalHandler = function (
    pageWidth: number,
    pagePaddingLeft: number,
    pagePaddingRight: number,
    horizontalAlign: HorizontalAlign,
    vertexAngleDeg = 0,
    centerAngleDeg = 0,
    cellValueType?: CellValueType,
  ): number {
    if (
      horizontalAlign === HorizontalAlign.UNSPECIFIED &&
      !vertexAngleDeg &&
      !centerAngleDeg &&
      cellValueType !== CellValueType.NUMBER &&
      cellValueType !== CellValueType.BOOLEAN
    ) {
      const text = this.getSkeleton?.()?.getViewModel().getDataModel().getBody()?.dataStream ?? ''
      if (resolveBidiDirection(text) === 'rtl') horizontalAlign = HorizontalAlign.RIGHT
    }
    return origHorizontal.call(
      this,
      pageWidth,
      pagePaddingLeft,
      pagePaddingRight,
      horizontalAlign,
      vertexAngleDeg,
      centerAngleDeg,
      cellValueType,
    )
  }

  // The in-cell editor paints through the same Documents component, so the
  // page handler above right-aligns its text too — but Univer seats the caret
  // (and hit-tests clicks) from the document margin, which the editor bridge
  // only widens for an explicit RIGHT alignment. Resolve General the same way
  // in the bridge state so caret and text agree, and the editor grows away
  // from its right edge like an explicitly right-aligned cell.
  const bridgeProto = EditorBridgeService.prototype as unknown as {
    getLatestEditCellState(): EditCellStateLike | null | undefined
  }
  const origLatestEditCellState = bridgeProto.getLatestEditCellState
  bridgeProto.getLatestEditCellState = function () {
    const state = origLatestEditCellState.call(this)
    rightAlignRtlGeneralEditor(state?.documentLayoutObject)
    return state
  }

  const textClass = Text as unknown as {
    drawWith(ctx: DirectionalContext, props: { text?: unknown }, skeleton?: unknown): void
  }
  const previousDrawWith = textClass.drawWith
  if (typeof previousDrawWith !== 'function') return
  textClass.drawWith = function (
    this: unknown,
    ctx: DirectionalContext,
    props: { text?: unknown },
    skeleton?: unknown,
  ): void {
    if (typeof props?.text !== 'string' || resolveBidiDirection(props.text) !== 'rtl') {
      return previousDrawWith.call(this, ctx, props, skeleton)
    }
    ctx.save()
    ctx.direction = 'rtl'
    ctx.textAlign = 'left'
    try {
      return previousDrawWith.call(this, ctx, props, skeleton)
    } finally {
      ctx.restore()
    }
  }
}
