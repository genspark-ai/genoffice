/**
 * Table width geometry: one precedence table for how wide a table box is and
 * where it sits horizontally, shared by the convert-time column fit
 * (resolveTableWidth), the editor's DocTable.renderHTML and the read-only /
 * nested renderTableSpec (tableWidthCss, tablePositionCss).
 *
 * Precedence (probe dates in the rule comments):
 *   fixed  w:tblLayout fixed: the grid is drawn as declared, past the paper if need be
 *   dxa    explicit w:tblW dxa: preferred widths; columns grow to their widest word
 *          and the growth is reclaimed from the others, the declared total is kept
 *          (never shrunk to the column, Word clips at the paper; 2026-09-23)
 *   pct    share of the text column; a saved non-uniform grid is Word's layout,
 *          otherwise tcW, otherwise content distribution (2026-10-20)
 *   auto   tcW / content; over the column it shrinks to column - signed tblInd,
 *          right edge pinned to the margin (2026-09-17); compat < 15 hangs the
 *          box by the side cell margins
 * Horizontal origin: jc left/absent measures w:tblInd from the column's left
 * edge; jc center centres the box on the margin box and jc right pins its
 * right edge to the right margin, both overhanging the margins symmetrically /
 * leftwards when over-wide.
 */
import type { TableModel } from '@genoffice/docx-engine'
import {
  clampTableColWidths,
  expandAutofitColWidths,
  legacyIndentTable,
  type RunDefaults,
} from './convert'

export type TableWidthPolicy = 'fixed' | 'dxa' | 'pct' | 'auto'

export function tableWidthPolicy(model: TableModel): TableWidthPolicy {
  if (model.fixedLayout) return 'fixed'
  if (model.widthPct) return 'pct'
  if (model.dxaWidth) return 'dxa'
  return 'auto'
}

/** schema-neutral width inputs (PM docTable attrs or a display TableModel) */
export interface TableWidthInput {
  /** outer box px (grid plus outer half-borders / spacing); null without an absolute grid */
  widthPx: number | null
  widthPct: number | null
  /** editor display mode; null on the static surface (no autofit classes there) */
  autoFit: 'contents' | 'window' | 'fixed' | null
  /** the model's own mode when the display mode differs (imported grid kept under 'contents') */
  autoFitDeclared?: 'contents' | 'window' | 'fixed'
  fixedLayout: boolean
  dxaWidth: boolean
  align: 'left' | 'center' | 'right' | null
  indentTwips: number
  /** carries w:tblpPr (floated or flowed inline) */
  floatSource: boolean
  /** currently CSS-floated */
  floated: boolean
  /** w:tblpXSpec center on a margin/text anchor: centred like jc center */
  tblpCentered: boolean
  bidiVisual: boolean
}

export interface TableGeometryCtx {
  kind: 'body' | 'nested'
  surface: 'editor' | 'static'
}

export interface TableWidthCss {
  /** CSS width expression (null = auto width) */
  widthExpr: string | null
  styles: string[]
  /** margin-left expression centring a held width on the margin box */
  centerMargin: string | null
  /** margin-left expression pinning a held width's right edge to the right margin */
  rightMargin: string | null
  hold: boolean
  rtlStart: boolean
  origin: 'column' | 'marginBox' | 'rightMargin'
  bleed: 'none' | 'right' | 'left' | 'both'
  log: string[]
}

export const CONTENT_W = 'var(--doc-content-w,100%)'
const PAPER_W = `calc(${CONTENT_W} + var(--doc-margin-left,var(--doc-margin-right,0px)) + var(--doc-margin-right,0px))`

/**
 * Whether the declared width is drawn as is, past the column and the paper:
 * w:tblLayout fixed always; an explicit dxa w:tblW (no autofit mode, no
 * w:tblpPr) whatever the jc (Word clips at the paper and never shrinks: left
 * 2026-09-23, center / right 2026-10-29). Nested tables never hold past their
 * cell. bidiVisual tables are excluded on every surface: no Word probe covers
 * an over-wide right-to-left table yet, so they keep the paper clamp.
 */
export function holdsDeclaredWidthInput(input: TableWidthInput, ctx: TableGeometryCtx): boolean {
  if (ctx.kind !== 'body' || input.widthPx === null || input.widthPct) return false
  return (
    input.fixedLayout ||
    (input.dxaWidth &&
      (input.autoFit === null || (input.autoFitDeclared ?? input.autoFit) === 'fixed') &&
      !input.floatSource &&
      !input.bidiVisual)
  )
}

/** convert-time form of holdsDeclaredWidthInput over the display model (fixed layout excluded: it never clamps) */
export function holdsDeclaredWidth(model: TableModel): boolean {
  return (
    !model.fixedLayout &&
    holdsDeclaredWidthInput(modelWidthInput(model, colSumTwips(model)), {
      kind: 'body',
      surface: 'static',
    })
  )
}

function colSumTwips(model: TableModel): number | null {
  const sum = model.colWidthsTwips?.reduce((a, b) => a + b, 0) ?? 0
  return sum > 0 ? sum : null
}

/** start-aligned bidiVisual table: hangs from the RIGHT margin, w:tblInd mirrored */
export function rtlStartAligned(input: TableWidthInput, ctx: TableGeometryCtx): boolean {
  return (
    input.bidiVisual &&
    ctx.kind === 'body' &&
    !input.floated &&
    input.align !== 'center' &&
    input.align !== 'right'
  )
}

export function tableWidthCss(input: TableWidthInput, ctx: TableGeometryCtx): TableWidthCss {
  const styles: string[] = []
  const log: string[] = []
  const nested = ctx.kind === 'nested'
  const rtlStart = rtlStartAligned(input, ctx)
  const spillMargin = rtlStart ? 'var(--doc-margin-left,0px)' : 'var(--doc-margin-right,0px)'
  const centred = (input.align === 'center' && !input.floated) || input.tblpCentered
  let widthExpr: string | null = null
  let centerMargin: string | null = null
  let rightMargin: string | null = null
  let hold = false
  let origin: TableWidthCss['origin'] = centred
    ? 'marginBox'
    : input.align === 'right'
      ? 'rightMargin'
      : 'column'
  let bleed: TableWidthCss['bleed'] = 'none'
  if (input.autoFit === 'contents') {
    styles.push('width:auto')
    log.push('autofit contents: width auto')
  } else if (input.autoFit === 'window') {
    // 'window' and pct are shares of the section's TEXT COLUMN, not of the
    // canvas padding box: differing-margin sections resolve through --doc-content-w
    widthExpr = CONTENT_W
    styles.push(`width:${widthExpr}`)
    log.push('autofit window: text column')
  } else if (input.widthPct) {
    widthExpr = nested ? `${input.widthPct}%` : `calc(${CONTENT_W} * ${input.widthPct / 100})`
    styles.push(`width:${widthExpr}`)
    log.push(`pct ${input.widthPct}% of the ${nested ? 'cell' : 'text column'}`)
  } else if (input.widthPx) {
    const widthPx = input.widthPx
    // w:tblLayout fixed holds the declared widths even past the paper edge (Word
    // clips there); narrowing to fit would rewrap every column. So does a
    // left-aligned explicit dxa w:tblW (see holdsDeclaredWidth)
    hold = holdsDeclaredWidthInput(input, ctx)
    if (hold) {
      widthExpr = `${widthPx}px`
      styles.push(`width:${widthExpr}`, 'max-width:none')
      // auto margins resolve to 0 on overflow, so a held box wider than the
      // column is pinned by an explicit (negative-capable) margin
      if (centred) centerMargin = `margin-left:calc((${CONTENT_W} - ${widthPx}px)/2)`
      else if (input.align === 'right' && !input.floated)
        rightMargin = `margin-left:calc(${CONTENT_W} - ${widthPx}px)`
      bleed = centred ? 'both' : input.align === 'right' ? 'left' : 'right'
      log.push(`${input.fixedLayout ? 'fixed' : 'dxa'} ${widthPx}px held, no clamp`)
    } else if (!nested && centred) {
      // over-wide grids spill both margins symmetrically (negative-margin
      // centring: auto margins resolve to 0 on overflow and would push the spill
      // right only), never past the paper
      styles.push(`width:min(${widthPx}px,${PAPER_W})`)
      centerMargin = `margin-left:calc((${CONTENT_W} - min(${widthPx}px,${PAPER_W}))/2)`
      bleed = 'both'
      log.push(`centred ${widthPx}px, clamped to the paper`)
    } else {
      // left-aligned grids spill right up to the paper edge; a negative
      // w:tblInd moves the box into the left margin, so the same amount is
      // added back to the right-hand spill allowance
      const indented =
        !input.floated &&
        input.align !== 'center' &&
        input.align !== 'right' &&
        input.indentTwips !== 0
      const indentPx = indented ? input.indentTwips / 15 : 0
      const shift = indentPx < 0 ? `+ ${(-indentPx).toFixed(1)}px` : `- ${indentPx.toFixed(1)}px`
      const base = nested ? '100%' : `calc(${CONTENT_W} + ${spillMargin})`
      const avail = indentPx
        ? nested
          ? indentPx > 0
            ? `calc(100% - ${indentPx.toFixed(1)}px)`
            : base
          : `calc(${CONTENT_W} + ${spillMargin} ${shift})`
        : base
      widthExpr = `min(${widthPx}px,${avail})`
      styles.push(`width:${widthExpr}`)
      bleed = nested ? 'none' : rtlStart ? 'left' : 'right'
      log.push(`${widthPx}px spilling ${bleed}, capped at ${nested ? 'the cell' : 'the paper'}`)
    }
  } else {
    origin = 'column'
    log.push('no absolute grid: auto width')
  }
  return { widthExpr, styles, centerMargin, rightMargin, hold, rtlStart, origin, bleed, log }
}

/**
 * Margins of a non-floating table: jc center (or a centred w:tblpXSpec) on the
 * margin box, jc right on the right margin, bidiVisual start from the right
 * margin, otherwise w:tblInd from the column edge.
 */
export function tablePositionCss(input: TableWidthInput, css: TableWidthCss): string[] {
  if (input.align === 'center' || input.tblpCentered)
    return css.centerMargin ? [css.centerMargin] : ['margin-left:auto', 'margin-right:auto']
  if (input.align === 'right') return [css.rightMargin ?? 'margin-left:auto']
  if (css.rtlStart) return rtlStartMarginCss(css.widthExpr, input.indentTwips / 15)
  if (input.indentTwips) return [`margin-left:${(input.indentTwips / 15).toFixed(1)}px`]
  return []
}

/**
 * Margins of a start-aligned bidiVisual table: Word hangs it from the RIGHT
 * margin and measures w:tblInd from that side (negative indents push it into
 * the margin, over-wide tables spill left). widthExpr null = auto width.
 */
export function rtlStartMarginCss(widthExpr: string | null, indentPx: number): string[] {
  const indent = `${indentPx.toFixed(1)}px`
  if (!widthExpr) return ['margin-left:auto', `margin-right:${indent}`]
  const shift = indentPx < 0 ? `+ ${(-indentPx).toFixed(1)}px` : `- ${indent}`
  return [`margin-left:calc(${CONTENT_W} - ${widthExpr} ${shift})`]
}

/** display TableModel → width inputs for the static surface */
export function modelWidthInput(model: TableModel, widthPx: number | null): TableWidthInput {
  return {
    widthPx,
    widthPct: model.widthPct ?? null,
    autoFit: null,
    fixedLayout: model.fixedLayout === true,
    dxaWidth: model.dxaWidth === true && !model.autoLayout,
    align: model.align ?? null,
    indentTwips: model.indentTwips ?? 0,
    floatSource: !!model.floatSide,
    floated: false,
    tblpCentered: false,
    bidiVisual: model.bidiVisual === true,
  }
}

export interface TableWidthBudgetTwips {
  /** growth bound: paper minus the left margin (left-aligned), the paper (centred) */
  avail: number
  /** the text column */
  fit: number
  paper: number
}

export interface ResolveTableWidthCtx {
  /** compat < 15: w:tblInd measured to the cell text, box hangs by the cell margins */
  legacyIndent: boolean
  defaults?: RunDefaults
}

export interface ResolvedTableWidth {
  /** display model with the fitted grid */
  model: TableModel
  policy: TableWidthPolicy
  /** Σ grid twips; null without an absolute grid (pct / auto) */
  widthTwips: number | null
  colTwips: number[] | null
  widthPct: number | null
  origin: 'column' | 'marginBox' | 'rightMargin'
  /** signed w:tblInd from the origin (twips, left-aligned tables only) */
  x: number
  /** compat < 15 cell-margin hang added to the box (twips) */
  hang: number
  /** declared width held past the column / paper */
  hold: boolean
  bleed: 'none' | 'right' | 'left' | 'both'
  log: string[]
}

/**
 * Convert-time width resolution: legacy indent, autofit growth / reclaim
 * within the section budget, then the paper clamp unless the declared width
 * is held. The returned model is what the PM attrs and the static renderer
 * draw; the geometry fields are the decision log for tests and probes.
 */
export function resolveTableWidth(
  model: TableModel,
  budget: TableWidthBudgetTwips | null,
  ctx: ResolveTableWidthCtx,
): ResolvedTableWidth {
  const log: string[] = []
  const policy = tableWidthPolicy(model)
  let t = model
  if (ctx.legacyIndent) {
    t = legacyIndentTable(t)
    if (t !== model) log.push(`legacy indent ${model.indentTwips ?? 0} -> ${t.indentTwips}`)
  }
  const hang = ctx.legacyIndent && t !== model ? (model.indentTwips ?? 0) - (t.indentTwips ?? 0) : 0
  let hold = false
  if (budget) {
    const avail =
      model.align === 'center' || model.floatPos?.xSpec === 'center' ? budget.paper : budget.avail
    const before = t.colWidthsTwips
    t = expandAutofitColWidths(
      t,
      avail,
      budget.fit,
      undefined,
      false,
      ctx.legacyIndent,
      ctx.defaults,
    )
    if (t.colWidthsTwips !== before) log.push(`${policy} autofit within ${avail}/${budget.fit}`)
    hold = holdsDeclaredWidth(t)
    if (hold) log.push('declared dxa width held, no paper clamp')
    else {
      const clamped = clampTableColWidths(t, budget.paper)
      if (clamped !== t) log.push(`clamped to the paper ${budget.paper}`)
      t = clamped
    }
  }
  const colTwips = t.colWidthsTwips ?? null
  const widthTwips = colTwips ? colTwips.reduce((a, b) => a + b, 0) : null
  const origin =
    t.align === 'center' || t.floatPos?.xSpec === 'center'
      ? 'marginBox'
      : t.align === 'right'
        ? 'rightMargin'
        : 'column'
  const bleed: ResolvedTableWidth['bleed'] =
    widthTwips != null &&
    budget &&
    widthTwips > budget.fit - (origin === 'column' ? (t.indentTwips ?? 0) : 0)
      ? origin === 'marginBox'
        ? 'both'
        : origin === 'rightMargin' || t.bidiVisual
          ? 'left'
          : 'right'
      : 'none'
  return {
    model: t,
    policy,
    widthTwips,
    colTwips,
    widthPct: t.widthPct ?? null,
    origin,
    x: origin === 'column' ? (t.indentTwips ?? 0) : 0,
    hang,
    hold,
    bleed,
    log,
  }
}
