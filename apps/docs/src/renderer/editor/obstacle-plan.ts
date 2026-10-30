/**
 * Model-side wrap obstacles of an anchor paragraph (X10): the side-wrapped
 * drawings and wrapTopAndBottom bands of a docProtected text-box block, as
 * rectangles in the wrapper's own coordinates. syncWrapObstacles used to
 * read them off the rendered boxes; here they come from the same
 * TextboxDisplay facts the renderer positions the boxes with, so only the
 * anchor paragraph's measured top, the column it sits in and the live height
 * of an auto-fit box stay layout reads. Kept outside BlockPlan until the plan
 * grows obstacle facts.
 */
import type { TextboxDisplay } from '@genoffice/docx-engine'
import type { BlockPlan, PlanDiff, PlanOf } from '../pagination-plan'
import type { BlockBox } from '../pagination-types'
import { resolveAnchor, textboxAnchor, type AnchorCtx } from './anchor-geometry'
import { quarterTurnInsetPx } from './image-rotation'

const CTX: AnchorCtx = { kind: 'body', surface: 'measure' }

/** stylesheet margin of .doc-textbox: an absolutely positioned box lands its margin edge on `top` */
const BOX_MARGIN_TOP_PX = 2

/** text column the wrapper spans (px); the margins resolve the page-edge clamp of a column X */
export interface ColumnBox {
  w: number
  marginLeft: number
  marginRight: number
}

/** one side-wrapped box: layout box relative to the wrapper origin (before rotation) */
export interface ObstacleBox {
  x: number
  y: number
  w: number
  /** fixed height (a:noAutofit); an auto-fit box takes its measured height */
  h?: number
  rotDeg?: number
  /** wp:anchor distL/distR/distT/distB (px) */
  dist: [number, number, number, number]
  /** page-positioned box: not anchor-relative, no obstacle (pinned by syncFloatShifts) */
  pageRelV: boolean
}

export interface ObstacleFacts {
  /** every box of the wrapper that carries data-wrap-side, document order */
  sides: ObstacleBox[]
  /** every drawing of the anchor is a side wrap: the page-turn rule (floatKeepWhole) applies */
  allSides: boolean
  bandKeep: boolean
  /** wrapTopAndBottom band intervals, the data-bands twin (wrapper-relative px) */
  bands: Array<[number, number]>
}

/** wrapTopAndBottom band bottom (px) for a box at the given (live) height */
export function textboxBandBottom(box: TextboxDisplay, height = box.heightPx): number {
  if (box.bandTopPx !== undefined && height !== undefined) {
    // a quarter-turned box turns about its centre: its visual band is the
    // swapped extent, sticking out (w-h)/2 above and below the layout box
    return box.bandTopPx + height + quarterTurnInsetPx(box.widthPx ?? 0, height, box.rotDeg)
  }
  return box.bandBottomPx ?? 0
}

/** wrapTopAndBottom band top (px), lifted by the quarter-turn overhang */
export function textboxBandTop(box: TextboxDisplay, height = box.heightPx): number {
  return (box.bandTopPx ?? 0) - quarterTurnInsetPx(box.widthPx ?? 0, height ?? 0, box.rotDeg)
}

const round1 = (n: number) => Math.round(n * 10) / 10

/** the renderer's clamp (textboxBoxStyle): a column X stays on the paper */
function clampedLeft(box: TextboxDisplay, raw: number, col: ColumnBox | undefined): number {
  if (!box.widthPx || box.pagePinned || box.pageRelX || !col) return raw
  const min = -col.marginLeft
  const max = col.w - box.widthPx + col.marginRight
  return Math.max(min, Math.min(raw, max))
}

export interface Bounds {
  x: number
  y: number
  w: number
  h: number
}

/** axis-aligned bounds of a box turned by rotDeg about its centre (getBoundingClientRect) */
export function turnedBounds(b: Bounds, rotDeg: number | undefined): Bounds {
  if (!rotDeg) return b
  const [c, s] = turnFactors(rotDeg)
  const bw = b.w * c + b.h * s
  const bh = b.w * s + b.h * c
  return { x: b.x + (b.w - bw) / 2, y: b.y + (b.h - bh) / 2, w: bw, h: bh }
}

/** layout height of a box whose rendered (turned) bounds are bw x bh */
export function layoutHeightOf(
  bw: number,
  bh: number,
  w: number,
  rotDeg: number | undefined,
): number {
  if (!rotDeg) return bh
  const [c, s] = turnFactors(rotDeg)
  return c === 0 ? bw : (bh - w * s) / c
}

/** |cos|, |sin| of the turn, exact at quarter turns */
function turnFactors(rotDeg: number): [number, number] {
  const rad = (rotDeg * Math.PI) / 180
  const snap = (v: number) =>
    Math.abs(v) < 1e-9 ? 0 : Math.abs(1 - Math.abs(v)) < 1e-9 ? 1 : Math.abs(v)
  return [snap(Math.cos(rad)), snap(Math.sin(rad))]
}

/** rendered bounds of a side box at `h` (its fixed or measured layout height) */
export const obstacleBounds = (b: ObstacleBox, h: number): Bounds =>
  turnedBounds({ x: b.x, y: b.y, w: b.w, h }, b.rotDeg)

/**
 * The wrap obstacles of a docProtected block's attrs; undefined when the
 * block is not a floating anchor wrapper (or is page-pinned cover art).
 * `heights` are the live layout heights per box (an auto-fit box is as tall
 * as its text).
 */
export function obstacleFactsOf(
  attrs: Record<string, unknown>,
  col?: ColumnBox,
  heights?: ReadonlyArray<number | undefined>,
): ObstacleFacts | undefined {
  const boxes = attrs.textboxes as TextboxDisplay[] | null
  if (!Array.isArray(boxes) || boxes.length === 0) return undefined
  const diagram = attrs.diagramDisplay as { floating?: boolean } | null
  const anyFloating = boxes.some((b) => b.floating) && (!diagram || !!diagram.floating)
  if (!anyFloating || boxes.every((b) => b.pagePinned)) return undefined
  const liveH = (i: number) => (boxes[i].heightPx === undefined ? heights?.[i] : boxes[i].heightPx)
  const sides: ObstacleBox[] = []
  boxes.forEach((box, i) => {
    if (!box.floating || !box.wrapSides) return
    const geom = resolveAnchor(textboxAnchor(box), {}, CTX)
    const d = box.wrapDistPx
    sides.push({
      x: clampedLeft(box, round1(geom.x ?? 0), col),
      y: round1(geom.y ?? 0) + BOX_MARGIN_TOP_PX,
      w: box.widthPx ?? 0,
      ...(liveH(i) !== undefined ? { h: liveH(i) } : {}),
      ...(box.rotDeg ? { rotDeg: box.rotDeg } : {}),
      dist: d ? [d.l, d.r, d.t, d.b] : [0, 0, 0, 0],
      pageRelV: !!box.pageRelV,
    })
  })
  return {
    sides,
    allSides: sides.length === boxes.length,
    bandKeep: boxes.some((b) => b.bandOverflow),
    // the renderer's data-bands at load: parse-time heights (autogrow refreshes both)
    bands: boxes
      .filter((b) => textboxBandBottom(b) > 0)
      .map((b) => [Math.max(0, Math.round(textboxBandTop(b))), Math.round(textboxBandBottom(b))]),
  }
}

/** the wrapper's column box: its own width plus the section margins the clamp reaches into */
export function wrapperColumnBox(el: HTMLElement, widthPx: number): ColumnBox {
  const cs = getComputedStyle(el)
  return {
    w: widthPx,
    marginLeft: parseFloat(cs.getPropertyValue('--doc-margin-left')) || 0,
    marginRight: parseFloat(cs.getPropertyValue('--doc-margin-right')) || 0,
  }
}

export const wrapperBoxes = (el: HTMLElement): HTMLElement[] =>
  Array.from(
    el.querySelectorAll<HTMLElement>(':scope > .doc-textbox, :scope > .box-rotator > .doc-textbox'),
  )

/** live layout heights of the wrapper's boxes (undefined for a collapsed box) */
export function wrapperBoxHeights(
  boxes: readonly HTMLElement[],
  attrs: Record<string, unknown>,
  factor: number,
): Array<number | undefined> {
  const model = (attrs.textboxes as TextboxDisplay[] | null) ?? []
  return boxes.map((el, i) => {
    const r = el.getBoundingClientRect()
    if (r.height <= 0) return undefined
    const m = model[i]
    return layoutHeightOf(r.width / factor, r.height / factor, m?.widthPx ?? 0, m?.rotDeg)
  })
}

/** the model obstacles of a rendered wrapper (undefined without a plan or for a non-anchor block) */
export function wrapperObstacleFacts(
  el: HTMLElement,
  plan: BlockPlan | undefined,
  widthPx: number,
  factor: number,
): ObstacleFacts | undefined {
  if (!plan || plan.kind !== 'protected') return undefined
  const boxes = wrapperBoxes(el)
  if (boxes.length === 0) return undefined
  return obstacleFactsOf(
    plan.node.attrs,
    wrapperColumnBox(el, widthPx),
    wrapperBoxHeights(boxes, plan.node.attrs, factor),
  )
}

/** the data-obs-bottom / data-obs-maxh facts: lowest side-box bottom below the
 *  wrapper top and the tallest box, when the page-turn rule applies */
export function floatKeepWholeOf(f: ObstacleFacts): { bottom: number; maxH: number } | undefined {
  if (!f.allSides || f.bandKeep) return undefined
  let bottom = 0
  let maxH = 0
  for (const b of f.sides) {
    if (b.pageRelV || b.h === undefined || b.h <= 0) continue
    const r = obstacleBounds(b, b.h)
    bottom = Math.max(bottom, r.y + r.h)
    maxH = Math.max(maxH, r.h)
  }
  return bottom > 0 ? { bottom: round1(bottom), maxH: round1(maxH) } : undefined
}

/** one plan-vs-DOM disagreement of the obstacle pass (I9 class `obstacle`) */
export interface ObstacleDiff {
  el: HTMLElement
  cls: string
  plan: unknown
  scrape: unknown
}

/** I9 shadow sink of one pass: the disagreements plus what was compared */
export interface ObstacleShadow {
  diffs: ObstacleDiff[]
  wrappers: number
  boxes: number
}

export const newObstacleShadow = (): ObstacleShadow => ({ diffs: [], wrappers: 0, boxes: 0 })

/** the obstacle shadow as I9 plan diffs, attributed to the measured blocks */
export function obstaclePlanDiffs(
  diffs: readonly ObstacleDiff[],
  blocks: readonly BlockBox[],
  planOf: PlanOf,
): PlanDiff[] {
  return diffs.map((d) => {
    const docxIndex = planOf(d.el)?.docxIndex
    return {
      cls: d.cls,
      block: blocks.findIndex((b) => b.el === d.el),
      ...(docxIndex !== undefined ? { docxIndex } : {}),
      plan: d.plan,
      scrape: d.scrape,
    }
  })
}
