import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { imageDropPages } from './extensions'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { SHAPE_INLINE_WRAP, shapeWrapOf } from './floating-z-order'
import {
  selectedObjectPositions,
  setObjectSelection,
  type OffsetFrame,
} from './object-multi-select'

const EMU_PER_PX = 9525

export const ALIGN_MODES = [
  'left',
  'center',
  'right',
  'top',
  'middle',
  'bottom',
  'distribute-h',
  'distribute-v',
] as const
export type AlignMode = (typeof ALIGN_MODES)[number]

/** what the objects line up against: the paper, its text area, or each other */
export const ALIGN_FRAMES = ['page', 'margin', 'objects'] as const
export type AlignFrame = (typeof ALIGN_FRAMES)[number]

export interface Box {
  left: number
  top: number
  width: number
  height: number
}

// ---- geometry (pure, unit-tested) ----

export function boundsOf(boxes: Box[]): Box {
  const left = Math.min(...boxes.map((b) => b.left))
  const top = Math.min(...boxes.map((b) => b.top))
  const right = Math.max(...boxes.map((b) => b.left + b.width))
  const bottom = Math.max(...boxes.map((b) => b.top + b.height))
  return { left, top, width: right - left, height: bottom - top }
}

/** Equal gaps along one axis between the frame's two edges, in the objects' current order. */
function distributed(starts: number[], sizes: number[], from: number, span: number): number[] {
  if (starts.length === 1) return [from + (span - sizes[0]) / 2]
  const order = starts.map((_, i) => i).sort((a, b) => starts[a] - starts[b] || a - b)
  const gap = (span - sizes.reduce((sum, size) => sum + size, 0)) / (starts.length - 1)
  const out = [...starts]
  let at = from
  for (const i of order) {
    out[i] = at
    at += sizes[i] + gap
  }
  return out
}

/** Where each box goes; a box keeps the coordinate its mode does not touch. */
export function alignedBoxes(boxes: Box[], mode: AlignMode, frame: Box): Box[] {
  if (mode === 'distribute-h' || mode === 'distribute-v') {
    const horizontal = mode === 'distribute-h'
    const starts = distributed(
      boxes.map((b) => (horizontal ? b.left : b.top)),
      boxes.map((b) => (horizontal ? b.width : b.height)),
      horizontal ? frame.left : frame.top,
      horizontal ? frame.width : frame.height,
    )
    return boxes.map((b, i) => (horizontal ? { ...b, left: starts[i] } : { ...b, top: starts[i] }))
  }
  return boxes.map((b) => {
    switch (mode) {
      case 'left':
        return { ...b, left: frame.left }
      case 'center':
        return { ...b, left: frame.left + (frame.width - b.width) / 2 }
      case 'right':
        return { ...b, left: frame.left + frame.width - b.width }
      case 'top':
        return { ...b, top: frame.top }
      case 'middle':
        return { ...b, top: frame.top + (frame.height - b.height) / 2 }
      default:
        return { ...b, top: frame.top + frame.height - b.height }
    }
  })
}

/** Objects line up against each other only when there are several; one alone uses the page. */
export function effectiveAlignFrame(frame: AlignFrame, count: number): AlignFrame {
  return frame === 'objects' && count < 2 ? 'page' : frame
}

// ---- measuring ----

export function pageZoom(): number {
  const el = document.querySelector('.doc-zoom')
  return el ? parseFloat(getComputedStyle(el).zoom || '1') || 1 : 1
}

/** the object as it shows, client px; null when it is not laid out */
export function objectBox(view: EditorView, pos: number): Box | null {
  const node = view.state.doc.nodeAt(pos)
  const dom = view.nodeDOM(pos)
  if (!node || !(dom instanceof HTMLElement)) return null
  let parts: Element[]
  if (node.type.name === 'docInlineImage') {
    parts = [dom instanceof HTMLImageElement ? dom : (dom.querySelector('img') ?? dom)]
  } else if (node.attrs.blockType === 'image') {
    parts = [dom.querySelector('.doc-img-wrap') ?? dom]
  } else {
    parts = Array.from(dom.querySelectorAll('.doc-textbox'))
  }
  const rects = parts
    .map((el) => el.getBoundingClientRect())
    .filter((r) => r.width > 0 && r.height > 0)
  if (rects.length === 0) return null
  return boundsOf(
    rects.map((r) => ({ left: r.left, top: r.top, width: r.width, height: r.height })),
  )
}

/** The paper (or its text area) around `box`, client px: the page the object's middle is on. */
function pageFrame(view: EditorView, box: Box, margin: boolean, zoom: number): Box | null {
  const paper = view.dom.closest<HTMLElement>('.doc-page')
  if (!paper) return null
  const paperRect = paper.getBoundingClientRect()
  const style = getComputedStyle(paper)
  const pad = (v: string) => (parseFloat(v) || 0) * zoom
  const pages = imageDropPages(view.dom, zoom)
  const mid = box.top + box.height / 2
  const page =
    pages.find((p) => mid < p.paperBottom) ?? (pages.length > 0 ? pages[pages.length - 1] : null)
  if (!page) return null
  const top = Math.max(page.paperTop, paperRect.top)
  if (!margin)
    return { left: paperRect.left, top, width: paperRect.width, height: page.paperBottom - top }
  const left = paperRect.left + pad(style.paddingLeft)
  const right = paperRect.right - pad(style.paddingRight)
  return { left, top: page.bodyTop, width: right - left, height: page.bodyBottom - page.bodyTop }
}

// ---- moving ----

/** Text wraps on the open side of a block picture: the side re-picked from where it now stands, like a drag. */
function wrapOnSide(wrap: unknown, side: 'left' | 'right' | null): unknown {
  const kind = side ? /^(square|tight|through)-(?:left|right)$/.exec(String(wrap ?? ''))?.[1] : null
  return kind ? `${kind}-${side}` : wrap
}

/** The object's attrs after a move of (dx, dy) document px that leaves it on `side` of the text area. */
function movedAttrs(
  node: PmNode,
  dx: number,
  dy: number,
  side: 'left' | 'right' | null,
): Record<string, unknown> {
  const ex = Math.round(dx * EMU_PER_PX)
  const ey = Math.round(dy * EMU_PER_PX)
  const attrs = node.attrs
  if (node.type.name === 'docInlineImage') {
    // a picture in a run of text keeps its side: its offset measures from the same edge either way
    return {
      ...attrs,
      offsetXEmu: Number(attrs.offsetXEmu ?? 0) + ex,
      offsetYEmu: Number(attrs.offsetYEmu ?? 0) + ey,
      lineCenterV: false,
    }
  }
  const boxes = attrs.textboxes as TextboxDisplay[] | null
  // a shape placed by alignment has no offsets on its node: the ones it is
  // drawn at (resolved when the file was read) are where it is, and what a
  // move starts from
  const shown = Array.isArray(boxes) ? boxes[0] : undefined
  const next: Record<string, unknown> = {
    ...attrs,
    imageWrap: wrapOnSide(attrs.imageWrap, side),
    imageOffsetXEmu: Number(attrs.imageOffsetXEmu ?? shown?.offsetXEmu ?? 0) + ex,
    imageOffsetYEmu: Number(attrs.imageOffsetYEmu ?? shown?.offsetYEmu ?? 0) + ey,
    imagePosH: null,
    imagePosV: null,
  }
  if (!Array.isArray(boxes) || boxes.length === 0) return next
  // a shape shows at its box's own offsets; the node's are what its anchor saves
  next.textboxes = boxes.map((box) => ({
    ...box,
    offsetXEmu: (box.offsetXEmu ?? 0) + ex,
    offsetYEmu: (box.offsetYEmu ?? 0) + ey,
  }))
  // a moved shape saves through its wrap mirror, like a dragged one
  if (!attrs.imageWrap || attrs.imageWrap === SHAPE_INLINE_WRAP)
    next.imageWrap = shapeWrapOf(boxes[0]) ?? 'square-left'
  return next
}

interface Shift {
  dx: number
  dy: number
}

/** how far each object still is from its target, document px; objects already there are left out */
function distances(view: EditorView, targets: Map<number, Box>, zoom: number): Map<number, Shift> {
  const out = new Map<number, Shift>()
  for (const [pos, target] of targets) {
    const box = objectBox(view, pos)
    if (!box) continue
    const dx = (target.left - box.left) / zoom
    const dy = (target.top - box.top) / zoom
    if (Math.abs(dx) >= 0.5 || Math.abs(dy) >= 0.5) out.set(pos, { dx, dy })
  }
  return out
}

function shiftObjects(
  view: EditorView,
  shifts: Map<number, Shift>,
  targets: Map<number, Box>,
  zoom: number,
  follows?: Transaction,
): Transaction | undefined {
  const { state } = view
  const tr = state.tr
  // a pass that follows from the move joins its undo step
  if (follows) tr.setMeta('appendedTransaction', follows)
  for (const [pos, { dx, dy }] of shifts) {
    const node = state.doc.nodeAt(pos)
    const target = targets.get(pos)
    if (!node || !target) continue
    const area = pageFrame(view, target, true, zoom)
    const side = !area
      ? null
      : target.left + target.width / 2 > area.left + area.width / 2
        ? 'right'
        : 'left'
    tr.setNodeMarkup(pos, undefined, movedAttrs(node, dx, dy, side))
  }
  if (!tr.docChanged) return follows
  view.dispatch(setObjectSelection(tr, selectedObjectPositions(state)))
  return tr
}

const reach = (shift: Shift): number => Math.max(Math.abs(shift.dx), Math.abs(shift.dy))

/**
 * Moves each object to its target (client px). The offsets measure from
 * anchors whose own place can shift with the move (a preset position giving
 * way to an offset, the wrap side changing), so a second pass folds what the
 * first one could not predict into the offsets. A correction that does not
 * land the object (side floats cannot overlap on the canvas, so one is held
 * off its target) is taken back: the saved offset stays the one asked for.
 * The passes join the move's undo step, and the move joins the step of the
 * edit it `follows` (the others of a drag follow the dragged object).
 */
export function moveObjectsTo(
  view: EditorView,
  targets: Map<number, Box>,
  follows?: Transaction,
): boolean {
  const zoom = pageZoom()
  const first = distances(view, targets, zoom)
  if (first.size === 0) return false
  const move = shiftObjects(view, first, targets, zoom, follows)
  // an axis the canvas pins to the page again after every change shows at a
  // passing place until then: along it the shape moved by what it was shifted,
  // and there is nothing to correct. The other axis settles like any object's.
  const settled = (): Map<number, Shift> => {
    const out = new Map<number, Shift>()
    for (const [pos, shift] of distances(view, targets, zoom)) {
      const pinned = repinnedAxes(view.state.doc.nodeAt(pos))
      const left = { dx: pinned.x ? 0 : shift.dx, dy: pinned.y ? 0 : shift.dy }
      if (reach(left) >= 0.5) out.set(pos, left)
    }
    return out
  }
  const correction = settled()
  if (correction.size > 0) {
    shiftObjects(view, correction, targets, zoom, move)
    const left = settled()
    const back = new Map<number, Shift>()
    for (const [pos, asked] of correction) {
      const still = left.get(pos)
      if (still && reach(still) >= 1) back.set(pos, { dx: -asked.dx, dy: -asked.dy })
    }
    shiftObjects(view, back, targets, zoom, move)
  }
  writePaperOffsets(view, targets, move)
  return true
}

/** the axes of a shape that the canvas pins to its page again after each render */
function repinnedAxes(node: PmNode | null): { x: boolean; y: boolean } {
  const shown = (node?.attrs.textboxes as TextboxDisplay[] | null | undefined)?.[0]
  return { x: !!shown?.pageRelX, y: !!shown?.pageRelV }
}

/** the frame an axis is saved in when the file places it against the paper: null when it flows with the text */
function paperFrame(rel: unknown): 'page' | 'margin' | null {
  if (
    rel == null ||
    rel === 'column' ||
    rel === 'paragraph' ||
    rel === 'character' ||
    rel === 'line'
  )
    return null
  // the other frames (an inner or outer margin, the top or bottom one) have no writer here: the page says it exactly
  return rel === 'margin' ? 'margin' : 'page'
}

/**
 * Shapes the file places against the page or its margins get the offsets
 * they are saved with from where they stand on the paper: `places` is where
 * each one is (it was put there, or it never moved). The offsets a shape is
 * drawn at come in whatever frame the reader resolved (an alignment, a
 * percentage, an offset, from the page or from the body), so they cannot be
 * trusted to be the frame the file names; the paper can.
 */
function writePaperOffsets(
  view: EditorView,
  places: Map<number, Box>,
  follows?: Transaction,
): void {
  const zoom = pageZoom()
  const { state } = view
  const tr = state.tr
  if (follows) tr.setMeta('appendedTransaction', follows)
  for (const [pos, box] of places) {
    const node = state.doc.nodeAt(pos)
    const shown = (node?.attrs.textboxes as TextboxDisplay[] | null)?.[0]
    if (!node || node.type.name !== 'docProtected' || !shown) continue
    const h = paperFrame(node.attrs.imageRelH)
    const v = paperFrame(node.attrs.imageRelV)
    if (!h && !v) continue
    const paper = pageFrame(view, box, false, zoom)
    const body = pageFrame(view, box, true, zoom)
    if (!paper || !body) continue
    // from the middle: a turned shape's corner is not the corner of what shows
    const width = shown.widthPx ?? box.width / zoom
    const height = shown.heightPx ?? box.height / zoom
    const x = (box.left + box.width / 2 - (h === 'margin' ? body : paper).left) / zoom - width / 2
    const y = (box.top + box.height / 2 - (v === 'margin' ? body : paper).top) / zoom - height / 2
    const frame: OffsetFrame = { h: h ?? 'column', v: v ?? 'paragraph' }
    tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      ...(h ? { imageOffsetXEmu: Math.round(x * EMU_PER_PX) } : {}),
      ...(v ? { imageOffsetYEmu: Math.round(y * EMU_PER_PX) } : {}),
      imageOffsetFrame: frame,
    })
  }
  if (tr.docChanged) view.dispatch(setObjectSelection(tr, selectedObjectPositions(state)))
}

/** the object says where it is with offsets of its own, not a preset place or an inherited one */
function hasOwnOffsets(node: PmNode): boolean {
  const attrs = node.attrs
  if (node.type.name === 'docInlineImage')
    return attrs.offsetXEmu != null && attrs.offsetYEmu != null
  return (
    attrs.imageOffsetXEmu != null &&
    attrs.imageOffsetYEmu != null &&
    !attrs.imagePosH &&
    !attrs.imagePosV
  )
}

/**
 * Gives every object offsets of its own without moving it on the page. A
 * group is saved by shifting its first member's offsets, which a member
 * placed by alignment (a freshly inserted shape is centred) does not have.
 * Returns the edit, for what follows to join its undo step.
 */
export function pinObjectOffsets(view: EditorView, positions: number[]): Transaction | undefined {
  const zoom = pageZoom()
  const targets = new Map<number, Box>()
  for (const pos of positions) {
    const node = view.state.doc.nodeAt(pos)
    const box = objectBox(view, pos)
    if (node && box && !hasOwnOffsets(node)) targets.set(pos, box)
  }
  if (targets.size === 0) return undefined
  const still = new Map(Array.from(targets.keys(), (pos) => [pos, { dx: 0, dy: 0 }]))
  const pinned = shiftObjects(view, still, targets, zoom)
  // a shape's offsets started from where it is drawn, so it did not move. A
  // picture has no such offsets: its own started from nothing, and the passes
  // bring it back to where it was
  const pictures = new Map(
    Array.from(targets).filter(([pos]) => !view.state.doc.nodeAt(pos)?.attrs.textboxes),
  )
  if (pictures.size > 0) moveObjectsTo(view, pictures, pinned)
  writePaperOffsets(view, targets, pinned)
  return pinned
}

/** Align or distribute the selected floating objects. False when nothing is selected or laid out. */
export function alignSelectedObjects(editor: Editor, mode: AlignMode, frame: AlignFrame): boolean {
  const { view } = editor
  const positions = selectedObjectPositions(view.state)
  const measured = positions.flatMap((pos) => {
    const box = objectBox(view, pos)
    return box ? [{ pos, box }] : []
  })
  if (measured.length === 0) return false
  const boxes = measured.map((m) => m.box)
  const zoom = pageZoom()
  const against = effectiveAlignFrame(frame, measured.length)
  const targets = new Map<number, Box>()
  if (against === 'objects') {
    alignedBoxes(boxes, mode, boundsOf(boxes)).forEach((box, i) =>
      targets.set(measured[i].pos, box),
    )
  } else if (mode === 'distribute-h' || mode === 'distribute-v') {
    // spread over the page the first object is on
    const page = pageFrame(view, boxes[0], against === 'margin', zoom)
    if (!page) return false
    alignedBoxes(boxes, mode, page).forEach((box, i) => targets.set(measured[i].pos, box))
  } else {
    for (const { pos, box } of measured) {
      const page = pageFrame(view, box, against === 'margin', zoom)
      if (page) targets.set(pos, alignedBoxes([box], mode, page)[0])
    }
  }
  const moved = moveObjectsTo(view, targets)
  view.focus()
  return moved
}
