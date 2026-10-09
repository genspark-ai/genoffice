import type { Editor } from '@tiptap/core'
import type { ImageWrap, TextboxDisplay } from '@genoffice/docx-engine'

const EMU_PER_PX = 9525

// ---- shape box-model mapping (pure, unit-tested) ----

/**
 * imageWrap mirror for a shape the user put back in line: the attr's null default
 * also means "untouched", and a freshly inserted shape's anchor must still save.
 */
export const SHAPE_INLINE_WRAP = 'inline'

/** The wrap mode a shape's box represents, as the Wrap Text dropdown spells it. */
export function shapeWrapOf(box: TextboxDisplay): ImageWrap | null {
  if (box.behind) return 'behind'
  if (box.wrapSide === 'left') return 'square-left'
  if (box.wrapSide === 'right') return 'square-right'
  if (box.bandBottomPx != null) return 'topBottom'
  if (box.floating) return 'front'
  return null
}

/** Wrap is the only thing this edits: geometry, position, text and fill stay untouched. */
function withWrapFields(box: TextboxDisplay, wrap: ImageWrap | null): TextboxDisplay {
  const next: TextboxDisplay = { ...box }
  delete next.floating
  delete next.behind
  delete next.noWrap
  delete next.wrapSides
  delete next.wrapSide
  delete next.wrapEdgePx
  delete next.wrapGapPx
  delete next.bandTopPx
  delete next.bandBottomPx
  delete next.bandOverflow
  delete next.bandBeside
  if (wrap === 'front') return { ...next, floating: true, noWrap: true }
  if (wrap === 'behind') return { ...next, floating: true, behind: true, noWrap: true }
  if (wrap === 'square-left' || wrap === 'square-right') {
    return {
      ...next,
      wrapSides: true,
      wrapSide: wrap === 'square-left' ? 'left' : 'right',
      wrapEdgePx: box.wrapEdgePx ?? 0,
      wrapGapPx: box.wrapGapPx ?? 12,
    }
  }
  if (wrap === 'topBottom') {
    const top = box.bandTopPx ?? Math.round((box.offsetYEmu ?? 0) / EMU_PER_PX)
    const bottom = box.bandBottomPx ?? top + (box.heightPx ?? 0)
    return { ...next, floating: true, bandTopPx: top, bandBottomPx: bottom }
  }
  return next
}

/** Apply a rank to a shape box; an in-flow box floats in front first so the rank is visible (Word parity). */
export function boxWithZ(box: TextboxDisplay, z: number): TextboxDisplay {
  const next = { ...box, z }
  // a banded box already floats; a side-wrapped one keeps its CSS float, so
  // its rank only paints once the box is front/behind (Word keeps the wrap)
  if (box.wrapSides || box.bandBottomPx != null) return next
  if (!box.floating) return { ...next, floating: true, behind: false, noWrap: true }
  return next
}

// ---- commands ----

function selectedShapeBoxes(editor: Editor): TextboxDisplay[] | null {
  const boxes = editor.getAttributes('docProtected').textboxes as TextboxDisplay[] | null
  return Array.isArray(boxes) && boxes.length > 0 ? boxes : null
}

/** Write the first box back, mirroring the mode on the node attrs (dropdown + save-path compatibility). */
function writeFirstBox(editor: Editor, box: TextboxDisplay, mirror: Record<string, unknown>): void {
  const boxes = selectedShapeBoxes(editor)!
  editor
    .chain()
    .focus()
    .updateAttributes('docProtected', { ...mirror, textboxes: [box, ...boxes.slice(1)] })
    .run()
}

/** tight/through have no box representation: the nearest renderable mode keeps text beside the shape. */
function normalizedShapeWrap(wrap: string | null): ImageWrap | null | undefined {
  if (
    wrap === null ||
    wrap === 'front' ||
    wrap === 'behind' ||
    wrap === 'topBottom' ||
    wrap === 'square-left' ||
    wrap === 'square-right'
  )
    return wrap
  if (wrap === 'tight-left' || wrap === 'through-left') return 'square-left'
  if (wrap === 'tight-right' || wrap === 'through-right') return 'square-right'
  return undefined
}

/** Wrap Text on the selection: the shape's box model for a shape, the anchor attrs for an image. */
export function setFloatingWrap(editor: Editor, wrap: string | null): void {
  const boxes = selectedShapeBoxes(editor)
  if (boxes) {
    const mode = normalizedShapeWrap(wrap)
    if (mode === undefined) return
    writeFirstBox(editor, withWrapFields(boxes[0], mode), {
      imageWrap: mode ?? SHAPE_INLINE_WRAP,
    })
    return
  }
  // an inline image drops its floating-position attrs (the existing image behaviour)
  const cleared =
    wrap === null
      ? { imagePosH: null, imagePosV: null, imageOffsetXEmu: null, imageOffsetYEmu: null }
      : {}
  editor
    .chain()
    .focus()
    .updateAttributes('docProtected', { imageWrap: wrap, ...cleared })
    .run()
}

/** Rank of the selection: the shape's box rank, or imageZOrder for an image anchor. */
function currentZOrder(editor: Editor): number {
  const boxes = selectedShapeBoxes(editor)
  if (boxes) return Number(boxes[0].z ?? 0)
  return Number(editor.getAttributes('docProtected').imageZOrder ?? 0)
}

/** Ranks of every anchor in the document plus the selection's own; Word's to-front/to-back are document-global. */
function documentZOrders(editor: Editor, rank: number): number[] {
  const zs: number[] = [rank]
  editor.state.doc.descendants((n) => {
    if (n.type.name !== 'docProtected') return
    const boxes = n.attrs.textboxes as TextboxDisplay[] | undefined
    if (Array.isArray(boxes) && boxes.length > 0) {
      const box = boxes[0]
      if (box.floating || box.wrapSides || box.bandBottomPx != null) zs.push(Number(box.z ?? 0))
      return
    }
    if (n.attrs.imageWrap === 'front' || n.attrs.imageWrap === 'behind')
      zs.push(Number(n.attrs.imageZOrder ?? 0))
  })
  return zs
}

function setZOrder(editor: Editor, z: number): void {
  const boxes = selectedShapeBoxes(editor)
  if (boxes) {
    const next = boxWithZ(boxes[0], z)
    writeFirstBox(editor, next, { imageWrap: shapeWrapOf(next), imageZOrder: z })
    return
  }
  const attrs: Record<string, unknown> = { imageZOrder: z }
  const wrap = editor.getAttributes('docProtected').imageWrap as string | null | undefined
  if (wrap !== 'front' && wrap !== 'behind') attrs.imageWrap = 'front'
  editor.chain().focus().updateAttributes('docProtected', attrs).run()
}

export function bringToFront(editor: Editor): void {
  setZOrder(editor, Math.max(...documentZOrders(editor, currentZOrder(editor))) + 1)
}

export function bringForward(editor: Editor): void {
  setZOrder(editor, currentZOrder(editor) + 1)
}

export function sendBackward(editor: Editor): void {
  setZOrder(editor, currentZOrder(editor) - 1)
}

export function sendToBack(editor: Editor): void {
  setZOrder(editor, Math.min(...documentZOrders(editor, currentZOrder(editor))) - 1)
}
