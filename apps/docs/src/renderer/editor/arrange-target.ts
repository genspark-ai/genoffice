import type { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseLazyMediaUrl, type TextboxDisplay } from '@genoffice/docx-engine'
import { SHAPE_INLINE_WRAP, shapeWrapOf } from './floating-z-order'
import { selectedBoxIndex } from './selected-box'

export type ArrangeKind = 'blockPicture' | 'runPicture' | 'pictureFrame' | 'shape'

/** what the Arrange group can do with the selected object */
export interface ArrangeState {
  kind: ArrangeKind | null
  /** the wrap as Wrap Text spells it; null is in line with text */
  wrap: string | null
  canWrap: boolean
  /** a preset place is a place on the page: a picture block has one, a picture among text takes one by becoming a block */
  canPosition: boolean
  /** stacking ranks are kept by floating pictures and by shapes */
  canOrder: boolean
  canRotate: boolean
  canFlip: boolean
  posH: string | null
  posV: string | null
}

const NOTHING: ArrangeState = {
  kind: null,
  wrap: null,
  canWrap: false,
  canPosition: false,
  canOrder: false,
  canRotate: false,
  canFlip: false,
  posH: null,
  posV: null,
}

/** a header, footer or text box holds its own pictures: the editor whose object is selected */
export function arrangeEditor(editor: Editor, sub: Editor | null): Editor {
  return sub && sub.state.selection instanceof NodeSelection ? sub : editor
}

function runPictureOf(editor: Editor): { pos: number; attrs: Record<string, unknown> } | null {
  const sel = editor.state.selection
  if (!(sel instanceof NodeSelection) || sel.node.type.name !== 'docInlineImage') return null
  const { attrs } = sel.node
  return attrs.rule || !attrs.dataUrl ? null : { pos: sel.from, attrs }
}

const NEW_ANCHOR_DIST_EMU = 114300
const RASTER_SRC = /^data:(image\/(?:png|jpeg|gif));base64,(.+)$/
// what a new picture block does not write: a crop, a link, an outline
const BLOCK_LOSES = /<a:srcRect\s[^>]*\b[ltrb]="-?[1-9]|<a:hlinkClick\b/

/** a picture among the text of a body paragraph that can be written again as a picture block */
function movesToBlock(editor: Editor, run: { pos: number; attrs: Record<string, unknown> }) {
  const { doc } = editor.state
  const $pos = doc.resolve(run.pos)
  const blockType = editor.schema.nodes.docProtected
  if ($pos.depth !== 1 || !blockType) return false
  if (!doc.canReplaceWith($pos.index(0), $pos.index(0), blockType)) return false
  const src = String(run.attrs.dataUrl)
  if (!RASTER_SRC.test(src) && !parseLazyMediaUrl(src)) return false
  if (!Number(run.attrs.widthPx) || !Number(run.attrs.heightPx)) return false
  return !run.attrs.layers && !run.attrs.border && !BLOCK_LOSES.test(String(run.attrs.xml ?? ''))
}

function selectedBox(editor: Editor): { boxes: TextboxDisplay[]; index: number } | null {
  const boxes = editor.getAttributes('docProtected').textboxes as TextboxDisplay[] | null
  if (!Array.isArray(boxes) || boxes.length === 0) return null
  const index = selectedBoxIndex(editor.state)
  return boxes[index] ? { boxes, index } : null
}

export function arrangeStateOf(editor: Editor, sub: Editor | null): ArrangeState {
  const target = arrangeEditor(editor, sub)
  const run = runPictureOf(target)
  if (run) {
    // a picture drawn from several of a group has no one a:xfrm to turn
    const single = !run.attrs.layers
    return {
      ...NOTHING,
      kind: 'runPicture',
      wrap: (run.attrs.wrap as string | null) ?? null,
      canWrap: true,
      canPosition: target === editor && movesToBlock(target, run),
      canOrder: !!run.attrs.wrap,
      canRotate: single,
      canFlip: single,
    }
  }
  const picked = selectedBox(target)
  if (picked) {
    const box = picked.boxes[picked.index]
    // a frame's turn is saved through its own drawing: one of a group, or of a header, has none to address
    const canRotate = !box.picture || (target === editor && box.drawingIndex !== undefined)
    return {
      ...NOTHING,
      kind: box.picture ? 'pictureFrame' : 'shape',
      wrap: shapeWrapOf(box),
      canWrap: true,
      canOrder: true,
      canRotate,
      // a member's flip is saved along its group's axes, which a turned group does not share with the page
      canFlip: canRotate && !box.grouped,
    }
  }
  const attrs = target.getAttributes('docProtected')
  if (attrs.blockType !== 'image' || !attrs.imageDataUrl) return NOTHING
  const wrap = (attrs.imageWrap as string | null) ?? null
  return {
    kind: 'blockPicture',
    wrap,
    canWrap: true,
    canPosition: true,
    canOrder: wrap !== null && wrap !== SHAPE_INLINE_WRAP,
    canRotate: true,
    canFlip: true,
    posH: (attrs.imagePosH as string | null) ?? null,
    posV: (attrs.imagePosV as string | null) ?? null,
  }
}

const turned = (from: unknown, delta: number): number =>
  ((Math.round((Number(from) || 0) + delta) % 360) + 360) % 360

function setRunPicture(editor: Editor, pos: number, attrs: Record<string, unknown>): void {
  const { state, view } = editor
  const tr = state.tr.setNodeMarkup(pos, undefined, attrs)
  view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)))
  view.focus()
}

/** Rotate Right / Left 90°: the selected picture or shape turns by `delta` degrees clockwise. */
export function rotateSelectedObject(editor: Editor, sub: Editor | null, delta: number): boolean {
  if (!Number.isFinite(delta) || !arrangeStateOf(editor, sub).canRotate) return false
  const target = arrangeEditor(editor, sub)
  const run = runPictureOf(target)
  if (run) {
    setRunPicture(target, run.pos, {
      ...run.attrs,
      rotDeg: turned(run.attrs.rotDeg, delta) || null,
    })
    return true
  }
  const picked = selectedBox(target)
  if (picked) {
    const box = { ...picked.boxes[picked.index] }
    const rotDeg = turned(box.rotDeg, delta)
    // no turn is no attribute, so the file keeps a clean shape
    if (rotDeg) box.rotDeg = rotDeg
    else delete box.rotDeg
    const textboxes = picked.boxes.map((b, i) => (i === picked.index ? box : b))
    return target.chain().focus().updateAttributes('docProtected', { textboxes }).run()
  }
  const now = target.getAttributes('docProtected').imageRotDeg
  return target
    .chain()
    .focus()
    .updateAttributes('docProtected', { imageRotDeg: turned(now, delta) || null })
    .run()
}

/** Flip Horizontal / Vertical of the selected picture or shape. */
export function flipSelectedObject(editor: Editor, sub: Editor | null, axis: 'h' | 'v'): boolean {
  if (!arrangeStateOf(editor, sub).canFlip) return false
  const target = arrangeEditor(editor, sub)
  const run = runPictureOf(target)
  if (run) {
    const key = axis === 'h' ? 'flipH' : 'flipV'
    setRunPicture(target, run.pos, { ...run.attrs, [key]: !run.attrs[key] })
    return true
  }
  const picked = selectedBox(target)
  if (picked) {
    const key = axis === 'h' ? 'flipH' : 'flipV'
    const box = { ...picked.boxes[picked.index] }
    if (box[key]) delete box[key]
    else box[key] = true
    const textboxes = picked.boxes.map((b, i) => (i === picked.index ? box : b))
    return target.chain().focus().updateAttributes('docProtected', { textboxes }).run()
  }
  const key = axis === 'h' ? 'imageFlipH' : 'imageFlipV'
  const now = target.getAttributes('docProtected')[key]
  return target
    .chain()
    .focus()
    .updateAttributes('docProtected', { [key]: !now })
    .run()
}

export const POSITIONS_H = ['left', 'center', 'right'] as const
export const POSITIONS_V = ['top', 'center', 'bottom'] as const
export type PositionH = (typeof POSITIONS_H)[number]
export type PositionV = (typeof POSITIONS_V)[number]

/** a Position preset as picture block attrs: square wrap away from its side, against the margins */
const presetAttrs = (h: PositionH, v: PositionV) => ({
  imageWrap: h === 'right' ? 'square-right' : 'square-left',
  imagePosH: h,
  imagePosV: v,
  // the frame the save writes the pair against, so the page shows the saved place
  imageRelH: 'margin',
  imageRelV: 'margin',
  imageOffsetXEmu: null,
  imageOffsetYEmu: null,
})

/**
 * Position > a preset. A picture among text has no place on the page while it
 * stays in its line, so it leaves the line first: it becomes a picture block
 * anchored just above its paragraph, in one undo step with the placement.
 */
export function positionSelectedPicture(
  editor: Editor,
  sub: Editor | null,
  h: PositionH,
  v: PositionV,
): boolean {
  if (!arrangeStateOf(editor, sub).canPosition) return false
  const target = arrangeEditor(editor, sub)
  const run = target === editor ? runPictureOf(editor) : null
  if (!run) return target.chain().focus().updateAttributes('docProtected', presetAttrs(h, v)).run()
  const { state, view } = editor
  const a = run.attrs
  const src = String(a.dataUrl)
  const raster = RASTER_SRC.exec(src)
  const widthPx = Number(a.widthPx)
  const heightPx = Number(a.heightPx)
  const block = state.schema.nodes.docProtected.create({
    docxIndex: null,
    blockType: 'image',
    label: 'Picture',
    imageDataUrl: src,
    imageVector: a.vector ?? null,
    imageWidthPx: widthPx,
    imageHeightPx: heightPx,
    imageAlt: a.altText ?? null,
    imageRotDeg: a.rotDeg ?? null,
    imageFlipH: !!a.flipH,
    imageFlipV: !!a.flipV,
    // a picture still in the file's media is written from there (no bytes at hand)
    ...(raster ? { genImage: { base64: raster[2], mime: raster[1], widthPx, heightPx } } : {}),
    // as the file will read: the anchor's own empty paragraph keeps a line, and
    // the text stands off the sides by the distance a new anchor is written with
    anchorLine: { styleId: null },
    imageWrapDistTopEmu: 0,
    imageWrapDistBottomEmu: 0,
    imageWrapDistLeftEmu: NEW_ANCHOR_DIST_EMU,
    imageWrapDistRightEmu: NEW_ANCHOR_DIST_EMU,
    ...presetAttrs(h, v),
  })
  const $pos = state.doc.resolve(run.pos)
  const at = $pos.before(1)
  const tr = state.tr
  // alone in its paragraph the picture is that paragraph: the block takes its place
  if ($pos.parent.childCount === 1) tr.replaceWith(at, $pos.after(1), block)
  else tr.delete(run.pos, run.pos + 1).insert(at, block)
  view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, at)).scrollIntoView())
  view.focus()
  return true
}
