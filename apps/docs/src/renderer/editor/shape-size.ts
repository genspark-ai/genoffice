import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import {
  relSizeTwips,
  type RelSize,
  type SectionSettings,
  type TextboxDisplay,
} from '@genoffice/docx-engine'
import { selectedBoxIndex } from './selected-box'

const TWIPS_PER_PX = 15
const MIN_SIDE_PX = 1

/** one side of a shape in Size and Position: so many px, or a share of the page */
export interface ShapeSide {
  px: number
  rel: RelSize | null
}

/** what Size and Position shows of the selected shape or text box */
export interface ShapeSizeGeometry {
  w: ShapeSide
  /** null: the shape is as tall as its text */
  h: ShapeSide | null
}

/** a box whose size the dialog can set: a shape or text box of its own, not a line, a picture or a group member */
function sizable(box: TextboxDisplay | undefined): box is TextboxDisplay {
  return (
    !!box &&
    !box.picture &&
    !box.grouped &&
    !box.lineVert &&
    !box.lineDiag &&
    !(box.prst ?? '').startsWith('line') &&
    (box.widthPx ?? 0) > 0
  )
}

function selectedBox(editor: Editor): { boxes: TextboxDisplay[]; index: number } | null {
  const boxes = editor.getAttributes('docProtected')?.textboxes as TextboxDisplay[] | null
  if (!Array.isArray(boxes) || boxes.length === 0) return null
  const index = selectedBoxIndex(editor.state)
  return sizable(boxes[index]) ? { boxes, index } : null
}

export function selectedShapeSize(editor: Editor): ShapeSizeGeometry | null {
  const sel = selectedBox(editor)
  if (!sel) return null
  const box = sel.boxes[sel.index]
  return {
    w: { px: box.widthPx!, rel: box.sizeRel?.w ?? null },
    h: box.heightPx ? { px: box.heightPx, rel: box.sizeRel?.h ?? null } : null,
  }
}

/** the px a relative side comes to on a page of the section */
export const relSidePx = (section: SectionSettings, axis: 'H' | 'V', rel: RelSize): number =>
  Math.max(MIN_SIDE_PX, Math.round(relSizeTwips(section, axis, rel) / TWIPS_PER_PX))

/** the box with its sides set; a side given without a share is absolute from here on */
export function boxWithSize(
  box: TextboxDisplay,
  size: { w?: ShapeSide; h?: ShapeSide },
): TextboxDisplay {
  const { sizeRel: was, ...rest } = box
  const w = size.w ? size.w.rel : (was?.w ?? null)
  const h = size.h ? size.h.rel : (was?.h ?? null)
  const sizeRel = w || h ? { ...(w ? { w } : {}), ...(h ? { h } : {}) } : null
  return {
    ...rest,
    ...(size.w ? { widthPx: Math.max(MIN_SIDE_PX, Math.round(size.w.px)) } : {}),
    ...(size.h
      ? {
          heightPx: Math.max(MIN_SIDE_PX, Math.round(size.h.px)),
          minHeightPx: Math.max(MIN_SIDE_PX, Math.round(size.h.px)),
        }
      : {}),
    ...(sizeRel ? { sizeRel } : {}),
  }
}

/** Size and Position's OK for a shape: both sides in one undo step. */
export function applyShapeSize(
  editor: Editor,
  next: ShapeSizeGeometry,
  section: SectionSettings | null | undefined,
): boolean {
  // what is being typed into the box goes into its model first, or the new model would overwrite it
  window.dispatchEvent(new Event('ai-docs-commit-tables'))
  const sel = selectedBox(editor)
  if (!sel || !editor.isEditable) return false
  const side = (axis: 'H' | 'V', s: ShapeSide): ShapeSide =>
    s.rel && section ? { px: relSidePx(section, axis, s.rel), rel: s.rel } : { px: s.px, rel: null }
  const box = boxWithSize(sel.boxes[sel.index], {
    w: side('H', next.w),
    ...(next.h ? { h: side('V', next.h) } : {}),
  })
  return editor
    .chain()
    .focus()
    .updateAttributes('docProtected', {
      textboxes: sel.boxes.map((b, i) => (i === sel.index ? box : b)),
    })
    .run()
}

/** the boxes of a node resized to the shares they state, on a page of the section; null when none moves */
export function boxesFittedToSection(
  node: PmNode,
  section: SectionSettings,
): TextboxDisplay[] | null {
  const boxes = node.attrs.textboxes as TextboxDisplay[] | null
  if (!Array.isArray(boxes) || !boxes.some((box) => box.sizeRel)) return null
  let moved = false
  const next = boxes.map((box) => {
    if (!box.sizeRel || !sizable(box)) return box
    const w = box.sizeRel.w ? relSidePx(section, 'H', box.sizeRel.w) : null
    // a shape as tall as its text keeps following the text
    const h = box.sizeRel.h && box.heightPx ? relSidePx(section, 'V', box.sizeRel.h) : null
    if ((w === null || w === box.widthPx) && (h === null || h === box.heightPx)) return box
    moved = true
    return boxWithSize(box, {
      ...(w !== null ? { w: { px: w, rel: box.sizeRel.w! } } : {}),
      ...(h !== null ? { h: { px: h, rel: box.sizeRel.h! } } : {}),
    })
  })
  return moved ? next : null
}

/** the lengths of a section a relative size can be a share of, as one comparable value */
export const pageGeometryKey = (s: SectionSettings | null | undefined): string =>
  s
    ? [s.pageWidth, s.pageHeight, s.marginTop, s.marginRight, s.marginBottom, s.marginLeft].join()
    : ''

/**
 * Shapes sized as a share of the page take the size the page now gives them.
 * `sectionAt` names the section of a top-level block, or null to leave its
 * shapes alone. The step is not one of the user's: undoing the page setup
 * runs it again.
 */
export function fitRelativeShapes(
  editor: Editor,
  sectionAt: (topIndex: number) => SectionSettings | null,
): boolean {
  const { state, view } = editor
  const tr = state.tr
  state.doc.forEach((top, offset, topIndex) => {
    let section: SectionSettings | null | undefined
    const fit = (node: PmNode, pos: number) => {
      if (node.type.name !== 'docProtected') return
      section ??= sectionAt(topIndex)
      const textboxes = section && boxesFittedToSection(node, section)
      if (textboxes) tr.setNodeMarkup(pos, undefined, { ...node.attrs, textboxes })
    }
    fit(top, offset)
    top.descendants((node, pos) => fit(node, offset + 1 + pos))
  })
  if (!tr.docChanged) return false
  view.dispatch(tr.setMeta('addToHistory', false))
  return true
}
