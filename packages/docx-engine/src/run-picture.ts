import {
  applyImageWrap,
  applyImageZOrder,
  patchImageParagraphXml,
  resizeAnchoredPictureXml,
  resizeInlinePictureXml,
  setDrawingAltText,
  swapAnchorWrap,
} from './generate'
import { HF_DRAWING_RE, hfMovablePictureFrag, picTransformOf, runPictureAnchorOf } from './parse'
import { decodeEntities } from './parse-xml-text'
import type { HfPictureEdit, ImageWrap } from './types'

const EMU_PER_PX = 9525

/**
 * Write the floating pictures a header/footer strip moved or resized back into
 * the part: each edited <w:drawing> (by its HfImage.drawing ordinal) takes
 * the new extents, and each axis it moved along is pinned to the page at the
 * new top-left corner. An axis it did not move keeps its own base and offset,
 * so a resize from the far edge leaves a margin-relative banner relative to
 * the margin. Every other byte of the part is kept.
 */
export function patchHfPicturesXml(partXml: string, edits: Record<number, HfPictureEdit>): string {
  let ordinal = -1
  return partXml.replace(HF_DRAWING_RE, (frag) => {
    ordinal++
    const edit = edits[ordinal]
    if (!edit || !hfMovablePictureFrag(frag)) return frag
    let out = frag
    if (edit.posXPx != null || edit.posYPx != null) {
      // simplePos="1" places by wp:simplePos instead of the position elements,
      // which the strip (like the parser) positioned the picture from
      out = out.replace(/(<wp:anchor\b[^>]*\bsimplePos=")(?:1|true)"/, '$10"')
      out = setAnchorPosition(out, 'H', edit.posXPx)
      out = setAnchorPosition(out, 'V', edit.posYPx)
    }
    return resizeAnchoredPictureXml(out, edit.widthPx, edit.heightPx)
  })
}

/**
 * An anchor's wp:positionH / wp:positionV set to a page-relative offset, or
 * left alone when px is absent. A missing element is inserted where the
 * schema orders it (after wp:simplePos, V after H), so the file agrees with
 * the screen.
 */
function setAnchorPosition(xml: string, axis: 'H' | 'V', px: number | undefined): string {
  if (px == null) return xml
  const emu = Math.round(px * EMU_PER_PX)
  const el = `<wp:position${axis} relativeFrom="page"><wp:posOffset>${emu}</wp:posOffset></wp:position${axis}>`
  const existing = new RegExp(`<wp:position${axis}\\b[\\s\\S]*?</wp:position${axis}>`)
  if (existing.test(xml)) return xml.replace(existing, el)
  // the element it follows: wp:positionH for V, else wp:simplePos, else the anchor's open tag
  const follows =
    (axis === 'V' ? /<wp:positionH\b[\s\S]*?<\/wp:positionH>/.exec(xml) : null) ??
    /<wp:simplePos\b[^>]*\/>|<wp:simplePos\b[\s\S]*?<\/wp:simplePos>/.exec(xml) ??
    /<wp:anchor\b[^>]*>/.exec(xml)
  if (!follows) return xml
  const end = follows.index + follows[0].length
  return xml.slice(0, end) + el + xml.slice(end)
}

/** The editor's geometry of a run-level picture (its docInlineImage attrs). */
export interface RunPictureGeometry {
  widthPx?: number | null
  heightPx?: number | null
  wrap?: ImageWrap | null
  offsetXEmu?: number | null
  offsetYEmu?: number | null
  relH?: 'page' | 'margin' | null
  relV?: 'page' | 'margin' | 'line' | null
  lineCenterV?: boolean | null
  /** degrees clockwise and mirror flips of the picture's own a:xfrm; undefined keeps the file's */
  rotDeg?: number | null
  flipH?: boolean | null
  flipV?: boolean | null
  /** stacking rank above Word's base relativeHeight; undefined keeps the file's */
  zOrder?: number
}

/**
 * Write a run-level picture's edited geometry back into its <w:drawing>
 * fragment. Only what differs from the fragment's own geometry is rewritten,
 * so a picture nobody touched keeps its bytes: a wrap change swaps the wrap
 * element (or converts in-line <-> floating), a move writes numeric
 * posOffsets on the bases the editor positioned it from, a resize the extents,
 * a turn or flip the picture's a:xfrm.
 * VML pictures (no wp:inline / wp:anchor) come back unchanged.
 */
export function patchRunPictureXml(xml: string, next: RunPictureGeometry): string {
  if (!/<wp:(?:inline|anchor)[\s>]/.test(xml)) return xml
  const cur = runPictureAnchorOf(xml)
  const wrap = next.wrap ?? null
  const x = next.offsetXEmu ?? null
  const y = next.offsetYEmu ?? null
  let out = turn(xml, next)
  if (next.zOrder !== undefined) out = applyImageZOrder(out, next.zOrder)
  const curWrap = cur.wrap ?? null
  if (wrap !== curWrap) {
    if (!wrap || !/<wp:anchor[\s>]/.test(out)) {
      // in-line <-> floating: the anchor is rebuilt; a known offset pins it
      // on the column / paragraph like a dragged block picture
      const placed = x != null || y != null
      out = applyImageWrap(out, wrap, placed ? { x: x ?? 0, y: y ?? 0 } : undefined)
      return resize(out, next)
    }
    // the side of a square/tight/through float is read off its position, so
    // a side flip keeps the wrap element (and a tight wrap's polygon)
    if (wrapKind(wrap) !== wrapKind(curWrap)) out = swapAnchorWrap(out, wrap)
    const side = wrapSide(wrap)
    if (x == null && side && side !== wrapSide(curWrap)) {
      out = out.replace(
        /<wp:positionH\b[\s\S]*?<\/wp:positionH>/,
        `<wp:positionH relativeFrom="column"><wp:align>${side}</wp:align></wp:positionH>`,
      )
    }
  }
  if (wrap && /<wp:anchor[\s>]/.test(out)) {
    const relH = next.relH ?? null
    const relV = next.relV ?? null
    if (x != null && (x !== (cur.offsetXEmu ?? null) || relH !== (cur.relH ?? null))) {
      out = out.replace(
        /<wp:positionH\b[\s\S]*?<\/wp:positionH>/,
        `<wp:positionH relativeFrom="${relH ?? 'column'}"><wp:posOffset>${Math.round(x)}</wp:posOffset></wp:positionH>`,
      )
    }
    if (
      y != null &&
      (y !== (cur.offsetYEmu ?? null) ||
        relV !== (cur.relV ?? null) ||
        !!next.lineCenterV !== !!cur.lineCenterV)
    ) {
      out = out.replace(
        /<wp:positionV\b[\s\S]*?<\/wp:positionV>/,
        `<wp:positionV relativeFrom="${relV ?? 'paragraph'}"><wp:posOffset>${Math.round(y)}</wp:posOffset></wp:positionV>`,
      )
    }
  }
  return resize(out, next)
}

/** the turn and flips, where they differ from the drawing's own; the effect extent follows a new turn */
function turn(xml: string, next: RunPictureGeometry): string {
  const cur = picTransformOf(xml)
  const patch: { rotDeg?: number; flipH?: boolean; flipV?: boolean } = {}
  const rot =
    next.rotDeg === undefined ? undefined : ((Math.round(next.rotDeg ?? 0) % 360) + 360) % 360
  if (rot !== undefined && rot !== (cur.rotDeg ?? 0)) patch.rotDeg = rot
  if (next.flipH !== undefined && !!next.flipH !== !!cur.flipH) patch.flipH = !!next.flipH
  if (next.flipV !== undefined && !!next.flipV !== !!cur.flipV) patch.flipV = !!next.flipV
  if (Object.keys(patch).length === 0) return xml
  const holder = withPictureXfrm(xml)
  return PICTURE_XFRM.test(holder) ? patchImageParagraphXml(holder, patch) : xml
}

const PICTURE_XFRM = /<pic:spPr\b[^>]*>[\s\S]*?<a:xfrm\b/

/** a picture whose pic:spPr has no a:xfrm gets one of its extent, for the turn to be written on */
function withPictureXfrm(xml: string): string {
  if (PICTURE_XFRM.test(xml)) return xml
  const extent = /<wp:extent\b[^>]*?\bcx="(\d+)"[^>]*?\bcy="(\d+)"/.exec(xml)
  if (!extent) return xml
  const xfrm = `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${extent[1]}" cy="${extent[2]}"/></a:xfrm>`
  if (/<pic:spPr\b/.test(xml)) {
    return xml.replace(/<pic:spPr\b([^>]*?)(\/?)>/, (_tag, attrs: string, empty: string) =>
      empty ? `<pic:spPr${attrs}>${xfrm}</pic:spPr>` : `<pic:spPr${attrs}>${xfrm}`,
    )
  }
  return xml.replace('</pic:blipFill>', `</pic:blipFill><pic:spPr>${xfrm}</pic:spPr>`)
}

const wrapKind = (wrap: ImageWrap | null): string | null =>
  wrap ? wrap.replace(/-(?:left|right)$/, '') : null

const wrapSide = (wrap: ImageWrap | null): 'left' | 'right' | null =>
  (/-(left|right)$/.exec(wrap ?? '')?.[1] as 'left' | 'right' | undefined) ?? null

function resize(xml: string, next: RunPictureGeometry): string {
  const w = Number(next.widthPx)
  const h = Number(next.heightPx)
  if (!(w > 0) || !(h > 0)) return xml
  const out = /<wp:anchor[\s>]/.test(xml)
    ? resizeAnchoredPictureXml(xml, w, h)
    : resizeInlinePictureXml(xml, w, h)
  // the room a turn takes is that of the new size
  const { rotDeg } = picTransformOf(out)
  return out === xml || !rotDeg ? out : patchImageParagraphXml(out, { rotDeg })
}

/** The alternative text of a run-level picture; null when its fragment has nowhere to hold one (VML). */
export function runPictureAltOf(xml: string): string | null {
  const docPr = /<wp:docPr\b[^>]*>/.exec(xml)?.[0]
  if (!docPr) return null
  const descr = /\bdescr\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(docPr)
  return decodeEntities(descr?.[1] ?? descr?.[2] ?? '')
}

export const setRunPictureAlt = setDrawingAltText
