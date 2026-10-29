import {
  applyImageWrap,
  resizeAnchoredPictureXml,
  resizeInlinePictureXml,
  swapAnchorWrap,
} from './generate'
import { HF_DRAWING_RE, runPictureAnchorOf } from './parse'
import type { HfPictureEdit, ImageWrap } from './types'

const EMU_PER_PX = 9525

/**
 * Write the floating pictures a header/footer strip moved or resized back into
 * the part: each edited <w:drawing> (by its HfImage.drawing ordinal) is pinned
 * to the page at its new top-left corner and takes the new extents. Every
 * other byte of the part is kept.
 */
export function patchHfPicturesXml(partXml: string, edits: Record<number, HfPictureEdit>): string {
  let ordinal = -1
  return partXml.replace(HF_DRAWING_RE, (frag) => {
    ordinal++
    const edit = edits[ordinal]
    if (!edit || !/<wp:anchor[\s>]/.test(frag)) return frag
    const emu = (px: number) => Math.round(px * EMU_PER_PX)
    const out = frag
      // simplePos="1" places by wp:simplePos instead of the position elements
      .replace(/(<wp:anchor\b[^>]*\bsimplePos=")(?:1|true)"/, '$10"')
      .replace(
        /<wp:positionH\b[\s\S]*?<\/wp:positionH>/,
        `<wp:positionH relativeFrom="page"><wp:posOffset>${emu(edit.posXPx)}</wp:posOffset></wp:positionH>`,
      )
      .replace(
        /<wp:positionV\b[\s\S]*?<\/wp:positionV>/,
        `<wp:positionV relativeFrom="page"><wp:posOffset>${emu(edit.posYPx)}</wp:posOffset></wp:positionV>`,
      )
    return resizeAnchoredPictureXml(out, edit.widthPx, edit.heightPx)
  })
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
}

/**
 * Write a run-level picture's edited geometry back into its <w:drawing>
 * fragment. Only what differs from the fragment's own geometry is rewritten,
 * so a picture nobody touched keeps its bytes: a wrap change swaps the wrap
 * element (or converts in-line <-> floating), a move writes numeric
 * posOffsets on the bases the editor positioned it from, a resize the extents.
 * VML pictures (no wp:inline / wp:anchor) come back unchanged.
 */
export function patchRunPictureXml(xml: string, next: RunPictureGeometry): string {
  if (!/<wp:(?:inline|anchor)[\s>]/.test(xml)) return xml
  const cur = runPictureAnchorOf(xml)
  const wrap = next.wrap ?? null
  const x = next.offsetXEmu ?? null
  const y = next.offsetYEmu ?? null
  let out = xml
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

const wrapKind = (wrap: ImageWrap | null): string | null =>
  wrap ? wrap.replace(/-(?:left|right)$/, '') : null

const wrapSide = (wrap: ImageWrap | null): 'left' | 'right' | null =>
  (/-(left|right)$/.exec(wrap ?? '')?.[1] as 'left' | 'right' | undefined) ?? null

function resize(xml: string, next: RunPictureGeometry): string {
  const w = Number(next.widthPx)
  const h = Number(next.heightPx)
  if (!(w > 0) || !(h > 0)) return xml
  return /<wp:anchor[\s>]/.test(xml)
    ? resizeAnchoredPictureXml(xml, w, h)
    : resizeInlinePictureXml(xml, w, h)
}
