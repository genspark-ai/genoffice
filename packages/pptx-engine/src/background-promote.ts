/// Full-page background normalization. The cloud html→pptx converter encodes the
/// page background as ordinary bottom-of-z-order full-page shapes when its own
/// promotion heuristics miss (e.g. a 1px fully-transparent border on the
/// container). Such shapes swallow every click/marquee on the slide, so:
/// - promoteSlideBackground rewrites them into a native <p:bg> at landing time
///   (strict conditions, destructive);
/// - isBackgroundLikeElement is the loose predicate the render layer uses to mark
///   surviving full-page backgrounds (existing files, image backgrounds) as
///   interaction-layer background nodes (non-destructive).

import type {
  PictureElement,
  ResolvedColor,
  Slide,
  SlideElement,
  SlideSize,
  Stroke,
  TextElement,
  Transform,
} from './types'
import { patchSlideBackgroundXml } from './generate'
import { elementSpid } from './animation'

const EMU_PER_PX = 9525
const COVER_TOL = 2 * EMU_PER_PX

function coversPage(t: Transform, size: SlideSize, maxAreaRatio: number): boolean {
  const o = t.offset
  return (
    o.x <= COVER_TOL &&
    o.y <= COVER_TOL &&
    o.x + o.cx >= size.cx - COVER_TOL &&
    o.y + o.cy >= size.cy - COVER_TOL &&
    o.cx * o.cy <= size.cx * size.cy * maxAreaRatio
  )
}

function isOpaqueColor(c: ResolvedColor): boolean {
  return /^#?[0-9a-fA-F]{6}$/.test(c) || /^#?[0-9a-fA-F]{6}[fF]{2}$/.test(c)
}

/** No stroke, zero width, or a fully transparent stroke color (#RRGGBB00). */
function strokeInvisible(stroke: Stroke | undefined): boolean {
  if (!stroke || stroke.width === 0 || stroke.fill.type === 'none') return true
  return (
    stroke.fill.type === 'solid' &&
    (stroke.fill.color === 'none' || /^#?[0-9a-fA-F]{6}00$/.test(stroke.fill.color))
  )
}

function hasVisibleText(el: TextElement): boolean {
  return !!el.text?.paragraphs.some((p) => p.runs.some((r) => r.text.trim() !== ''))
}

/**
 * Strict predicate for the destructive rewrite: an unrotated, unflipped,
 * non-placeholder full-page rect with an opaque solid fill, invisible stroke, no
 * effects and no text — nothing is lost by replacing it with <p:bg>.
 */
function isPromotableBackgroundShape(el: SlideElement, size: SlideSize): el is TextElement {
  if (el.type !== 'shape' && el.type !== 'text') return false
  const t = el as TextElement
  const tr = el.transform
  return (
    !el.placeholder &&
    tr.rot === 0 &&
    !tr.flipH &&
    !tr.flipV &&
    coversPage(tr, size, 1.05) &&
    (t.presetGeometry ?? 'rect') === 'rect' &&
    !t.customGeometry &&
    t.fill?.type === 'solid' &&
    isOpaqueColor(t.fill.color) &&
    strokeInvisible(t.stroke) &&
    !t.shadow &&
    !t.glow &&
    !hasVisibleText(t)
  )
}

/** Animations target shapes by cNvPr id; a referenced shape must not be deleted. */
function isReferencedByTiming(slide: Slide, el: SlideElement): boolean {
  const id = elementSpid(el)
  if (id == null) return true
  const ref = `spid="${id}"`
  return slide.bodySuffix.includes(ref) || slide.bodyPrefix.includes(ref)
}

/**
 * Promote the bottom-of-z-order contiguous run of full-page opaque solid shapes
 * into a native <p:bg> (topmost color wins — it is the one visible) and delete
 * them. Returns whether anything was promoted; on success the slide is
 * structureDirty and its model background is updated.
 */
export function promoteSlideBackground(slide: Slide, size: SlideSize): boolean {
  const leading: TextElement[] = []
  for (const el of slide.elements) {
    if (!isPromotableBackgroundShape(el, size) || isReferencedByTiming(slide, el)) break
    leading.push(el)
  }
  if (leading.length === 0) return false
  const fill = leading[leading.length - 1]!.fill
  if (fill?.type !== 'solid') return false
  slide.elements.splice(0, leading.length)
  slide.bodyPrefix = patchSlideBackgroundXml(slide.bodyPrefix, fill.color)
  slide.background = { type: 'solid', color: fill.color }
  slide.structureDirty = true
  return true
}

/**
 * Loose predicate for interaction-layer marking: a full-page unrotated textless
 * fill (solid/gradient/image shape, or a picture) with an invisible stroke.
 * Marked nodes stay visible and editable but stop swallowing marquee drags.
 */
export function isBackgroundLikeElement(el: SlideElement, size: SlideSize): boolean {
  if (el.placeholder || el.transform.rot !== 0) return false
  if (!coversPage(el.transform, size, 1.5)) return false
  if (el.type === 'picture') {
    const p = el as PictureElement
    return !p.media && strokeInvisible(p.stroke)
  }
  if (el.type !== 'shape' && el.type !== 'text') return false
  const t = el as TextElement
  const fillKind = t.fill?.type
  return (
    (fillKind === 'solid' || fillKind === 'gradient' || fillKind === 'image') &&
    (t.presetGeometry ?? 'rect') === 'rect' &&
    !t.customGeometry &&
    strokeInvisible(t.stroke) &&
    !hasVisibleText(t)
  )
}
