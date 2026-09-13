import type { ShapeRenderNode } from '@genoffice/pptx-render'

/** Text frames need a full-box hit target because their glyph runs do not cover spacing/insets. */
export function needsTextFrameHitArea(shape: ShapeRenderNode): boolean {
  return !!shape.text
}

const PROMPT_PLACEHOLDERS = new Set(['title', 'ctrTitle', 'subTitle', 'body'])

/** Empty placeholder that draws a click-to-add prompt in its frame. */
export function isPromptPlaceholder(shape: ShapeRenderNode): boolean {
  if (!shape.placeholder || !PROMPT_PLACEHOLDERS.has(shape.placeholder)) return false
  return !shape.text?.lines.some((l) => l.runs.some((r) => !r.isBullet && r.text.trim()))
}

/**
 * PowerPoint/WPS: a single click on the text itself places the caret (edit mode), while
 * the frame around the text still selects the shape for dragging. The text region is the
 * union of the laid-out line boxes (glyph extent plus a small pad, in box-local px); an
 * empty prompt placeholder counts as text over its whole frame.
 */
export function textHitAtPoint(
  shape: ShapeRenderNode,
  box: { w: number; h: number },
  p: { x: number; y: number },
  padPx = 4,
): boolean {
  // Empty placeholders carry no layout at all; their prompt makes the whole frame text
  if (isPromptPlaceholder(shape)) return p.x >= 0 && p.y >= 0 && p.x <= box.w && p.y <= box.h
  const text = shape.text
  if (!text || text.txWarp) return false
  // vert/vert270 lines keep pre-rotation geometry; eaVert/wordArtVert columns are real boxes
  if (text.vert === 'vert' || text.vert === 'vert270') return false
  for (const line of text.lines) {
    const top = text.insets.t + line.top
    if (p.y < top - padPx || p.y > top + line.height + padPx) continue
    let x0 = Infinity
    let x1 = -Infinity
    for (const r of line.runs) {
      if (!r.text) continue
      // A rotated Latin word in a column: x is the rotation anchor at the glyph box's right edge
      x0 = Math.min(x0, r.rotate90 ? r.x - r.fontSizePx : r.x)
      x1 = Math.max(x1, r.rotate90 ? r.x : r.x + r.widthPx)
    }
    if (x0 === Infinity) continue
    if (p.x >= text.insets.l + x0 - padPx && p.x <= text.insets.l + x1 + padPx) return true
  }
  return false
}
