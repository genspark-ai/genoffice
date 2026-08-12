import type { ShapeRenderNode } from '@genoffice/pptx-render'

/** Text frames need a full-box hit target because their glyph runs do not cover spacing/insets. */
export function needsTextFrameHitArea(shape: ShapeRenderNode): boolean {
  return !!shape.text
}
