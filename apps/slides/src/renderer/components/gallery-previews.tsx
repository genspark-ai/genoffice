/**
 * Geometry-driven gallery previews. The shape geometry/preview implementation
 * is shared across apps via @genoffice/ui (re-exported here for existing
 * imports); SmartArt thumbs reuse the engine's layout math so preview and
 * insert result always match.
 */
import { shapePreviewPath } from '@genoffice/ui'
import {
  layoutShapes,
  type SmartArtLayout,
} from '../../../../../packages/pptx-engine/src/smartart-layout'

export { ShapePreview, shapePreviewBox, shapePreviewPath } from '@genoffice/ui'

/** Same virtual canvas scale as real insertion (EMU) so layout ratios/minimums behave identically. */
const SA_CX = 4800000
const SA_CY = 3000000
/** Node counts mirror SMARTART_GALLERY defaultItems */
const SA_ITEMS: Record<SmartArtLayout, number> = {
  list: 3,
  process: 4,
  cycle: 4,
  hierarchy: 4,
  pyramid: 3,
  matrix: 4,
  venn: 3,
}

export function smartArtPreviewShapes(
  layout: SmartArtLayout,
): Array<{ d: string; color: string; x: number; y: number; alpha?: number }> {
  const items = Array.from({ length: SA_ITEMS[layout] }, () => '')
  return layoutShapes(layout, items, SA_CX, SA_CY).map((s) => ({
    d: shapePreviewPath(s.prst, s.box.cx, s.box.cy) ?? '',
    color: `#${s.color}`,
    x: s.box.x,
    y: s.box.y,
    alpha: s.alpha,
  }))
}

export function SmartArtPreview({
  layout,
  width = 44,
}: {
  layout: SmartArtLayout
  width?: number
}) {
  return (
    <svg
      width={width}
      height={(width * SA_CY) / SA_CX}
      viewBox={`0 0 ${SA_CX} ${SA_CY}`}
      aria-hidden
    >
      {smartArtPreviewShapes(layout).map((s, i) => (
        <path
          key={i}
          d={s.d}
          transform={`translate(${s.x} ${s.y})`}
          fill={s.color}
          fillOpacity={s.alpha}
        />
      ))}
    </svg>
  )
}
