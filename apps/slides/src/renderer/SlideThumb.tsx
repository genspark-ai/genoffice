/**
 * #4 Thumbnail content rendering — draws a RenderSlide scaled down into a thumbnail (read-only, no interaction).
 * Shares NodeBody/StaticNode static rendering with the main canvas, keeping thumbnails and the full view identical.
 */
import React from 'react'
import { Stage, Layer, Rect } from 'react-konva'
import type Konva from 'konva'
import type { RenderSlide } from '@genoffice/pptx-render'
import { fillToKonva } from './konva-adapter'
import { StaticNode } from './NodeBody'

// Fallback width (px): matches the default 150px sidebar minus padding 20 and .thumb border 4;
// the sidebar itself is drag-resizable and always passes an explicit width
const THUMB_W = 126

// memo: App-level state changes (zoom gestures especially) must not re-render every thumbnail Stage
export const SlideThumb = React.memo(function SlideThumb({
  slide,
  images,
  width = THUMB_W,
  stageRef,
}: {
  slide: RenderSlide
  images: Map<string, HTMLImageElement>
  /** Render width (px): sidebar thumbnails default to 116; slide sorter/reading view pass larger sizes */
  width?: number
  /** Get the underlying Konva Stage (for offscreen toDataURL when exporting PNG/PDF) */
  stageRef?: (stage: Konva.Stage | null) => void
}) {
  const scale = width / slide.widthPx
  const h = slide.heightPx * scale
  return (
    <Stage
      ref={stageRef}
      width={width}
      height={h}
      listening={false}
      style={{ pointerEvents: 'none' }}
    >
      <Layer scaleX={scale} scaleY={scale} listening={false}>
        <Rect
          x={0}
          y={0}
          width={slide.widthPx}
          height={slide.heightPx}
          {...bgFill(slide, images)}
        />
        {slide.nodes.map((n) => (
          <StaticNode key={n.id} node={n} images={images} />
        ))}
      </Layer>
    </Stage>
  )
})

function bgFill(slide: RenderSlide, images: Map<string, HTMLImageElement>) {
  const f = fillToKonva(slide.background, slide.widthPx, slide.heightPx, images)
  return Object.keys(f).length ? f : { fill: '#ffffff' }
}
