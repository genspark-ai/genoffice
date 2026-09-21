/**
 * Offscreen rendering for export — draws each RenderSlide page into an offscreen Konva Stage
 * and exports high-resolution PNGs. Reuses SlideThumb (same rendering as the main canvas and
 * thumbnails), guaranteeing the export matches what the editor shows.
 */
import React from 'react'
import { createRoot } from 'react-dom/client'
import type Konva from 'konva'
import { Layer, Stage } from 'react-konva'
import type { RenderNode, RenderSlide } from '@genoffice/pptx-render'
import { SlideThumb } from './SlideThumb'
import { StaticNode } from './NodeBody'
import { rotatedBounds, type NodeRaster } from './export-svg'

/** Pixel ratio of the exported bitmap (2x hi-res, 1280 viewport width → 2560px PNG) */
const EXPORT_PIXEL_RATIO = 2

/**
 * Render each page to PNG base64 (without the data: prefix).
 * Reuse a single offscreen root page by page, grabbing each page as it's drawn, so the whole
 * deck never sits in memory at once.
 * pixelRatio 1 is enough for AI-vision screenshots (half the tokens of the 2x export default).
 */
export async function renderSlidesToPngBase64(
  slides: RenderSlide[],
  images: Map<string, HTMLImageElement>,
  pixelRatio: number = EXPORT_PIXEL_RATIO,
): Promise<string[]> {
  // Offscreen container: mounted outside the body viewport (display:none would give the Konva canvas zero size, unusable)
  const container = document.createElement('div')
  container.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none;'
  document.body.appendChild(container)
  const root = createRoot(container)
  const out: string[] = []
  try {
    for (const slide of slides) {
      const stage = await new Promise<Konva.Stage>((resolve) => {
        root.render(
          <SlideThumb
            slide={slide}
            images={images}
            width={slide.widthPx}
            stageRef={(s) => s && resolve(s)}
          />,
        )
      })
      // Wait one frame for Konva to finish batchDraw, then capture
      await new Promise((r) => requestAnimationFrame(r))
      const dataUrl = stage.toDataURL({ mimeType: 'image/png', pixelRatio })
      out.push(dataUrl.replace(/^data:image\/png;base64,/, ''))
    }
  } finally {
    root.unmount()
    container.remove()
  }
  return out
}

/** Ink allowance around a node's box for shadows, glows and reflections (px). */
const NODE_RASTER_PAD = 48

/**
 * One top-level node alone on a transparent stage, cropped to its rotated bounds:
 * the vector export's fallback for effects SVG cannot express (WordArt warp).
 */
export async function rasterizeSlideNode(
  node: RenderNode,
  slide: RenderSlide,
  images: Map<string, HTMLImageElement>,
  pixelRatio: number = EXPORT_PIXEL_RATIO,
): Promise<NodeRaster | null> {
  const b = rotatedBounds(node.box)
  const x = Math.floor(b.x - NODE_RASTER_PAD)
  const y = Math.floor(b.y - NODE_RASTER_PAD)
  const w = Math.ceil(b.w + 2 * NODE_RASTER_PAD)
  const h = Math.ceil(b.h + 2 * NODE_RASTER_PAD)
  if (w < 1 || h < 1) return null
  const container = document.createElement('div')
  container.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none;'
  document.body.appendChild(container)
  const root = createRoot(container)
  try {
    const stage = await new Promise<Konva.Stage>((resolve) => {
      root.render(
        <Stage
          width={slide.widthPx}
          height={slide.heightPx}
          listening={false}
          ref={(s) => {
            if (s) resolve(s)
          }}
        >
          <Layer listening={false}>
            <StaticNode node={node} images={images} />
          </Layer>
        </Stage>,
      )
    })
    await new Promise((r) => requestAnimationFrame(r))
    const href = stage.toDataURL({ mimeType: 'image/png', pixelRatio, x, y, width: w, height: h })
    return { href, x, y, w, h }
  } finally {
    root.unmount()
    container.remove()
  }
}
