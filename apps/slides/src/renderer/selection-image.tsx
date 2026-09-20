import React from 'react'
import { createRoot } from 'react-dom/client'
import { flushSync } from 'react-dom'
import { Stage, Layer } from 'react-konva'
import type Konva from 'konva'
import type { RenderFill, RenderNode, RenderSlide } from '@genoffice/pptx-render'
import { StaticNode } from './NodeBody'
import { createImageLoader } from './image-loader'

/** Load only the selection's resources, including picture fills and image bullets. */
async function selectionImages(nodes: RenderNode[], cached: Map<string, HTMLImageElement>) {
  const urls = new Set<string>()
  const fill = (f?: RenderFill) => {
    if (f?.kind === 'image' && f.dataUrl) urls.add(f.dataUrl)
  }
  const text = (t?: { lines: Array<{ runs: Array<{ image?: string }> }> }) => {
    for (const line of t?.lines ?? [])
      for (const run of line.runs) if (run.image) urls.add(run.image)
  }
  const walk = (items: RenderNode[]) => {
    for (const n of items) {
      if (n.type === 'picture' && n.dataUrl) urls.add(n.dataUrl)
      if (n.type === 'picture' || n.type === 'shape' || n.type === 'text') fill(n.fill)
      if (n.type === 'shape' || n.type === 'text') text(n.text)
      if (n.type === 'chart') {
        fill(n.bgFill)
        fill(n.plotRect?.fill)
      }
      if (n.type === 'table') fill(n.bgFill)
      if (n.type === 'group') walk(n.children)
      if (n.type === 'table')
        for (const c of n.cells) {
          fill(c.fill)
          text(c.text)
        }
    }
  }
  walk(nodes)
  const images = new Map(cached)
  const missing = [...urls].filter((url) => !images.get(url)?.naturalWidth)
  if (!missing.length) return images
  const loader = createImageLoader((entries) => {
    for (const [url, img] of entries) images.set(url, img)
  }, 1)
  try {
    loader.load(missing)
    const deadline = Date.now() + 5000
    while (loader.pending() && Date.now() < deadline) await new Promise((r) => setTimeout(r, 20))
    if (missing.some((url) => !images.get(url)?.naturalWidth))
      throw new Error('Selection image could not be loaded')
    return images
  } finally {
    loader.dispose()
  }
}

/** Transparent PNG of selected elements in slide z-order, cropped to their visual bounds. */
export async function renderSelectionToPngBase64(
  slide: RenderSlide,
  sourceIds: readonly string[],
  cachedImages: Map<string, HTMLImageElement>,
): Promise<string> {
  const selected = new Set(sourceIds)
  const nodes = slide.nodes.filter((n) => !n.decoration && selected.has(n.sourceId))
  if (!nodes.length) throw new Error('No selected elements to render')
  const images = await selectionImages(nodes, cachedImages)
  const container = document.createElement('div')
  container.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none;'
  document.body.appendChild(container)
  const root = createRoot(container)
  const layerRef = React.createRef<Konva.Layer>()
  try {
    // Rendering the layer directly into the crop also captures negative/off-page coordinates.
    flushSync(() =>
      root.render(
        <Stage width={1} height={1} listening={false}>
          <Layer ref={layerRef} listening={false}>
            {nodes.map((node) => (
              <StaticNode key={node.id} node={node} images={images} />
            ))}
          </Layer>
        </Stage>,
      ),
    )
    await document.fonts?.ready
    const layer = layerRef.current!
    layer.draw()
    const bounds = layer.getClientRect()
    if (
      ![bounds.x, bounds.y, bounds.width, bounds.height].every(Number.isFinite) ||
      bounds.width <= 0 ||
      bounds.height <= 0
    )
      throw new Error('Selection has no drawable bounds')
    // One pixel of transparent bleed preserves antialiasing at the crop edges.
    const x = Math.floor(bounds.x) - 1,
      y = Math.floor(bounds.y) - 1
    const width = Math.ceil(bounds.x + bounds.width) + 1 - x
    const height = Math.ceil(bounds.y + bounds.height) + 1 - y
    // Normal selections are 2x; oversized selections fit within a 16 MP / 8192px canvas.
    const pixelRatio = Math.min(
      2,
      8192 / width,
      8192 / height,
      Math.sqrt(16_000_000 / (width * height)),
    )
    const dataUrl = layer.toDataURL({ x, y, width, height, pixelRatio, mimeType: 'image/png' })
    if (!dataUrl.startsWith('data:image/png;base64,'))
      throw new Error('Selection PNG capture failed')
    return dataUrl.slice('data:image/png;base64,'.length)
  } finally {
    root.unmount()
    container.remove()
  }
}
