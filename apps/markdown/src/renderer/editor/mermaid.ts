import type { Mermaid } from 'mermaid'
import type { NewImage } from '@genoffice/docx-engine'

export const MERMAID_LANGUAGE = 'mermaid'

export const MERMAID_TEMPLATE = [
  '```mermaid',
  'flowchart LR',
  '    A[Start] --> B{Decision}',
  '    B -->|Yes| C[Done]',
  '    B -->|No| A',
  '```',
].join('\n')

let loading: Promise<Mermaid> | null = null

/** Loaded on first use so documents without diagrams never pay for the bundle */
export function loadMermaid(): Promise<Mermaid> {
  loading ??= import('mermaid').then(({ default: mermaid }) => {
    mermaid.initialize({
      startOnLoad: false,
      securityLevel: 'strict',
      // diagrams are document content: one fixed theme in both UI themes
      theme: 'default',
      // SVG text labels instead of <foreignObject> so the diagram can be drawn
      // onto a canvas (docx export) without tainting it
      htmlLabels: false,
      flowchart: { htmlLabels: false },
      suppressErrorRendering: true,
    })
    return mermaid
  })
  return loading
}

export type MermaidResult = { ok: true; svg: string } | { ok: false; error: string }

let renderSeq = 0

export async function renderMermaid(source: string): Promise<MermaidResult> {
  const mermaid = await loadMermaid()
  try {
    const { svg } = await mermaid.render(`md-mermaid-${++renderSeq}`, source)
    return { ok: true, svg }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

const VIEWBOX_RE = /\bviewBox="[\d.\s-]*?\s([\d.]+)\s([\d.]+)"/

/** Rasterize a rendered diagram for the docx export; null when it cannot be drawn */
export async function mermaidSvgToPng(svg: string, maxWidthPx: number): Promise<NewImage | null> {
  const box = VIEWBOX_RE.exec(svg)
  const width = Math.ceil(Number(box?.[1]))
  const height = Math.ceil(Number(box?.[2]))
  if (!width || !height) return null

  // mermaid emits width="100%": pin the intrinsic size so <img> decodes at the
  // viewBox dimensions instead of the 300×150 SVG default
  const sized = svg.replace(
    /<svg\b([^>]*?)\swidth="[^"]*"/,
    `<svg$1 width="${width}" height="${height}"`,
  )
  const url = URL.createObjectURL(new Blob([sized], { type: 'image/svg+xml' }))
  try {
    const img = new Image()
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve()
      img.onerror = () => reject(new Error('svg decode failed'))
      img.src = url
    })
    const scale = 2
    const canvas = document.createElement('canvas')
    canvas.width = width * scale
    canvas.height = height * scale
    const ctx = canvas.getContext('2d')
    if (!ctx) return null
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height)
    const dataUrl = canvas.toDataURL('image/png')
    const base64 = dataUrl.slice(dataUrl.indexOf(',') + 1)
    if (!base64) return null
    const widthPx = Math.min(width, maxWidthPx)
    const heightPx = Math.round((height * widthPx) / width)
    return { base64, mime: 'image/png', widthPx, heightPx, align: 'center' }
  } catch {
    return null
  } finally {
    URL.revokeObjectURL(url)
  }
}
