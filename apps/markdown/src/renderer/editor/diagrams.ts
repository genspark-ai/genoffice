import type { NewImage } from '@genoffice/docx-engine'
import { MERMAID_LANGUAGE, renderMermaid } from './mermaid'
import { WAVEDROM_LANGUAGE, renderWavedrom } from './wavedrom'

/** Fenced-code languages that render as a picture instead of source */
export const DIAGRAM_LANGUAGES = [MERMAID_LANGUAGE, WAVEDROM_LANGUAGE] as const
export type DiagramLanguage = (typeof DIAGRAM_LANGUAGES)[number]

export type DiagramResult = { ok: true; svg: string } | { ok: false; error: string }

export function diagramLanguage(language: unknown): DiagramLanguage | null {
  return (DIAGRAM_LANGUAGES as readonly unknown[]).includes(language)
    ? (language as DiagramLanguage)
    : null
}

export function renderDiagram(language: DiagramLanguage, source: string): Promise<DiagramResult> {
  return language === WAVEDROM_LANGUAGE ? renderWavedrom(source) : renderMermaid(source)
}

const VIEWBOX_RE = /\bviewBox="[\d.\s-]*?\s([\d.]+)\s([\d.]+)"/

/** Rasterize a rendered diagram for the docx export; null when it cannot be drawn */
export async function diagramSvgToPng(svg: string, maxWidthPx: number): Promise<NewImage | null> {
  const box = VIEWBOX_RE.exec(svg)
  const width = Math.ceil(Number(box?.[1]))
  const height = Math.ceil(Number(box?.[2]))
  if (!width || !height) return null

  // pin the intrinsic size (mermaid emits width="100%") so <img> decodes at the
  // viewBox dimensions instead of the 300×150 SVG default
  const sized = svg.replace(/<svg\b[^>]*>/, (tag) => {
    const stripped = tag.replace(/\s(?:width|height)="[^"]*"/g, '')
    return stripped.replace(/^<svg\b/, `<svg width="${width}" height="${height}"`)
  })
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
