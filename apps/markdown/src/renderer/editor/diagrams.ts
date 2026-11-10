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

const VIEWBOX_RE =
  /\bviewBox\s*=\s*["']\s*(-?[\d.]+)[\s,]+(-?[\d.]+)[\s,]+(-?[\d.]+)[\s,]+(-?[\d.]+)\s*["']/

/** Fallback doc width when maxWidthPx is missing or invalid (matches DOCX_MAX_IMAGE_PX) */
export const DIAGRAM_FALLBACK_MAX_WIDTH_PX = 620
/** Sane upper bound for the requested doc width so callers cannot blow up layout */
export const DIAGRAM_MAX_WIDTH_PX = 2048
/** Hard cap for a canvas dimension; larger rasters are scaled down proportionally */
export const DIAGRAM_MAX_CANVAS_DIM_PX = 4096
/** HiDPI raster scale for diagram PNGs */
export const DIAGRAM_RASTER_SCALE = 2

/** Comma-tolerant viewBox parse; null when missing, malformed, or non-positive */
export function parseSvgViewBox(svg: string): { width: number; height: number } | null {
  if (typeof svg !== 'string') return null
  const box = VIEWBOX_RE.exec(svg)
  if (!box) return null
  const width = Math.ceil(Number(box[3]))
  const height = Math.ceil(Number(box[4]))
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return null
  }
  return { width, height }
}

/** Clamp the requested doc width to a positive finite value within a sane cap */
export function clampDiagramMaxWidth(value: unknown): number {
  const n = typeof value === 'number' ? value : Number(value)
  if (!Number.isFinite(n) || n <= 0) return DIAGRAM_FALLBACK_MAX_WIDTH_PX
  return Math.min(Math.floor(n), DIAGRAM_MAX_WIDTH_PX)
}

export interface DiagramRasterSize {
  widthPx: number
  heightPx: number
  canvasWidth: number
  canvasHeight: number
}

/**
 * Doc size plus backing canvas size for a viewBox; null for invalid input.
 * Canvas dimensions are capped at DIAGRAM_MAX_CANVAS_DIM_PX (scaled down
 * proportionally) so huge diagrams cannot allocate unbounded bitmaps.
 */
export function computeDiagramRasterSize(
  viewBoxWidth: number,
  viewBoxHeight: number,
  maxWidthPx: unknown,
): DiagramRasterSize | null {
  if (
    !Number.isFinite(viewBoxWidth) ||
    !Number.isFinite(viewBoxHeight) ||
    viewBoxWidth <= 0 ||
    viewBoxHeight <= 0
  ) {
    return null
  }
  const clampedMax = clampDiagramMaxWidth(maxWidthPx)
  const widthPx = Math.max(1, Math.min(Math.floor(viewBoxWidth), clampedMax))
  const heightPx = Math.max(1, Math.round((viewBoxHeight * widthPx) / viewBoxWidth))
  let canvasWidth = Math.floor(viewBoxWidth * DIAGRAM_RASTER_SCALE)
  let canvasHeight = Math.floor(viewBoxHeight * DIAGRAM_RASTER_SCALE)
  canvasWidth = Math.max(1, canvasWidth)
  canvasHeight = Math.max(1, canvasHeight)
  const peak = Math.max(canvasWidth, canvasHeight)
  if (peak > DIAGRAM_MAX_CANVAS_DIM_PX) {
    const factor = DIAGRAM_MAX_CANVAS_DIM_PX / peak
    canvasWidth = Math.max(1, Math.floor(canvasWidth * factor))
    canvasHeight = Math.max(1, Math.floor(canvasHeight * factor))
  }
  return { widthPx, heightPx, canvasWidth, canvasHeight }
}

/** Pin intrinsic size so <img> decodes at viewBox dimensions; strips width/height first */
export function pinSvgIntrinsicSize(svg: string, width: number, height: number): string {
  return svg.replace(/<svg\b[^>]*>/, (tag) => {
    const stripped = tag.replace(/\s(?:width|height)=(?:"[^"]*"|'[^']*')/g, '')
    return stripped.replace(/^<svg\b/, `<svg width="${width}" height="${height}"`)
  })
}

/**
 * Rasterize a rendered diagram for the docx export.
 * Returns null when it cannot be drawn (missing/invalid viewBox, invalid size,
 * undecodable SVG, or no 2d context); callers fall back to the diagram source
 * as code in that case.
 */
export async function diagramSvgToPng(svg: string, maxWidthPx: number): Promise<NewImage | null> {
  const parsed = parseSvgViewBox(svg)
  if (!parsed) return null
  const raster = computeDiagramRasterSize(parsed.width, parsed.height, maxWidthPx)
  if (!raster) return null

  // pin the intrinsic size (mermaid emits width="100%") so <img> decodes at the
  // viewBox dimensions instead of the 300x150 SVG default
  const sized = pinSvgIntrinsicSize(svg, parsed.width, parsed.height)
  const url = URL.createObjectURL(new Blob([sized], { type: 'image/svg+xml' }))
  try {
    const img = new Image()
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve()
      img.onerror = () => reject(new Error('svg decode failed'))
      img.src = url
    })
    const canvas = document.createElement('canvas')
    canvas.width = raster.canvasWidth
    canvas.height = raster.canvasHeight
    const ctx = canvas.getContext('2d')
    if (!ctx) return null
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height)
    const dataUrl = canvas.toDataURL('image/png')
    const base64 = dataUrl.slice(dataUrl.indexOf(',') + 1)
    if (!base64) return null
    return {
      base64,
      mime: 'image/png',
      widthPx: raster.widthPx,
      heightPx: raster.heightPx,
      align: 'center',
    }
  } catch {
    return null
  } finally {
    URL.revokeObjectURL(url)
  }
}
