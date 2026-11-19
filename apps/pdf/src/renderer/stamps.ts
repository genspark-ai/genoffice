import type { StampInput } from '../shared/ipc'

/** Bitmap supersampling factor relative to PDF pt — stays sharp even when enlarged for print */
const SS = 4

export interface WatermarkConfig {
  text: string
  /** Counterclockwise angle */
  angle: number
  opacity: number
  color: string
  /** Font size as a ratio of page width */
  sizeRatio: number
}

export interface HeaderFooterConfig {
  headerLeft: string
  headerCenter: string
  headerRight: string
  footerLeft: string
  footerCenter: string
  footerRight: string
  /** Auto page number in the footer center (overrides footerCenter) */
  pageNumber: boolean
  startAt: number
  fontSize: number
  color: string
}

export const DEFAULT_WATERMARK: WatermarkConfig = {
  text: '',
  angle: 35,
  opacity: 0.18,
  color: '#d0342c',
  sizeRatio: 0.11,
}

export const DEFAULT_HEADER_FOOTER: HeaderFooterConfig = {
  headerLeft: '',
  headerCenter: '',
  headerRight: '',
  footerLeft: '',
  footerCenter: '',
  footerRight: '',
  pageNumber: true,
  startAt: 1,
  fontSize: 9,
  color: '#666666',
}

const FONT = (px: number, bold = false) =>
  `${bold ? '600 ' : ''}${px}px -apple-system, "PingFang SC", "Microsoft YaHei", "Segoe UI", sans-serif`

function toBase64(canvas: HTMLCanvasElement): string {
  return canvas.toDataURL('image/png').split(',')[1] ?? ''
}

/**
 * Watermark bitmap: full-page transparent canvas with text rotated around the center.
 * Uses a bitmap because pdf-lib's built-in fonts lack CJK, and embedding fonts would
 * bundle several MB of font data.
 */
export function renderWatermark(cfg: WatermarkConfig, pw: number, ph: number): string | null {
  const text = cfg.text.trim()
  if (!text) return null
  const canvas = document.createElement('canvas')
  canvas.width = Math.round(pw * SS)
  canvas.height = Math.round(ph * SS)
  const ctx = canvas.getContext('2d')
  if (!ctx) return null
  ctx.translate(canvas.width / 2, canvas.height / 2)
  ctx.rotate((-cfg.angle * Math.PI) / 180)
  // Shrink long text to the diagonal so slanted ends aren't clipped by the page
  let size = pw * cfg.sizeRatio * SS
  ctx.font = FONT(size, true)
  const maxW = Math.hypot(canvas.width, canvas.height) * 0.8
  const w = ctx.measureText(text).width
  if (w > maxW) {
    size *= maxW / w
    ctx.font = FONT(size, true)
  }
  ctx.fillStyle = cfg.color
  ctx.textAlign = 'center'
  ctx.textBaseline = 'middle'
  ctx.fillText(text, 0, 0)
  return toBase64(canvas)
}

/** Header or footer bar: transparent canvas of full page width × bar height, in left/center/right segments */
function renderBar(
  parts: [string, string, string],
  pw: number,
  fontSize: number,
  color: string,
): string | null {
  if (parts.every((p) => !p.trim())) return null
  const h = Math.round(fontSize * 2.2)
  const canvas = document.createElement('canvas')
  canvas.width = Math.round(pw * SS)
  canvas.height = h * SS
  const ctx = canvas.getContext('2d')
  if (!ctx) return null
  ctx.scale(SS, SS)
  ctx.font = FONT(fontSize)
  ctx.fillStyle = color
  ctx.textBaseline = 'middle'
  const pad = pw * 0.06
  const y = h / 2
  const [left, center, right] = parts
  if (left.trim()) {
    ctx.textAlign = 'left'
    ctx.fillText(left, pad, y)
  }
  if (center.trim()) {
    ctx.textAlign = 'center'
    ctx.fillText(center, pw / 2, y)
  }
  if (right.trim()) {
    ctx.textAlign = 'right'
    ctx.fillText(right, pw - pad, y)
  }
  return toBase64(canvas)
}

/** Replace {page}/{total} placeholders */
const fill = (tpl: string, page: number, total: number) =>
  tpl.replaceAll('{page}', String(page)).replaceAll('{total}', String(total))

/**
 * Build stamps for each target page. pages provides (original page index, unrotated
 * page size, display number). Headers/footers are placed in unrotated coordinates and
 * follow rotation consistently with the page's /Rotate.
 */
export function buildStamps(
  pages: { origIdx: number; pw: number; ph: number; displayNo: number }[],
  watermark: WatermarkConfig | null,
  hf: HeaderFooterConfig | null,
): StampInput[] {
  const out: StampInput[] = []
  const total = pages.length
  // Pages of the same size share one watermark bitmap, avoiding re-rendering in large docs
  const wmCache = new Map<string, string | null>()

  for (const p of pages) {
    if (watermark?.text.trim()) {
      const key = `${Math.round(p.pw)}x${Math.round(p.ph)}`
      if (!wmCache.has(key)) wmCache.set(key, renderWatermark(watermark, p.pw, p.ph))
      const image = wmCache.get(key)
      if (image) {
        out.push({
          pageIndex: p.origIdx,
          image,
          rect: [0, 0, p.pw, p.ph],
          opacity: watermark.opacity,
        })
      }
    }

    if (!hf) continue
    const no = hf.startAt + p.displayNo - 1
    const barH = hf.fontSize * 2.2
    const margin = Math.min(p.ph * 0.035, 26)

    const header = renderBar(
      [hf.headerLeft, hf.headerCenter, hf.headerRight].map((s) => fill(s, no, total)) as [string, string, string],
      p.pw,
      hf.fontSize,
      hf.color,
    )
    if (header) {
      out.push({
        pageIndex: p.origIdx,
        image: header,
        rect: [0, p.ph - margin - barH, p.pw, p.ph - margin],
      })
    }

    const footerCenter = hf.pageNumber ? `${no} / ${total}` : fill(hf.footerCenter, no, total)
    const footer = renderBar(
      [fill(hf.footerLeft, no, total), footerCenter, fill(hf.footerRight, no, total)],
      p.pw,
      hf.fontSize,
      hf.color,
    )
    if (footer) {
      out.push({ pageIndex: p.origIdx, image: footer, rect: [0, margin, p.pw, margin + barH] })
    }
  }
  return out
}
