/**
 * Draws the visible face of a certificate signature to a PNG. It is document content (like a
 * stamp), so its colors are fixed and never follow the UI theme; canvas text lets any script
 * (CJK names included) render without embedding a font into the PDF.
 */
const INK = '#1f2937'
const INK_DIM = '#4b5563'
const FRAME = '#1d4e89'
const PAPER = 'rgba(255, 255, 255, 0.92)'
const FONT_STACK =
  "'Noto Sans CJK SC', 'PingFang SC', 'Microsoft YaHei', 'Segoe UI', Arial, sans-serif"

export interface AppearanceLines {
  /** "Digitally signed by" */
  heading: string
  name: string
  date: string
  reason?: string
  location?: string
}

/** Pixels per PDF point in the rendered PNG: sharp when zoomed, still small as a file */
const SCALE = 3

function fit(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string {
  if (ctx.measureText(text).width <= maxWidth) return text
  let end = text.length
  while (end > 1 && ctx.measureText(`${text.slice(0, end)}…`).width > maxWidth) end--
  return `${text.slice(0, end)}…`
}

/** Base64 PNG (no data: prefix) of a `widthPt` x `heightPt` signature box */
export function renderSignatureAppearance(
  lines: AppearanceLines,
  widthPt: number,
  heightPt: number,
): string {
  const canvas = document.createElement('canvas')
  canvas.width = Math.round(widthPt * SCALE)
  canvas.height = Math.round(heightPt * SCALE)
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('canvas unavailable')
  ctx.scale(SCALE, SCALE)

  ctx.fillStyle = PAPER
  ctx.fillRect(0, 0, widthPt, heightPt)
  ctx.strokeStyle = FRAME
  ctx.lineWidth = 1
  ctx.strokeRect(0.5, 0.5, widthPt - 1, heightPt - 1)

  const pad = 6
  const maxWidth = widthPt - pad * 2
  ctx.textBaseline = 'top'
  let y = pad - 1

  ctx.fillStyle = INK_DIM
  ctx.font = `8px ${FONT_STACK}`
  ctx.fillText(fit(ctx, lines.heading, maxWidth), pad, y)
  y += 11

  ctx.fillStyle = INK
  let size = 14
  ctx.font = `600 ${size}px ${FONT_STACK}`
  while (size > 8 && ctx.measureText(lines.name).width > maxWidth) {
    size--
    ctx.font = `600 ${size}px ${FONT_STACK}`
  }
  ctx.fillText(fit(ctx, lines.name, maxWidth), pad, y)
  y += size + 4

  ctx.fillStyle = INK_DIM
  ctx.font = `8px ${FONT_STACK}`
  for (const line of [lines.date, lines.reason, lines.location]) {
    if (!line || y + 9 > heightPt - pad + 2) continue
    ctx.fillText(fit(ctx, line, maxWidth), pad, y)
    y += 10
  }
  return canvas.toDataURL('image/png').split(',')[1] ?? ''
}
