import type { SignatureData } from './SignatureDialog'

export type StaticFormFillKind = 'text' | 'check' | 'cross'

const SS = 4
const FONT_FAMILY = '-apple-system, "PingFang SC", "Microsoft YaHei", "Segoe UI", sans-serif'

function canvasImage(
  width: number,
  height: number,
  paint: (ctx: CanvasRenderingContext2D) => void,
): Extract<SignatureData, { kind: 'image' }> | null {
  const canvas = document.createElement('canvas')
  canvas.width = Math.ceil(width * SS)
  canvas.height = Math.ceil(height * SS)
  const ctx = canvas.getContext('2d')
  if (!ctx) return null
  ctx.scale(SS, SS)
  paint(ctx)
  return {
    kind: 'image',
    image: canvas.toDataURL('image/png').split(',')[1] ?? '',
    width,
    height,
  }
}

/** Rasterized, transparent typewriter text; bitmap output keeps CJK portable. */
export function renderStaticFormText(
  source: string,
  fontSize = 14,
  color = '#111111',
  align: CanvasTextAlign = 'left',
): Extract<SignatureData, { kind: 'image' }> | null {
  const text = source.trim()
  if (!text) return null
  const lines = text.split(/\r?\n/)
  const measure = document.createElement('canvas').getContext('2d')
  if (!measure) return null
  measure.font = `${fontSize}px ${FONT_FAMILY}`
  const pad = Math.max(3, fontSize * 0.25)
  const lineHeight = fontSize * 1.25
  const contentWidth = Math.max(...lines.map((line) => measure.measureText(line || ' ').width))
  const width = contentWidth + pad * 2
  const height = lineHeight * lines.length + pad * 2
  return canvasImage(width, height, (ctx) => {
    ctx.font = `${fontSize}px ${FONT_FAMILY}`
    ctx.fillStyle = color
    ctx.textAlign = align
    ctx.textBaseline = 'alphabetic'
    const x = align === 'center' ? width / 2 : align === 'right' ? width - pad : pad
    lines.forEach((line, index) => {
      ctx.fillText(line, x, pad + fontSize + index * lineHeight)
    })
  })
}

/** Vector-painted check/cross bitmap, inserted through the existing movable image pipeline. */
export function renderStaticFormMark(
  kind: Exclude<StaticFormFillKind, 'text'>,
  size = 22,
  color = '#111111',
): Extract<SignatureData, { kind: 'image' }> | null {
  return canvasImage(size, size, (ctx) => {
    ctx.strokeStyle = color
    ctx.lineWidth = Math.max(1.8, size * 0.11)
    ctx.lineCap = 'round'
    ctx.lineJoin = 'round'
    ctx.beginPath()
    if (kind === 'check') {
      ctx.moveTo(size * 0.14, size * 0.53)
      ctx.lineTo(size * 0.4, size * 0.78)
      ctx.lineTo(size * 0.87, size * 0.2)
    } else {
      ctx.moveTo(size * 0.2, size * 0.2)
      ctx.lineTo(size * 0.8, size * 0.8)
      ctx.moveTo(size * 0.8, size * 0.2)
      ctx.lineTo(size * 0.2, size * 0.8)
    }
    ctx.stroke()
  })
}
