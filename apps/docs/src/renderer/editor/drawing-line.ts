// Modified for ForaOffice: render authored line geometry independently of its hit area.
import type { TextboxDisplay } from '@genoffice/docx-engine'

const n = (value: number) => Math.round(value * 100) / 100
const factor = (value: string) => (value === 'sm' ? 2 : value === 'lg' ? 5 : 3)

export function drawingLineSize(box: TextboxDisplay): {
  w: number
  h: number
  dx: number
  dy: number
} {
  const w = box.widthPx ?? 0
  const h = box.heightPx ?? 0
  const stroke = box.borderWidthPx ?? 1
  const ends = [box.lineStyle?.headEnd, box.lineStyle?.tailEnd]
  const pad = Math.max(
    4,
    stroke / 2,
    ...ends.map((end) =>
      end ? (Math.max(factor(end.width), factor(end.length)) * Math.max(stroke, 2.5)) / 2 : 0,
    ),
  )
  return {
    w: Math.max(8, w + 2 * pad),
    h: Math.max(8, h + 2 * pad),
    dx: Math.max(pad, (8 - w) / 2),
    dy: Math.max(pad, (8 - h) / 2),
  }
}

/** DrawingML head/tail kinds, compound strokes and dash presets on a straight segment. */
export function drawingLineBackground(box: TextboxDisplay): string {
  const size = drawingLineSize(box)
  const w = box.widthPx ?? 0
  const h = box.lineDiag ? (box.heightPx ?? 0) : 0
  // The legacy display preset reverses head-only arrows; authored ends retain
  // their original start/end semantics, so undo that display normalization.
  const reversed = !!box.lineStyle?.headEnd && !box.lineStyle?.tailEnd
  const flipH = !!box.flipH !== reversed
  const flipV = !!box.flipV !== reversed
  const x1 = size.dx + (flipH ? w : 0)
  const y1 = size.dy + (flipV ? h : 0)
  const x2 = size.dx + (flipH ? 0 : w)
  const y2 = size.dy + (flipV ? 0 : h)
  const stroke = `#${box.borderColor ?? '000000'}`
  const sw = box.borderWidthPx ?? 1
  const style = box.lineStyle
  const marker = (id: string, end: NonNullable<TextboxDisplay['lineStyle']>['headEnd']): string => {
    if (!end) return ''
    const mw = factor(end.length) * Math.max(sw, 2.5)
    const mh = factor(end.width) * Math.max(sw, 2.5)
    const path =
      end.type === 'triangle'
        ? 'M0 0 L1 .5 L0 1 Z'
        : end.type === 'stealth'
          ? 'M0 0 L1 .5 L0 1 L.25 .5 Z'
          : end.type === 'diamond'
            ? 'M0 .5 L.5 0 L1 .5 L.5 1 Z'
            : end.type === 'arrow'
              ? 'M0 0 L1 .5 L0 1'
              : null
    const body =
      end.type === 'oval'
        ? `<ellipse cx=".5" cy=".5" rx=".5" ry=".5" fill="${stroke}"/>`
        : path
          ? `<path d="${path}" fill="${end.type === 'arrow' ? 'none' : stroke}" stroke="${stroke}" stroke-width="${end.type === 'arrow' ? '.12' : '0'}"/>`
          : ''
    return `<marker id="${id}" viewBox="0 0 1 1" preserveAspectRatio="none" markerUnits="userSpaceOnUse" markerWidth="${n(mw)}" markerHeight="${n(mh)}" refX="1" refY=".5" orient="auto-start-reverse" overflow="visible">${body}</marker>`
  }
  const presets: Record<string, number[]> = {
    dot: [1, 3],
    sysDot: [1, 1],
    dash: [4, 3],
    sysDash: [3, 1],
    lgDash: [8, 3],
    dashDot: [4, 3, 1, 3],
    sysDashDot: [3, 1, 1, 1],
    lgDashDot: [8, 3, 1, 3],
    lgDashDotDot: [8, 3, 1, 3, 1, 3],
    sysDashDotDot: [3, 1, 1, 1, 1, 1],
  }
  const dash = presets[style?.dash ?? '']
  const paint = `stroke="${stroke}" fill="none"${dash ? ` stroke-dasharray="${dash.map((v) => n(v * sw)).join(' ')}"` : ''}`
  const len = Math.hypot(x2 - x1, y2 - y1) || 1
  const nx = -(y2 - y1) / len
  const ny = (x2 - x1) / len
  const bands: Record<string, Array<[number, number]>> = {
    dbl: [
      [-1 / 3, 1 / 3],
      [1 / 3, 1 / 3],
    ],
    thickThin: [
      [-1 / 4, 1 / 2],
      [5 / 12, 1 / 6],
    ],
    thinThick: [
      [-5 / 12, 1 / 6],
      [1 / 4, 1 / 2],
    ],
    tri: [
      [-5 / 12, 1 / 6],
      [0, 1 / 3],
      [5 / 12, 1 / 6],
    ],
  }
  const strokes = (bands[style?.compound ?? ''] ?? [[0, 1]])
    .map(
      ([offset, width]) =>
        `<path d="M ${n(x1 + nx * offset * sw)} ${n(y1 + ny * offset * sw)} L ${n(x2 + nx * offset * sw)} ${n(y2 + ny * offset * sw)}" ${paint} stroke-width="${n(width * sw)}"/>`,
    )
    .join('')
  const ends =
    style?.headEnd || style?.tailEnd
      ? `<path d="M ${n(x1)} ${n(y1)} L ${n(x2)} ${n(y2)}" stroke="none" fill="none" stroke-width="${sw}"${style.headEnd ? ' marker-start="url(#head)"' : ''}${style.tailEnd ? ' marker-end="url(#tail)"' : ''}/>`
      : ''
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size.w}" height="${size.h}" viewBox="0 0 ${size.w} ${size.h}" overflow="visible"><defs>${marker('head', style?.headEnd)}${marker('tail', style?.tailEnd)}</defs>${strokes}${ends}</svg>`
  return `background-image:url("data:image/svg+xml,${encodeURIComponent(svg)}");background-size:100% 100%;background-repeat:no-repeat`
}
