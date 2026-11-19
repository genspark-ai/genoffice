/**
 * SmartArt preset layout math, split from smartart.ts so the renderer can import
 * it for gallery previews without pulling in the archive/XML machinery (Node-only deps).
 */
import type { EmuRect } from './types'

export type SmartArtLayout =
  'list' | 'process' | 'cycle' | 'hierarchy' | 'pyramid' | 'matrix' | 'venn'

/** PowerPoint default theme accent series (colors picked in order) */
export const SMARTART_PALETTE = [
  '4472C4',
  'ED7D31',
  'A5A5A5',
  'FFC000',
  '5B9BD5',
  '70AD47',
  '264478',
  '9E480E',
]

export interface SmartArtChildShape {
  prst: string
  box: EmuRect
  text?: string
  color: string
  fontSize?: number
  /** Fill opacity 0..1 (venn overlaps); omitted = opaque */
  alpha?: number
}

/** Generate child shapes per layout (child coordinate space = the group's own resolution-independent coords, in EMU). */
export function layoutShapes(
  layout: SmartArtLayout,
  items: string[],
  cx: number,
  cy: number,
): SmartArtChildShape[] {
  const n = Math.min(Math.max(items.length, 1), 8)
  const texts = items.slice(0, n)
  const color = (i: number) => SMARTART_PALETTE[i % SMARTART_PALETTE.length]!
  const out: SmartArtChildShape[] = []

  if (layout === 'list') {
    const gap = Math.round(cy * 0.04)
    const h = Math.round((cy - gap * (n - 1)) / n)
    texts.forEach((t, i) => {
      out.push({
        prst: 'roundRect',
        box: { x: 0, y: i * (h + gap), cx, cy: h },
        text: t,
        color: color(i),
      })
    })
  } else if (layout === 'process') {
    // Horizontal chevron flow: adjacent arrows overlap by 1/4 to interlock
    const overlap = 0.25
    const w = Math.round(cx / (n - (n - 1) * overlap))
    const h = Math.round(cy * 0.6)
    const y = Math.round((cy - h) / 2)
    texts.forEach((t, i) => {
      out.push({
        prst: i === 0 ? 'homePlate' : 'chevron',
        box: { x: Math.round(i * w * (1 - overlap)), y, cx: w, cy: h },
        text: t,
        color: color(i),
        fontSize: 12,
      })
    })
  } else if (layout === 'cycle') {
    // n ellipses distributed around a circle
    const nodeW = Math.round(cx * 0.3)
    const nodeH = Math.round(cy * 0.26)
    const rx = (cx - nodeW) / 2
    const ry = (cy - nodeH) / 2
    texts.forEach((t, i) => {
      const ang = -Math.PI / 2 + (i * 2 * Math.PI) / n
      out.push({
        prst: 'ellipse',
        box: {
          x: Math.round(cx / 2 - nodeW / 2 + Math.cos(ang) * rx),
          y: Math.round(cy / 2 - nodeH / 2 + Math.sin(ang) * ry),
          cx: nodeW,
          cy: nodeH,
        },
        text: t,
        color: color(i),
        fontSize: 12,
      })
    })
  } else if (layout === 'pyramid') {
    // Stacked slices: triangle apex, trapezoid layers widening toward the base
    const h = Math.round(cy / n)
    texts.forEach((t, i) => {
      const w = Math.round((cx * (i + 1)) / n)
      out.push({
        prst: i === 0 ? 'triangle' : 'trapezoid',
        box: { x: Math.round((cx - w) / 2), y: i * h, cx: w, cy: h },
        text: t,
        color: color(i),
        fontSize: 12,
      })
    })
  } else if (layout === 'matrix') {
    // 2-column quadrant grid
    const gap = Math.round(Math.min(cx, cy) * 0.03)
    const rows = Math.ceil(n / 2)
    const w = Math.round((cx - gap) / 2)
    const h = Math.round((cy - gap * (rows - 1)) / rows)
    texts.forEach((t, i) => {
      out.push({
        prst: 'roundRect',
        box: { x: (i % 2) * (w + gap), y: Math.floor(i / 2) * (h + gap), cx: w, cy: h },
        text: t,
        color: color(i),
        fontSize: 12,
      })
    })
  } else if (layout === 'venn') {
    // 2..4 translucent circles overlapping around the center
    const m = Math.min(n, 4)
    const s = Math.min(cx, cy)
    const d = Math.round(s * 0.6)
    const ringR = Math.round(s * 0.2)
    const start = m === 2 ? 0 : -Math.PI / 2
    texts.slice(0, m).forEach((t, i) => {
      const ang = start + (i * 2 * Math.PI) / m
      out.push({
        prst: 'ellipse',
        box: {
          x: Math.round(cx / 2 - d / 2 + Math.cos(ang) * ringR),
          y: Math.round(cy / 2 - d / 2 + Math.sin(ang) * ringR),
          cx: d,
          cy: d,
        },
        text: t,
        color: color(i),
        fontSize: 12,
        alpha: 0.55,
      })
    })
  } else {
    // hierarchy: first item centered on top, the rest in the bottom row
    const top = texts[0]!
    const rest = texts.slice(1)
    const h = Math.round(cy * 0.32)
    const topW = Math.round(cx * 0.36)
    out.push({
      prst: 'roundRect',
      box: { x: Math.round((cx - topW) / 2), y: 0, cx: topW, cy: h },
      text: top,
      color: color(0),
    })
    if (rest.length) {
      const gap = Math.round(cx * 0.03)
      const w = Math.round((cx - gap * (rest.length - 1)) / rest.length)
      rest.forEach((t, i) => {
        // Vertical connector from top node to child (thin rect, avoids connector semantics)
        const nodeX = i * (w + gap)
        out.push({
          prst: 'rect',
          box: {
            x: Math.round(nodeX + w / 2 - cx * 0.002),
            y: h,
            cx: Math.max(Math.round(cx * 0.004), 9525),
            cy: Math.round(cy * 0.18),
          },
          color: 'A5A5A5',
        })
        out.push({
          prst: 'roundRect',
          box: { x: nodeX, y: Math.round(cy * 0.5), cx: w, cy: Math.round(cy * 0.5) },
          text: t,
          color: color(i + 1),
          fontSize: 12,
        })
      })
    }
  }
  return out
}
