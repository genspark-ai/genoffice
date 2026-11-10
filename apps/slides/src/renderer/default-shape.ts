/**
 * "Set as Default Shape": the fill/line that new gallery shapes start with.
 * PowerPoint stores it per deck; ours is per install (localStorage).
 */
import type { RenderNode } from '@genoffice/pptx-render'

export interface DefaultShapeStyle {
  fillColor?: string
  stroke?: { color: string; widthPt: number }
}

const KEY = 'ai-slides-default-shape'
const FACTORY: DefaultShapeStyle = { fillColor: '#C43E1C' }

// #RRGGBB or #RRGGBBAA (the alpha byte is how the format pane stores transparency)
const hex = (c: string | undefined) =>
  c && /^#?[0-9a-fA-F]{6}(?:[0-9a-fA-F]{2})?$/.test(c)
    ? c.startsWith('#')
      ? c
      : `#${c}`
    : undefined

export function defaultShapeStyle(): DefaultShapeStyle {
  try {
    const raw = localStorage.getItem(KEY)
    if (raw) return JSON.parse(raw) as DefaultShapeStyle
  } catch {
    /* unreadable entry: fall back to the factory style */
  }
  return FACTORY
}

/** Seedable style of a shape node; undefined for connectors and shapes with neither solid fill nor line. */
export function shapeStyleOf(node: RenderNode | undefined): DefaultShapeStyle | undefined {
  if (node?.type !== 'shape' || node.line) return undefined
  const fillColor = node.fill.kind === 'solid' ? hex(node.fill.color) : undefined
  const strokeColor = hex(node.stroke?.color)
  const stroke = strokeColor ? { color: strokeColor, widthPt: node.stroke!.widthPt } : undefined
  if (!fillColor && !stroke) return undefined
  return { ...(fillColor ? { fillColor } : {}), ...(stroke ? { stroke } : {}) }
}

export function saveDefaultShapeStyle(style: DefaultShapeStyle): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(style))
  } catch {
    /* quota: the default is a convenience, not document data */
  }
}
