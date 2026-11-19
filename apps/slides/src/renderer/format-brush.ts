/**
 * Format painter (renderer side) — extracts formatting from a RenderNode and converts it into
 * IPC call parameters.
 *
 * Format scope:
 *  - shape/text: fill, stroke, default text run (first run of first paragraph), paragraph alignment
 *  - picture: stroke
 *  - cross-type application takes the intersection
 *
 * This module has no side effects; App.tsx calls it.
 */
import type {
  ShapeRenderNode,
  PictureRenderNode,
  RenderNode,
  RenderFill,
} from '@genoffice/pptx-render'
import type { EditParagraph, EditRun } from '../shared/ipc'

/** Format container captured by the render layer (pure data, serializable) */
export interface BrushFormat {
  sourceType: 'shape' | 'text' | 'picture'
  /** solid fill color (#RRGGBB) or 'none'; undefined = no fill info */
  fillColor?: string
  /** stroke; undefined = no stroke info */
  stroke?: { color: string; widthPt: number }
  /** Default text run format (extracted from text.lines[0].runs[0]) */
  run?: {
    fontFamily?: string
    fontSize?: number // pt
    bold?: boolean
    italic?: boolean
    underline?: boolean
    color?: string
  }
  /** Paragraph alignment (first paragraph) */
  align?: 'left' | 'center' | 'right' | 'justify'
}

/** px→pt conversion (96dpi) */
function pxToPt(px: number): number {
  return Math.round(((px * 72) / 96) * 100) / 100
}

/** RenderFill → fillColor string or undefined */
function renderFillToColor(fill: RenderFill | undefined): string | undefined {
  if (!fill || fill.kind === 'none') return 'none'
  if (fill.kind === 'solid') return fill.color
  // gradient/image: not carried over (the format painter only handles solid colors)
  return undefined
}

/** Extract format from a ShapeRenderNode / text node */
export function extractBrushFormat(node: RenderNode): BrushFormat | null {
  if (node.type === 'shape' || node.type === 'text') {
    const shape = node as ShapeRenderNode
    const fmt: BrushFormat = { sourceType: node.type as 'shape' | 'text' }

    // fill
    const fc = renderFillToColor(shape.fill)
    if (fc !== undefined) fmt.fillColor = fc

    // stroke
    if (shape.stroke) {
      fmt.stroke = { color: shape.stroke.color, widthPt: shape.stroke.widthPt }
    }

    // text: first run of first line (para start)
    if (shape.text?.lines?.length) {
      const firstLine = shape.text.lines.find((l) => l.paraStart !== false) ?? shape.text.lines[0]
      const firstRun = firstLine?.runs?.find((r) => !r.isBullet) ?? firstLine?.runs?.[0]
      if (firstRun) {
        fmt.run = {
          fontFamily: firstRun.fontFamily,
          fontSize: pxToPt(firstRun.fontSizePx),
          bold: firstRun.bold,
          italic: firstRun.italic,
          underline: firstRun.underline,
          color: firstRun.color,
        }
      }
      // align from first line
      const firstParaLine =
        shape.text.lines.find((l) => l.paraStart !== false) ?? shape.text.lines[0]
      if (firstParaLine?.align) fmt.align = firstParaLine.align
    }

    return fmt
  }

  if (node.type === 'picture') {
    const pic = node as PictureRenderNode
    const fmt: BrushFormat = { sourceType: 'picture' }
    if (pic.stroke) {
      fmt.stroke = { color: pic.stroke.color, widthPt: pic.stroke.widthPt }
    }
    return fmt
  }

  return null
}

/** Format application result (caller invokes IPC as needed) */
export interface BrushApplyPlan {
  /** editFill parameters */
  fillColor?: string
  /** editStroke parameters */
  stroke?: { color: string; widthPt: number } | null
  /** editText parameters (replaces format only, text unchanged; needs to read the current node's text) */
  runFormat?: BrushFormat['run']
  /** Paragraph alignment */
  align?: 'left' | 'center' | 'right' | 'justify'
}

/**
 * Apply BrushFormat to the target node type, producing an IPC call plan.
 * Cross-type takes the intersection: picture only receives stroke.
 */
export function computeBrushApply(fmt: BrushFormat, targetNode: RenderNode): BrushApplyPlan {
  const plan: BrushApplyPlan = {}

  if (targetNode.type === 'picture') {
    if (fmt.stroke !== undefined) plan.stroke = fmt.stroke
    return plan
  }

  if (targetNode.type === 'shape' || targetNode.type === 'text') {
    if (fmt.fillColor !== undefined) plan.fillColor = fmt.fillColor
    if (fmt.stroke !== undefined) plan.stroke = fmt.stroke
    if (fmt.run) plan.runFormat = fmt.run
    if (fmt.align) plan.align = fmt.align
    return plan
  }

  return plan
}

/**
 * Convert BrushApplyPlan's runFormat + align into EditParagraph[], keeping the target node's
 * original text structure and replacing only run format attributes.
 *
 * targetNode is a ShapeRenderNode; original run text is extracted from its text.lines.
 */
export function buildEditParagraphsWithFormat(
  node: RenderNode,
  runFmt: BrushFormat['run'] | undefined,
  align: 'left' | 'center' | 'right' | 'justify' | undefined,
): EditParagraph[] | null {
  if (node.type !== 'shape' && node.type !== 'text') return null
  const shape = node as ShapeRenderNode
  const lines = shape.text?.lines ?? []
  if (lines.length === 0) return null

  // Group lines by paraStart → paragraphs
  const paraGroups: (typeof lines)[number][][] = []
  for (const line of lines) {
    if (line.paraStart !== false || paraGroups.length === 0) {
      paraGroups.push([line])
    } else {
      paraGroups[paraGroups.length - 1]!.push(line)
    }
  }

  return paraGroups.map((group) => {
    const firstLine = group[0]!
    const effectiveAlign = align ?? firstLine.align

    // Merge each line's runs into paragraph runs
    const runs: EditRun[] = []
    for (const line of group) {
      for (const r of line.runs) {
        if (r.isBullet) continue
        const run: EditRun = { text: r.text }
        if (runFmt) {
          if (runFmt.fontFamily) run.fontFamily = runFmt.fontFamily
          if (runFmt.fontSize) run.fontSize = runFmt.fontSize
          if (runFmt.bold != null) run.bold = runFmt.bold
          if (runFmt.italic != null) run.italic = runFmt.italic
          if (runFmt.underline != null) run.underline = runFmt.underline
          if (runFmt.color) run.color = runFmt.color
        } else {
          // Keep original format
          run.bold = r.bold
          run.italic = r.italic
          run.underline = r.underline
          run.fontFamily = r.fontFamily
          run.fontSize = pxToPt(r.fontSizePx)
          run.color = r.color
        }
        runs.push(run)
      }
    }

    if (runs.length === 0) runs.push({ text: '' })

    return {
      runs,
      ...(effectiveAlign ? { align: effectiveAlign } : {}),
    }
  })
}
