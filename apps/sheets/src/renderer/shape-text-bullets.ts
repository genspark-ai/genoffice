import type { CSSProperties } from 'react'
import { formatAutoNum } from '@genoffice/pptx-render/auto-num'
import type { WorkbookVisualObject } from '../shared/desktop-api'
import { shapePtLength } from './shape-text-scale'

export type ShapeParagraph = NonNullable<WorkbookVisualObject['paragraphs']>[number]

/// Marker text per paragraph. Auto-numbering continues across consecutive
/// paragraphs of one scheme/startAt and restarts after a paragraph without
/// it (PowerPoint and Excel share this single-level rule); empty paragraphs
/// show nothing and leave the count alone.
export function shapeParagraphMarkers(
  paragraphs: readonly ShapeParagraph[],
): Array<string | undefined> {
  let count = 0
  let scheme: string | undefined
  let start = 1
  return paragraphs.map((paragraph) => {
    if (!paragraph.runs.some((run) => run.text.trim() !== '')) return undefined
    if (paragraph.bulletScheme) {
      const startAt = paragraph.bulletStartAt ?? 1
      const running = count > 0 && scheme === paragraph.bulletScheme && start === startAt
      count = running ? count + 1 : startAt
      scheme = paragraph.bulletScheme
      start = startAt
      return formatAutoNum(count, paragraph.bulletScheme)
    }
    count = 0
    scheme = undefined
    return paragraph.bulletChar || undefined
  })
}

/// marL is the left edge of wrapped lines, indent the first-line offset from
/// it. A negative indent hangs the marker in the gap before the text (clamped
/// so the first line never starts left of the frame).
export function shapeParagraphIndentStyle(paragraph: ShapeParagraph): CSSProperties {
  const marginLeft = Math.max(0, paragraph.marginLeft ?? 0)
  const indent = Math.max(-marginLeft, paragraph.indent ?? 0)
  const style: CSSProperties = {}
  if (marginLeft) style.paddingLeft = shapePtLength(marginLeft)
  if (indent) style.textIndent = shapePtLength(indent)
  return style
}

/// Width of a hanging marker so the text after it starts at marL.
export function shapeBulletWidth(paragraph: ShapeParagraph): string | undefined {
  const marginLeft = Math.max(0, paragraph.marginLeft ?? 0)
  const indent = Math.max(-marginLeft, paragraph.indent ?? 0)
  return indent < 0 ? shapePtLength(-indent) : undefined
}
