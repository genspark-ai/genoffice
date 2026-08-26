import type {
  GroupRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
} from '@genoffice/pptx-render'

/**
 * Deterministic layout audit (modeled on the Google Slides add-in review_google_slides_addin geometry-only checks):
 * pure geometric computation, no LLM calls, no screenshots. Checks three kinds of problems:
 *  1. Elements extending past the canvas
 *  2. Pairwise overlap of content elements (text-text / text-image/media)
 *  3. Text overflowing its text box (uses the render layer's already-laid-out text.contentHeight — exact, not estimated)
 * Results are appended to layout tools' return values so the AI "sees" the real post-edit state (write → verify → fix loop).
 */

interface AuditEntry {
  id: string
  type: string
  x: number
  y: number
  w: number
  h: number
  hasText: boolean
  preview: string
  /** Pixels by which the text content height exceeds the box height (only meaningful when >0) */
  overflowPx: number
  /** Pixels by which the widest laid-out line exceeds the box inner width (wrap=false lines, over-wide tokens) */
  overflowXPx: number
}

const PREVIEW_MAX = 18

function textPreview(node: ShapeRenderNode): string {
  const t = (node.text?.lines ?? [])
    .map((l) => l.runs.map((r) => r.text).join(''))
    .join(' ')
    .trim()
  return t.length > PREVIEW_MAX ? `${t.slice(0, PREVIEW_MAX)}…` : t
}

/** Collect the top-level nodes that take part in the audit (skip master/layout decoration; a group counts as one box). */
function collectEntries(nodes: RenderNode[]): AuditEntry[] {
  const out: AuditEntry[] = []
  for (const n of nodes) {
    if (n.decoration) continue
    const { x, y, w, h } = n.box
    let hasText = false
    let preview = ''
    let overflowPx = 0
    let overflowXPx = 0
    if (n.type === 'shape' || n.type === 'text') {
      const sn = n as ShapeRenderNode
      preview = textPreview(sn)
      hasText = preview.length > 0
      if (sn.text && hasText) {
        const inner = h - sn.text.insets.t - sn.text.insets.b
        overflowPx = Math.round(sn.text.contentHeight - inner)
        const innerW = w - sn.text.insets.l - sn.text.insets.r
        // Laid-out right edge of the widest line (run.x already includes marL/indent/alignment,
        // so the content-area comparison below is exact). Catches wrap=false lines and
        // unbreakable tokens wider than the box -- the height-only check misses both, and
        // the text then overlaps whatever sits next to it.
        const widest = sn.text.lines.reduce(
          (acc, ln) =>
            Math.max(
              acc,
              ln.runs.reduce((m, r) => Math.max(m, r.x + r.widthPx), -Infinity),
            ),
          -Infinity,
        )
        overflowXPx = Math.round(Number.isFinite(widest) ? widest - innerW : 0)
      }
    } else if (n.type === 'group') {
      // If any child in the group has text, treat it as text content for overlap detection
      hasText = groupHasText(n as GroupRenderNode)
      preview = '(group)'
    }
    out.push({
      id: n.sourceId,
      type: n.type,
      x,
      y,
      w,
      h,
      hasText,
      preview,
      overflowPx,
      overflowXPx,
    })
  }
  return out
}

function groupHasText(g: GroupRenderNode): boolean {
  for (const c of g.children) {
    if (c.type === 'group') {
      if (groupHasText(c as GroupRenderNode)) return true
    } else if (c.type === 'shape' || c.type === 'text') {
      const t = (c as ShapeRenderNode).text?.lines ?? []
      if (t.some((l) => l.runs.some((r) => r.text.trim()))) return true
    }
  }
  return false
}

const MEDIA_TYPES = new Set(['picture', 'table', 'chart', 'placeholder-chip'])

/** Whether it's a content element (participates in overlap detection): has text, or is a picture/table/chart. */
function isContent(e: AuditEntry): boolean {
  return e.hasText || MEDIA_TYPES.has(e.type)
}

function label(e: AuditEntry): string {
  return e.preview && e.preview !== '(group)' ? `${e.id}"${e.preview}"` : `${e.id}(${e.type})`
}

const EDGE_TOLERANCE_PX = 8
const OVERFLOW_TOLERANCE_PX = 4
/** Threshold for overlap area as a fraction of the smaller element's area */
const OVERLAP_RATIO = 0.12
/** Absolute overlap area floor (px²), filtering out noise like touching trims */
const OVERLAP_MIN_AREA = 400
/** Background color blocks (≥70% of canvas area) don't participate in overlap detection */
const BACKGROUND_AREA_RATIO = 0.7
const MAX_ISSUES = 12

/**
 * Audit one page's layout and return the list of problems (empty array = pass).
 */
export function auditSlideLayout(slide: RenderSlide): string[] {
  const entries = collectEntries(slide.nodes)
  const issues: string[] = []
  const W = slide.widthPx
  const H = slide.heightPx

  // 1. Out of bounds
  for (const e of entries) {
    const parts: string[] = []
    if (e.x < -EDGE_TOLERANCE_PX) parts.push(`${Math.round(-e.x)}px past the left edge`)
    if (e.y < -EDGE_TOLERANCE_PX) parts.push(`${Math.round(-e.y)}px past the top edge`)
    if (e.x + e.w > W + EDGE_TOLERANCE_PX)
      parts.push(`${Math.round(e.x + e.w - W)}px past the right edge`)
    if (e.y + e.h > H + EDGE_TOLERANCE_PX)
      parts.push(`${Math.round(e.y + e.h - H)}px past the bottom edge`)
    if (parts.length) issues.push(`Out of bounds: ${label(e)} ${parts.join(', ')}`)
  }

  // 2. Text overflow
  for (const e of entries) {
    if (e.overflowPx > OVERFLOW_TOLERANCE_PX) {
      issues.push(
        `Text overflow: ${label(e)} content exceeds the box height by ${e.overflowPx}px (make the box taller or reduce the font size)`,
      )
    }
  }

  // 2b. Horizontal overflow: a line wider than the box (wrap=false, or a single token wider
  // than the available width) spills over the box and overlaps neighbors. PowerPoint renders
  // nowrap overflow as-is, so this is an audit-only signal that lets the AI widen the box,
  // shrink the font, or enable wrapping instead of leaving invisible overlap.
  for (const e of entries) {
    if (e.overflowXPx > OVERFLOW_TOLERANCE_PX) {
      issues.push(
        `Text overflow (width): ${label(e)} content exceeds the box width by ${e.overflowXPx}px (widen the box, reduce the font size, or turn on wrapping)`,
      )
    }
  }

  // 3. Pairwise overlap of content elements
  const content = entries.filter((e) => isContent(e) && e.w * e.h < W * H * BACKGROUND_AREA_RATIO)
  for (let i = 0; i < content.length; i++) {
    for (let j = i + 1; j < content.length; j++) {
      const a = content[i]!
      const b = content[j]!
      // Only report text<->text and text<->media; media-on-media (e.g. a chart on an image) is often intentional design, don't report
      if (!a.hasText && !b.hasText) continue
      const ix = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)
      const iy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y)
      if (ix <= 0 || iy <= 0) continue
      const inter = ix * iy
      const minArea = Math.min(a.w * a.h, b.w * b.h)
      if (inter < OVERLAP_MIN_AREA || inter < minArea * OVERLAP_RATIO) continue
      issues.push(
        `Overlap: ${label(a)} and ${label(b)} intersect by ${Math.round(ix)}×${Math.round(iy)}px`,
      )
      if (issues.length >= MAX_ISSUES) break
    }
    if (issues.length >= MAX_ISSUES) break
  }

  return issues.slice(0, MAX_ISSUES)
}

/** Format the audit result as trailing text for a tool's return value. */
export function formatAudit(issues: string[], round?: string): string {
  if (issues.length === 0)
    return '\n<layout-audit>✅ Passed: no overlap/out-of-bounds/text overflow.</layout-audit>'
  const head = `\n<layout-audit>⚠️ Found ${issues.length} issue(s):\n`
  const body = issues.map((s) => `- ${s}`).join('\n')
  const tail = round
    ? `\n${round}\n</layout-audit>`
    : "\n→ Immediately write another execute_slide_script to fix these issues (don't stop, don't ask the user, don't declare completion). els reflects the new positions after the last apply; compute from it directly. At most 2 fix rounds; only if still unresolved tell the user honestly.\n</layout-audit>"
  return head + body + tail
}
