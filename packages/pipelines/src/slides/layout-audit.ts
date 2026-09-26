import type {
  GroupRenderNode,
  PictureRenderNode,
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
  rotationDeg: number
  srcRect?: { l: number; t: number; r: number; b: number }
  /** audio/video poster frame: sized to the media, not to the bitmap */
  poster: boolean
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
        // Vertical text (vert/vert270/eaVert/wordArtVert) keeps horizontal layout
        // convention in the renderer's lines, so run.x/widthPx mix axes — the width
        // check would false-positive on every vertical box; skip it (the height
        // check above still applies via contentHeight, which the renderer computes
        // correctly for both orientations).
        if (sn.text.vert) {
          overflowXPx = 0
        } else {
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
      }
    } else if (n.type === 'group') {
      // If any child in the group has text, treat it as text content for overlap detection
      hasText = groupHasText(n as GroupRenderNode)
      preview = '(group)'
    }
    const pic = n.type === 'picture' ? (n as PictureRenderNode) : undefined
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
      rotationDeg: n.box.rotationDeg,
      ...(pic?.srcRect ? { srcRect: pic.srcRect } : {}),
      poster: Boolean(pic?.media),
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

function label(e: AuditEntry, idOf: (sourceId: string) => string): string {
  const id = idOf(e.id)
  return e.preview && e.preview !== '(group)' ? `${id}"${e.preview}"` : `${id}(${e.type})`
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
/** slack a suggested box adds beyond the measured need */
const SUGGEST_SLACK_PX = 4
const EMU_PER_PX = 9525
/** box aspect may differ from the (cropped) source aspect by this fraction before a picture counts as stretched */
const DISTORTION_TOLERANCE = 0.05
/** a picture shrunk to its true aspect below this side length grows instead */
const MIN_PICTURE_SIDE_PX = 24

export type AuditCode =
  | 'out_of_bounds'
  | 'off_slide'
  | 'text_overflow'
  | 'text_overflow_width'
  | 'overlap'
  | 'picture_distorted'

export interface PictureSize {
  w: number
  h: number
}

export interface AuditOptions {
  /** Natural pixel size of a top-level picture's bitmap by node source id; undefined skips the distortion check. */
  pictureSize?: (sourceId: string) => PictureSize | undefined
}

export interface AuditBox {
  x: number
  y: number
  w: number
  h: number
}

/** A `setTransform` an agent can apply as is (EMU); `target.slide` is added by the deck-level caller. */
export interface AuditSuggest {
  op: 'setTransform'
  target: { slide?: number | string; el: string }
  box: { x: number; y: number; cx: number; cy: number }
  rotDeg?: number
}

export interface AuditFinding {
  code: AuditCode
  level: 'error' | 'warning'
  /** id of the element (idOf applied); overlap: the first of the pair */
  el: string
  els?: string[]
  message: string
  /** element box in slide px */
  box: AuditBox
  /** px the text exceeds the box (text_overflow: height, text_overflow_width: width) */
  overflowPx?: number
  /** picture_distorted: cropped source aspect (w/h), the box aspect and how far apart they are in percent */
  expected_ratio?: number
  actual_ratio?: number
  distortion_pct?: number
  suggest?: AuditSuggest
}

interface Distortion {
  expected: number
  actual: number
  pct: number
  /** the size that restores the source aspect (the longer side shrunk unless that leaves it tiny) */
  w: number
  h: number
}

function distortionOf(e: AuditEntry, size: PictureSize): Distortion | undefined {
  if (e.w <= 0 || e.h <= 0) return undefined
  const sr = e.srcRect ?? { l: 0, t: 0, r: 0, b: 0 }
  const srcW = size.w * (1 - sr.l - sr.r)
  const srcH = size.h * (1 - sr.t - sr.b)
  if (srcW <= 0 || srcH <= 0) return undefined
  const expected = srcW / srcH
  const actual = e.w / e.h
  const pct = Math.abs(actual / expected - 1) * 100
  if (pct <= DISTORTION_TOLERANCE * 100) return undefined
  const keepW = { w: e.w, h: e.w / expected }
  const keepH = { w: e.h * expected, h: e.h }
  // Shrinking keeps the picture inside its placeholder (a square logo in a banner stays
  // banner-high); growing only when the shrunk result would be too small to read.
  const shrink = actual > expected ? keepH : keepW
  const grow = shrink === keepH ? keepW : keepH
  const fit = Math.min(shrink.w, shrink.h) < MIN_PICTURE_SIDE_PX ? grow : shrink
  return { expected, actual, pct, ...fit }
}

/**
 * Audit one page's layout and return the list of problems (empty array = pass).
 * `idOf` maps a node's source id to the id the caller's tools accept (durable ids in the CLI).
 */
export function auditSlideLayout(
  slide: RenderSlide,
  idOf: (sourceId: string) => string = (id) => id,
  opts: AuditOptions = {},
): string[] {
  return auditSlideFindings(slide, idOf, opts).map((f) => f.message)
}

/** The same audit with each finding typed, located and, where geometry alone fixes it, carrying the op. */
export function auditSlideFindings(
  slide: RenderSlide,
  idOf: (sourceId: string) => string = (id) => id,
  opts: AuditOptions = {},
): AuditFinding[] {
  const entries = collectEntries(slide.nodes)
  const findings: AuditFinding[] = []
  const W = slide.widthPx
  const H = slide.heightPx
  const boxOf = (e: AuditEntry): AuditBox => ({ x: e.x, y: e.y, w: e.w, h: e.h })
  // Strict comparisons: a box touching the canvas edge from outside still reads as out_of_bounds.
  const offSlide = (e: AuditEntry): boolean => e.x + e.w < 0 || e.y + e.h < 0 || e.x > W || e.y > H
  const distortions = new Map<AuditEntry, Distortion>()
  if (opts.pictureSize) {
    for (const e of entries) {
      if (e.type !== 'picture' || e.poster) continue
      const size = opts.pictureSize(e.id)
      const d = size && distortionOf(e, size)
      if (d) distortions.set(e, d)
    }
  }
  const outside = (e: AuditEntry, tolerance = EDGE_TOLERANCE_PX): string[] => {
    const parts: string[] = []
    if (e.x < -tolerance) parts.push(`${Math.round(-e.x)}px past the left edge`)
    if (e.y < -tolerance) parts.push(`${Math.round(-e.y)}px past the top edge`)
    if (e.x + e.w > W + tolerance) parts.push(`${Math.round(e.x + e.w - W)}px past the right edge`)
    if (e.y + e.h > H + tolerance) parts.push(`${Math.round(e.y + e.h - H)}px past the bottom edge`)
    return parts
  }
  // One box per element that answers every finding on it at once (grown for the text,
  // then brought inside the canvas), so applying the suggestions in any order converges.
  interface Plan {
    suggest?: AuditSuggest
    grewW: boolean
    grewH: boolean
  }
  const plans = new Map<AuditEntry, Plan>()
  const planFor = (e: AuditEntry): Plan => {
    const cached = plans.get(e)
    if (cached) return cached
    // A grown axis that would not fit the slide is left alone: the message then asks for a
    // smaller font instead of a suggestion that cannot clear the finding.
    const wantW =
      e.w + (e.overflowXPx > OVERFLOW_TOLERANCE_PX ? e.overflowXPx + SUGGEST_SLACK_PX : 0)
    const wantH = e.h + (e.overflowPx > OVERFLOW_TOLERANCE_PX ? e.overflowPx + SUGGEST_SLACK_PX : 0)
    const grewW = wantW > e.w && wantW <= W
    const grewH = wantH > e.h && wantH <= H
    let w = Math.min(grewW ? wantW : e.w, W)
    let h = Math.min(grewH ? wantH : e.h, H)
    let x = e.x
    let y = e.y
    const fit = distortions.get(e)
    if (fit) {
      // a corrected picture larger than the canvas scales down uniformly so the aspect still holds
      const scale = Math.min(1, W / fit.w, H / fit.h)
      w = fit.w * scale
      h = fit.h * scale
      x = e.x + (e.w - w) / 2
      y = e.y + (e.h - h) / 2
    }
    const box = { x: clamp(x, 0, W - w), y: clamp(y, 0, H - h), w, h }
    const same = box.x === e.x && box.y === e.y && box.w === e.w && box.h === e.h
    const plan: Plan = {
      grewW,
      grewH,
      ...(same
        ? {}
        : {
            suggest: {
              op: 'setTransform' as const,
              target: { el: idOf(e.id) },
              box: {
                x: Math.round(box.x * EMU_PER_PX),
                y: Math.round(box.y * EMU_PER_PX),
                cx: Math.round(box.w * EMU_PER_PX),
                cy: Math.round(box.h * EMU_PER_PX),
              },
              ...(e.rotationDeg ? { rotDeg: e.rotationDeg } : {}),
            },
          }),
    }
    plans.set(e, plan)
    return plan
  }
  const withSuggest = (e: AuditEntry, when: (plan: Plan) => boolean = () => true) => {
    const plan = planFor(e)
    return plan.suggest && when(plan) ? { suggest: plan.suggest } : {}
  }

  // 1. Off the slide entirely (whatever the overhang), else out of bounds past the tolerance
  for (const e of entries) {
    const off = offSlide(e)
    const parts = off ? outside(e, 0) : outside(e)
    if (!parts.length) continue
    findings.push({
      code: off ? 'off_slide' : 'out_of_bounds',
      level: 'error',
      el: idOf(e.id),
      message: off
        ? `Off slide: ${label(e, idOf)} lies entirely outside the slide (${parts.join(', ')})`
        : `Out of bounds: ${label(e, idOf)} ${parts.join(', ')}`,
      box: boxOf(e),
      ...withSuggest(e),
    })
  }

  // 2. Text overflow
  for (const e of entries) {
    if (e.overflowPx > OVERFLOW_TOLERANCE_PX) {
      findings.push({
        code: 'text_overflow',
        level: 'error',
        el: idOf(e.id),
        message: `Text overflow: ${label(e, idOf)} content exceeds the box height by ${e.overflowPx}px (make the box taller or reduce the font size)`,
        box: boxOf(e),
        overflowPx: e.overflowPx,
        ...withSuggest(e, (plan) => plan.grewH),
      })
    }
  }

  // 2b. Horizontal overflow: a line wider than the box (wrap=false, or a single token wider
  // than the available width) spills over the box and overlaps neighbors. PowerPoint renders
  // nowrap overflow as-is, so this is an audit-only signal that lets the AI widen the box,
  // shrink the font, or enable wrapping instead of leaving invisible overlap.
  for (const e of entries) {
    if (e.overflowXPx > OVERFLOW_TOLERANCE_PX) {
      findings.push({
        code: 'text_overflow_width',
        level: 'warning',
        el: idOf(e.id),
        message: `Text overflow (width): ${label(e, idOf)} content exceeds the box width by ${e.overflowXPx}px (widen the box, reduce the font size, or turn on wrapping)`,
        box: boxOf(e),
        overflowPx: e.overflowXPx,
        ...withSuggest(e, (plan) => plan.grewW),
      })
    }
  }

  // 2c. Stretched pictures: the box aspect strays from the (cropped) bitmap aspect
  for (const [e, d] of distortions) {
    findings.push({
      code: 'picture_distorted',
      level: 'warning',
      el: idOf(e.id),
      message: `Picture distorted: ${label(e, idOf)} box aspect ${d.actual.toFixed(3)} vs source ${d.expected.toFixed(3)} (${d.pct.toFixed(1)}% off; resize one side or crop)`,
      box: boxOf(e),
      expected_ratio: round3(d.expected),
      actual_ratio: round3(d.actual),
      distortion_pct: Math.round(d.pct * 10) / 10,
      ...withSuggest(e),
    })
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
      findings.push({
        code: 'overlap',
        level: 'warning',
        el: idOf(a.id),
        els: [idOf(a.id), idOf(b.id)],
        message: `Overlap: ${label(a, idOf)} and ${label(b, idOf)} intersect by ${Math.round(ix)}×${Math.round(iy)}px`,
        box: boxOf(a),
      })
      if (findings.length >= MAX_ISSUES) break
    }
    if (findings.length >= MAX_ISSUES) break
  }

  return findings.slice(0, MAX_ISSUES)
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value))
}

function round3(value: number): number {
  return Math.round(value * 1000) / 1000
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
