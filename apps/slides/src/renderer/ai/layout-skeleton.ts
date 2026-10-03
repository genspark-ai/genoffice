/**
 * Layout skeleton: the recurring "chrome" geometry of a deck (title box, brand
 * image slot, accent shapes, background), extracted deterministically from the
 * current slides and stored inside a style template (#1821).
 *
 * A saved style template used to carry only the Style Skill text — colors and
 * fonts — so pages generated from it reinvented the layout every time and brand
 * elements (logo, header band) drifted page to page. The skeleton pins those
 * recurring boxes: save time extracts them from the deck, generation time hands
 * them to every page request as a mandatory block.
 *
 * Pure and Node-testable: input is the render tree (absolute px boxes), output
 * is plain JSON (what ai:save-style-template stores) or prompt text.
 */
import type { RenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'

export type SkeletonRole = 'cover' | 'content' | 'closing'

export interface SkeletonBox {
  x: number
  y: number
  w: number
  h: number
}

export interface SkeletonTitle extends SkeletonBox {
  sizePt: number
  color?: string
}

export interface SkeletonBrandImage extends SkeletonBox {
  /** How many pages of the role carry the image at this box */
  pages: number
}

export interface SkeletonAccent extends SkeletonBox {
  fill?: string
}

export interface RoleSkeleton {
  background?: string
  title?: SkeletonTitle
  brandImage?: SkeletonBrandImage
  accents?: SkeletonAccent[]
}

export interface LayoutSkeleton {
  canvas: { w: number; h: number }
  cover?: RoleSkeleton
  content?: RoleSkeleton
  closing?: RoleSkeleton
}

/** Page role from position: first page covers, last page closes (3+ pages), rest are content. */
export function skeletonRole(pageIndex: number, totalPages: number): SkeletonRole {
  if (pageIndex === 0) return 'cover'
  if (totalPages >= 3 && pageIndex === totalPages - 1) return 'closing'
  return 'content'
}

/** Two boxes are the same chrome element when every edge lands within this tolerance (px) */
const BOX_TOLERANCE_PX = 12
/** A recurring image on at least this many role pages reads as a brand slot (logo) */
const BRAND_MIN_PAGES = 2
/** A recurring filled shape on at least this fraction of role pages reads as template chrome */
const ACCENT_MIN_FRACTION = 0.6
/** Accent shapes larger than this fraction of the canvas are background blocks, not chrome */
const ACCENT_MAX_AREA_FRACTION = 0.9
/** Max accent shapes kept per role (most frequent first) */
const ACCENT_MAX = 2

interface FlatElement extends SkeletonBox {
  kind: 'title' | 'picture' | 'shape'
  sizePt?: number
  color?: string
  fill?: string
}

function flatElements(slide: RenderSlide): FlatElement[] {
  const out: FlatElement[] = []
  const walk = (nodes: RenderNode[], ox: number, oy: number): void => {
    for (const n of nodes) {
      const x = ox + n.box.x
      const y = oy + n.box.y
      const w = n.box.w
      const h = n.box.h
      if (n.type === 'group') {
        walk(n.children, x, y)
        continue
      }
      if (n.type === 'shape' || n.type === 'text') {
        const s = n as ShapeRenderNode
        let maxPx = 0
        let text = ''
        const weight = new Map<string, number>()
        for (const line of s.text?.lines ?? []) {
          for (const r of line.runs) {
            if (r.fontSizePx > maxPx) maxPx = r.fontSizePx
            text += r.text
            if (!r.isBullet) {
              const c = hex6(r.color)
              if (c) weight.set(c, (weight.get(c) ?? 0) + r.text.length)
            }
          }
        }
        if (maxPx > 0 && text.trim()) {
          let best: string | undefined
          let bestW = 0
          for (const [c, cw] of weight) {
            if (cw > bestW) {
              best = c
              bestW = cw
            }
          }
          out.push({
            kind: 'title',
            x,
            y,
            w,
            h,
            sizePt: Math.round((maxPx * 72) / 96),
            ...(best ? { color: best } : {}),
          })
        } else if (s.fill.kind === 'solid') {
          const fill = hex6(s.fill.color)
          out.push({ kind: 'shape', x, y, w, h, ...(fill ? { fill } : {}) })
        }
        continue
      }
      if (n.type === 'picture') out.push({ kind: 'picture', x, y, w, h })
    }
  }
  walk(slide.nodes, 0, 0)
  return out
}

function hex6(color: string | undefined): string | undefined {
  if (!color) return undefined
  const m = /^#([0-9a-fA-F]{6})/.exec(color.trim())
  return m ? `#${m[1].toUpperCase()}` : undefined
}

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b)
  return sorted[Math.floor(sorted.length / 2)]!
}

function sameBox(a: SkeletonBox, b: SkeletonBox): boolean {
  return (
    Math.abs(a.x - b.x) <= BOX_TOLERANCE_PX &&
    Math.abs(a.y - b.y) <= BOX_TOLERANCE_PX &&
    Math.abs(a.w - b.w) <= BOX_TOLERANCE_PX &&
    Math.abs(a.h - b.h) <= BOX_TOLERANCE_PX
  )
}

function medianBox(boxes: SkeletonBox[]): SkeletonBox {
  return {
    x: Math.round(median(boxes.map((b) => b.x))),
    y: Math.round(median(boxes.map((b) => b.y))),
    w: Math.round(median(boxes.map((b) => b.w))),
    h: Math.round(median(boxes.map((b) => b.h))),
  }
}

function picsOf(els: FlatElement[]): FlatElement[] {
  return els.filter((e) => e.kind === 'picture')
}

function accentShapes(els: FlatElement[], canvasArea: number): FlatElement[] {
  return els.filter((e) => e.kind === 'shape' && e.w * e.h < ACCENT_MAX_AREA_FRACTION * canvasArea)
}

function skeletonForRole(slides: RenderSlide[]): RoleSkeleton | undefined {
  if (slides.length === 0) return undefined
  const perSlide = slides.map(flatElements)

  // Background: the majority solid slide background
  const backgrounds = slides
    .map((s) => (s.background.kind === 'solid' ? hex6(s.background.color) : undefined))
    .filter((c): c is string => Boolean(c))
  const bgCount = new Map<string, number>()
  for (const c of backgrounds) bgCount.set(c, (bgCount.get(c) ?? 0) + 1)
  const background = [...bgCount.entries()].sort((a, b) => b[1] - a[1])[0]?.[0]

  // Title: the largest text on each page, median box/size across the role
  const titles = perSlide
    .map((els) => els.filter((e) => e.kind === 'title').sort((a, b) => b.sizePt! - a.sizePt!)[0])
    .filter((e): e is FlatElement & { sizePt: number } => Boolean(e))
  const title =
    titles.length > 0
      ? {
          ...medianBox(titles),
          sizePt: Math.round(median(titles.map((t) => t.sizePt))),
          ...(majorityColor(titles) ? { color: majorityColor(titles) } : {}),
        }
      : undefined

  // Brand image: a picture box recurring on 2+ role pages (the logo slot)
  const pictureClusters: { boxes: SkeletonBox[]; pages: number }[] = []
  for (const els of perSlide) {
    for (const p of picsOf(els)) {
      const cluster = pictureClusters.find((c) => sameBox(c.boxes[0]!, p))
      if (cluster) {
        cluster.pages++
        cluster.boxes.push({ x: p.x, y: p.y, w: p.w, h: p.h })
      } else {
        pictureClusters.push({ boxes: [{ x: p.x, y: p.y, w: p.w, h: p.h }], pages: 1 })
      }
    }
  }
  const brandCluster = pictureClusters
    .filter((c) => c.pages >= BRAND_MIN_PAGES)
    .sort((a, b) => b.pages - a.pages)[0]
  const brandImage = brandCluster
    ? { ...medianBox(brandCluster.boxes), pages: brandCluster.pages }
    : undefined

  // Accents: filled shapes recurring on most role pages, not background-sized
  const canvasArea = (slides[0]!.widthPx || 1) * (slides[0]!.heightPx || 1)
  const shapeClusters: { boxes: SkeletonBox[]; fill?: string; pages: number }[] = []
  for (const els of perSlide) {
    for (const s of accentShapes(els, canvasArea)) {
      const cluster = shapeClusters.find((c) => sameBox(c.boxes[0]!, s) && c.fill === s.fill)
      if (cluster) {
        cluster.pages++
        cluster.boxes.push({ x: s.x, y: s.y, w: s.w, h: s.h })
      } else {
        shapeClusters.push({
          boxes: [{ x: s.x, y: s.y, w: s.w, h: s.h }],
          ...(s.fill ? { fill: s.fill } : {}),
          pages: 1,
        })
      }
    }
  }
  const accents = shapeClusters
    .filter((c) => c.pages >= Math.max(BRAND_MIN_PAGES, ACCENT_MIN_FRACTION * slides.length))
    .sort((a, b) => b.pages - a.pages)
    .slice(0, ACCENT_MAX)
    .map((c) => ({ ...medianBox(c.boxes), ...(c.fill ? { fill: c.fill } : {}) }))

  if (!background && !title && !brandImage && accents.length === 0) return undefined
  return {
    ...(background ? { background } : {}),
    ...(title ? { title } : {}),
    ...(brandImage ? { brandImage } : {}),
    ...(accents.length > 0 ? { accents } : {}),
  }
}

function majorityColor(els: FlatElement[]): string | undefined {
  const count = new Map<string, number>()
  for (const e of els) {
    if (!e.color) continue
    count.set(e.color, (count.get(e.color) ?? 0) + 1)
  }
  const best = [...count.entries()].sort((a, b) => b[1] - a[1])[0]
  return best && best[1] >= Math.ceil(els.length / 2) ? best[0] : undefined
}

/**
 * Chrome skeleton of the deck, or null when nothing recurs (a one-page deck or
 * fully free-form pages carry no reusable structure).
 */
export function extractLayoutSkeleton(slides: RenderSlide[]): LayoutSkeleton | null {
  const visible = slides.filter((s) => !s.hidden && s.nodes.length > 0)
  if (visible.length === 0) return null
  const byRole = new Map<SkeletonRole, RenderSlide[]>()
  visible.forEach((slide, i) => {
    const role = skeletonRole(i, visible.length)
    byRole.set(role, [...(byRole.get(role) ?? []), slide])
  })
  const cover = skeletonForRole(byRole.get('cover') ?? [])
  const content = skeletonForRole(byRole.get('content') ?? [])
  const closing = skeletonForRole(byRole.get('closing') ?? [])
  if (!cover && !content && !closing) return null
  return {
    canvas: { w: visible[0]!.widthPx, h: visible[0]!.heightPx },
    ...(cover ? { cover } : {}),
    ...(content ? { content } : {}),
    ...(closing ? { closing } : {}),
  }
}

/** One prompt line per pinned feature; empty when the role carries no skeleton data. */
export function formatSkeletonForPrompt(skeleton: LayoutSkeleton, role: SkeletonRole): string {
  const r = skeleton[role]
  if (!r) return ''
  const lines: string[] = []
  if (r.background) lines.push(`- background: ${r.background}`)
  if (r.title)
    lines.push(
      `- title box: x=${r.title.x} y=${r.title.y} w=${r.title.w} h=${r.title.h}, font ${r.title.sizePt}pt${r.title.color ? `, color ${r.title.color}` : ''} — put this page's title text in exactly this box`,
    )
  if (r.brandImage)
    lines.push(
      `- brand image slot: x=${r.brandImage.x} y=${r.brandImage.y} w=${r.brandImage.w} h=${r.brandImage.h} — when this page carries a brand/identifying image, place it exactly here (the same spot on every page)`,
    )
  for (const [i, a] of (r.accents ?? []).entries()) {
    lines.push(
      `- template accent shape ${i + 1}: x=${a.x} y=${a.y} w=${a.w} h=${a.h}${a.fill ? `, fill ${a.fill}` : ''} — reproduce on every page`,
    )
  }
  if (lines.length === 0) return ''
  return (
    `Template chrome for ${role} pages (extracted from the saved template — keep exactly, on every ${role} page):\n` +
    lines.join('\n') +
    '\nContent layout below the chrome is yours; adapt the chosen layout variant so the pinned chrome still fits.'
  )
}
