/**
 * PDF-export link overlays: slides are rasterized to page PNGs, so element and
 * run hyperlinks must be re-emitted as clickable rectangles for the print
 * window to turn into PDF link annotations (transparent <a> boxes over each
 * page image; Chromium's printToPDF converts them to /URI actions and
 * in-document destinations).
 *
 * Geometry mirrors the slideshow's hitLink (SlideShowView): unrotated node
 * boxes, group children clipped to the group box, run rects from the laid-out
 * glyph geometry. What the show would follow on click is what the PDF links.
 */
import type { RenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import type { ExportPdfLink, LinkTargetOp } from '../shared/ipc'
import { safeExternalUrl } from '@genoffice/electron-utils/safe-external-url'
import { DECK_LINK_PROTOCOLS } from '../shared/run-link'

interface TargetRect {
  x: number
  y: number
  w: number
  h: number
  target: LinkTargetOp
}

interface Bounds {
  x: number
  y: number
  w: number
  h: number
}

/** Clip a rect to bounds; null when nothing visible remains. */
function clipRect(r: Bounds, b: Bounds): Bounds | null {
  const x1 = Math.max(r.x, b.x)
  const y1 = Math.max(r.y, b.y)
  const x2 = Math.min(r.x + r.w, b.x + b.w)
  const y2 = Math.min(r.y + r.h, b.y + b.h)
  if (!(x2 - x1 > 0) || !(y2 - y1 > 0)) return null
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 }
}

/**
 * All linked rectangles of one slide in px, bottom-to-top z-order (the DOM
 * overlay stacks later anchors on top, matching the show's topmost-wins hit
 * test). Element rects come before the element's own run rects so a run link
 * inside a linked shape stays clickable (more specific beats whole-element).
 */
export function collectSlideLinkRects(
  slide: RenderSlide,
  links: ReadonlyMap<string, LinkTargetOp>,
  runLinks: ReadonlyMap<string, LinkTargetOp>,
): TargetRect[] {
  const out: TargetRect[] = []
  // Clip to the innermost group box like hitLink (a point outside the group
  // never reaches its children), then to the page.
  const walk = (nodes: readonly RenderNode[], dx: number, dy: number, clip: Bounds): void => {
    for (const n of nodes) {
      const { box } = n
      const push = (r: Bounds, target: LinkTargetOp): void => {
        const c = clipRect(r, clip)
        if (c) out.push({ ...c, target })
      }
      const target = links.get(n.sourceId)
      if (target) push({ x: dx + box.x, y: dy + box.y, w: box.w, h: box.h }, target)
      const text = (n as ShapeRenderNode).text
      // hitLink rejects a point outside the node box before looking at runs, so
      // overflowing glyphs are not followable in the show and must not link here.
      const nodeClip =
        text && runLinks.size
          ? clipRect({ x: dx + box.x, y: dy + box.y, w: box.w, h: box.h }, clip)
          : null
      if (text && nodeClip) {
        let para = -1
        for (const ln of text.lines) {
          if (ln.paraStart !== false) para++
          for (const r of ln.runs) {
            if (r.srcRunIdx == null) continue
            const runTarget = runLinks.get(`${n.sourceId}:${para}:${r.srcRunIdx}`)
            if (!runTarget) continue
            const c = clipRect(
              {
                x: dx + box.x + text.insets.l + r.x,
                y: dy + box.y + text.insets.t + ln.top,
                w: r.widthPx,
                h: ln.height,
              },
              nodeClip,
            )
            if (c) out.push({ ...c, target: runTarget })
          }
        }
      }
      if (n.type === 'group') {
        const inner = clipRect({ x: dx + box.x, y: dy + box.y, w: box.w, h: box.h }, clip)
        if (inner) walk(n.children, dx + box.x, dy + box.y, inner)
      }
    }
  }
  walk(slide.nodes, 0, 0, { x: 0, y: 0, w: slide.widthPx, h: slide.heightPx })
  return out
}

/**
 * Resolve a link target to the export HTML href: URLs pass through, slide
 * jumps and navigation actions become in-document page anchors ("#pgN", the
 * id scheme of main/pdf-export.ts). Null = the link has no PDF meaning
 * (jump to a hidden page, end-show, last-slide-viewed) and is dropped —
 * matching the show's skip semantics for hidden pages.
 */
export function exportLinkHref(
  target: LinkTargetOp,
  page: number,
  pageOfModelIndex: ReadonlyMap<number, number>,
  pageCount: number,
): string | null {
  const anchor = (p: number): string | null => (p >= 0 && p < pageCount ? `#pg${p + 1}` : null)
  if (target.kind === 'url')
    return safeExternalUrl(target.url, { allowedProtocols: DECK_LINK_PROTOCOLS })
  if (target.kind === 'slide') {
    const p = pageOfModelIndex.get(target.slideIndex)
    return p == null ? null : anchor(p)
  }
  switch (target.action) {
    case 'nextslide':
      return anchor(page + 1)
    case 'previousslide':
      return anchor(page - 1)
    case 'firstslide':
      return anchor(0)
    case 'lastslide':
      return anchor(pageCount - 1)
    default:
      return null
  }
}

/**
 * Per-page clickable overlays for the PDF export payload: rects as fractions
 * of the page box (the print window scales pages to physical inches, so px
 * would be meaningless there), hrefs fully resolved.
 */
export function collectExportPdfLinks(
  visible: readonly RenderSlide[],
  linkLists: ReadonlyArray<ReadonlyArray<{ sourceId: string; target: LinkTargetOp }>>,
  runLinkLists: ReadonlyArray<
    ReadonlyArray<{ sourceId: string; paraIndex: number; runIndex: number; target: LinkTargetOp }>
  >,
  pageOfModelIndex: ReadonlyMap<number, number>,
): ExportPdfLink[][] {
  return visible.map((slide, page) => {
    const links = new Map((linkLists[page] ?? []).map((l) => [l.sourceId, l.target] as const))
    const runLinks = new Map(
      (runLinkLists[page] ?? []).map(
        (l) => [`${l.sourceId}:${l.paraIndex}:${l.runIndex}`, l.target] as const,
      ),
    )
    if (!links.size && !runLinks.size) return []
    const out: ExportPdfLink[] = []
    for (const r of collectSlideLinkRects(slide, links, runLinks)) {
      const href = exportLinkHref(r.target, page, pageOfModelIndex, visible.length)
      if (!href) continue
      if (!(slide.widthPx > 0) || !(slide.heightPx > 0)) continue
      out.push({
        x: r.x / slide.widthPx,
        y: r.y / slide.heightPx,
        w: r.w / slide.widthPx,
        h: r.h / slide.heightPx,
        href,
      })
    }
    return out
  })
}
